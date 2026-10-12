"""Tasks and graders for the revise benchmark.

Everything below the FIXTURES table is the grading logic as published; run.py
imports it. Graders read only the agent's output file and the original
fixtures, using regular expressions over the OOXML parts (no third-party
libraries), so a result can be re-graded anywhere with `run.py --grade-only`.
"""
import html, re, zipfile
from pathlib import Path

FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures"
# Task name -> fixture path. set_fixtures_dir() repoints these in place, so the
# graders below always read the active fixture set.
FIXTURE_FILES = ["contract.docx", "contract-v2.docx", "contract-tracked.docx", "nda.docx", "handbook.docx"]
FIXTURES = {name: FIXTURES_DIR / name for name in FIXTURE_FILES}


def set_fixtures_dir(directory):
    for name in FIXTURE_FILES:
        FIXTURES[name] = Path(directory).resolve() / name

# ── OOXML views ──────────────────────────────────────────────────────────────

P = re.compile(r"<w:p[ >].*?</w:p>", re.S)
REVIEW_PARTS = {
    "word/document.xml", "word/comments.xml", "word/commentsExtended.xml",
    "word/commentsIds.xml", "word/commentsExtensible.xml", "word/people.xml",
    "[Content_Types].xml", "word/_rels/document.xml.rels", "word/settings.xml",
}


def read_part(path, name):
    try:
        with zipfile.ZipFile(path) as z:
            return z.read(name).decode("utf8")
    except (KeyError, FileNotFoundError, zipfile.BadZipFile):
        return None


def drop_rows(xml, accepted):
    tag = "del" if accepted else "ins"
    return re.sub(
        r"<w:tr\b(?:(?!</w:tr>).)*?<w:trPr>(?:(?!</w:trPr>).)*<w:" + tag + r" (?:(?!</w:tr>).)*</w:tr>",
        "", xml, flags=re.S)


def para_text(p, accepted=True):
    p = re.sub(r"<w:del\b.*?</w:del>", "", p, flags=re.S) if accepted else re.sub(r"<w:ins\b.*?</w:ins>", "", p, flags=re.S)
    p = re.sub(r"<w:moveFrom\b.*?</w:moveFrom>", "", p, flags=re.S) if accepted else re.sub(r"<w:moveTo\b.*?</w:moveTo>", "", p, flags=re.S)
    p = p.replace("<w:delText", "<w:t").replace("</w:delText>", "</w:t>")
    return html.unescape("".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p)))


def paragraphs(xml, accepted=True):
    return [para_text(p, accepted) for p in P.findall(drop_rows(xml, accepted))]


def flat(texts):
    return re.sub(r"\s+", " ", " ".join(texts)).strip()


def doc_text(path, accepted=True):
    xml = read_part(path, "word/document.xml")
    return flat(paragraphs(xml, accepted)) if xml else None


def revision_authors(xml):
    return re.findall(r"<w:(?:ins|del|moveFrom|moveTo) [^>]*w:author=\"([^\"]*)\"", xml)


def comments(path):
    xml = read_part(path, "word/comments.xml") or ""
    return re.findall(r"<w:comment [^>]*w:author=\"([^\"]*)\"", xml)


def changed_parts(orig, out):
    with zipfile.ZipFile(orig) as zo, zipfile.ZipFile(out) as zn:
        names = set(zn.namelist())
        return sorted(n for n in zo.namelist() if n not in REVIEW_PARTS and (n not in names or zo.read(n) != zn.read(n)))


EMPTY = re.compile(r"<([\w:]+)((?:\s[^<>]*?)?)></\1>")


def canonical(paragraph):
    """<w:t ...></w:t> and <w:t .../> are the same XML; an XML library re-saving
    a file may write either, so only the spelling changes."""
    return EMPTY.sub(r"<\1\2/>", paragraph)


def changed_paragraph_count(orig, out):
    po = [canonical(p) for p in P.findall(read_part(orig, "word/document.xml"))]
    pn = {canonical(p) for p in P.findall(read_part(out, "word/document.xml"))}
    return sum(1 for p in po if p not in pn)


def fidelity_checks(orig, out, max_changed):
    parts = changed_parts(orig, out)
    changed = changed_paragraph_count(orig, out)
    return {
        "other parts byte-identical": not parts,
        f"at most {max_changed} paragraphs changed": changed <= max_changed,
    }, {"changed_parts": parts, "changed_paragraphs": changed}


def tracked_checks(orig, out, author):
    xml = read_part(out, "word/document.xml")
    authors = revision_authors(xml)
    return {
        "has tracked changes": bool(authors),
        f'every revision by "{author}"': bool(authors) and set(authors) == {author},
        "rejecting all restores the original text": doc_text(out, False) == doc_text(orig, True),
    }


# ── tasks ────────────────────────────────────────────────────────────────────

def g_redline3(d):
    o, n = FIXTURES["contract.docx"], d / "contract-redlined.docx"
    if not n.exists():
        return {"output exists": False}, {}
    acc = doc_text(n)
    checks = {
        "9.2 mutual, ninety days": bool(re.search(r"9\.2 Convenience\..{0,200}(either party|each party).{0,120}ninety", acc, re.I)),
        "4.1 99.9% + service credits": bool(re.search(r"4\.1 Availability\..{0,400}99\.9%.{0,600}credit", acc, re.I)),
        "7.2 48 hours + Provider bears cost": bool(re.search(r"7\.2 Security incident\..{0,600}(48|forty-eight).{0,600}Provider (bears|shall bear|will bear)", acc, re.I)),
        "3+ comments by Claude": len([a for a in comments(n) if a == "Claude"]) >= 3,
        **tracked_checks(o, n, "Claude"),
    }
    fid, info = fidelity_checks(o, n, 4)
    return {**checks, **fid}, info


def g_bulk_rename(d):
    o, n = FIXTURES["contract.docx"], d / "contract-renamed.docx"
    if not n.exists():
        return {"output exists": False}, {}
    before, after = doc_text(o), doc_text(n)
    providers = len(re.findall(r"\bProvider\b", before))
    expected = providers + len(re.findall(r"\bNorthstar\b", before))
    paragraphs_with_provider = sum(1 for t in paragraphs(read_part(o, "word/document.xml")) if re.search(r"\bProvider\b", t))
    checks = {
        "no Provider left": not re.search(r"\bProvider\b", after),
        f"Northstar x{expected}": len(re.findall(r"\bNorthstar\b", after)) == expected,
        **tracked_checks(o, n, "Claude"),
    }
    fid, info = fidelity_checks(o, n, paragraphs_with_provider)
    return {**checks, **fid}, info


def g_long_doc_fix(d):
    o, n = FIXTURES["handbook.docx"], d / "handbook-updated.docx"
    if not n.exists():
        return {"output exists": False}, {}
    acc = doc_text(n)
    checks = {
        "accrual is 2 days per month": "2 days per month" in acc and "1.5 days per month" not in acc,
        "no annual total of 18 days left": "18 days of paid leave" not in acc,
        "annual total of 24 days, 3 places": len(re.findall(r"24 days", acc)) >= 3,
        "distractor notice period untouched": "18 days' written notice" in acc,
        **tracked_checks(o, n, "Claude"),
    }
    fid, info = fidelity_checks(o, n, 3)
    return {**checks, **fid}, info


def g_compare(d):
    o, v2, n = FIXTURES["contract.docx"], FIXTURES["contract-v2.docx"], d / "redline.docx"
    if not n.exists():
        return {"output exists": False}, {}
    authors = revision_authors(read_part(n, "word/document.xml"))
    return {
        "has tracked changes": bool(authors),
        'every revision by "Northstar Legal"': bool(authors) and set(authors) == {"Northstar Legal"},
        "accepting all gives contract-v2": doc_text(n, True) == doc_text(v2, True),
        "rejecting all gives contract": doc_text(n, False) == doc_text(o, True),
    }, {}


def g_to_pdf(d):
    n = d / "contract.pdf"
    if not n.exists():
        return {"output exists": False}, {}
    data = n.read_bytes()
    pages = len(re.findall(rb"/Type\s*/Page[^s]", data))
    return {
        "is a PDF": data.startswith(b"%PDF"),
        "4-6 pages (source is 5)": 4 <= pages <= 6,
    }, {"pages": pages}


def table_rows(xml):
    rows = []
    for tr in re.findall(r"<w:tr\b.*?</w:tr>", xml, re.S):
        cells = [flat([para_text(p) for p in P.findall(tc)]) for tc in re.findall(r"<w:tc>.*?</w:tc>|<w:tc .*?</w:tc>", tr, re.S)]
        rows.append(cells)
    return rows


def g_table_row(d):
    o, n = FIXTURES["contract.docx"], d / "contract-support.docx"
    if not n.exists():
        return {"output exists": False}, {}
    rows = table_rows(drop_rows(read_part(n, "word/document.xml"), True))
    texts = [" | ".join(r) for r in rows]
    total = next((i for i, t in enumerate(texts) if t.startswith("Year-one total")), None)
    support = next((i for i, t in enumerate(texts) if t.startswith("Premium support")), None)
    checks = {
        "support row has the three cells": support is not None and rows[support][:3] == ["Premium support", "Annual flat fee", "$12,000"],
        "support row directly above the total": support is not None and total is not None and support == total - 1,
        "total is $174,000": total is not None and "$174,000" in texts[total],
        **tracked_checks(o, n, "Claude"),
    }
    fid, info = fidelity_checks(o, n, 6)
    return {**checks, **fid}, info


def g_generate_memo(d):
    n = d / "memo.docx"
    if not n.exists():
        return {"output exists": False}, {}
    xml = read_part(n, "word/document.xml") or ""
    styles = read_part(n, "word/styles.xml") or ""
    heading_ids = set(re.findall(r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>\s*<w:name w:val="(?:heading 1|Heading 1)"', styles)) | {"Heading1"}
    headings = [para_text(p) for p in P.findall(xml) if any(f'w:pStyle w:val="{h}"' in p for h in heading_ids)]
    # Numbering on the paragraph, or on its style (Word's "List Bullet").
    numbered_styles = {
        m.group(1)
        for m in re.finditer(r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>(?:(?!</w:style>).)*<w:numPr>', styles, re.S)
    }
    bullets = [
        para_text(p) for p in P.findall(xml)
        if "<w:numPr>" in p
        or any(f'w:pStyle w:val="{sid}"' in p for sid in numbered_styles)
    ]
    tables = re.findall(r"<w:tbl>.*?</w:tbl>", xml, re.S)
    rows = table_rows(tables[0]) if tables else []
    footers = " ".join(read_part(n, f"word/footer{i}.xml") or "" for i in range(1, 4)) + (read_part(n, "word/footer.xml") or "")
    sz = re.search(r"<w:pgSz [^>]*>", xml)
    mar = re.search(r"<w:pgMar [^>]*>", xml)
    return {
        "title Q4 Planning": "Q4 Planning" in flat(paragraphs(xml)),
        "Heading 1 Goals and Budget": {"Goals", "Budget"} <= {h.strip() for h in headings},
        "3 bulleted goals": len(bullets) >= 3 and all(any(g.lower() in b.lower() for b in bullets) for g in ["Ship the CLI", "Publish the benchmark", "Sign two design partners"]),
        "table: header + 2 rows": len(rows) >= 3 and [c.strip() for c in rows[0][:3]] == ["Item", "Owner", "Amount"],
        "footer has PAGE and NUMPAGES fields": bool(re.search(r"\bPAGE\b", footers)) and "NUMPAGES" in footers,
        "US Letter": bool(sz) and 'w:w="12240"' in sz.group(0) and 'w:h="15840"' in sz.group(0),
        "1-inch margins": bool(mar) and all(f'w:{k}="1440"' in mar.group(0) for k in ["top", "right", "bottom", "left"]),
    }, {}


def section_of(texts):
    sec, out = None, []
    for t in texts:
        m = re.match(r"(\d+)\. ", t)
        if m and len(t) < 60:
            sec = int(m.group(1))
        out.append(sec)
    return out


def g_resolve_tracked(d):
    o, n = FIXTURES["contract-tracked.docx"], d / "contract-resolved.docx"
    if not n.exists():
        return {"output exists": False}, {}
    oxml = read_part(o, "word/document.xml")
    paras = P.findall(drop_rows(oxml, True))
    secs = section_of([para_text(p, True) for p in paras])
    expected = flat(para_text(p, s != 8) for p, s in zip(paras, secs))
    nxml = read_part(n, "word/document.xml")
    return {
        "no tracked changes left": not re.search(r"<w:(ins|del|moveFrom|moveTo) ", nxml),
        "text = accepted except section 8 rejected": flat(paragraphs(nxml)) == expected,
    }, {}


def g_comment_review(d):
    o, n = FIXTURES["nda.docx"], d / "nda-reviewed.docx"
    if not n.exists():
        return {"output exists": False}, {}
    xml = read_part(n, "word/document.xml")
    anchored = 0
    for cid in set(re.findall(r'<w:commentRangeStart w:id="(\d+)"', xml)):
        span = re.search(r'<w:commentRangeStart w:id="' + cid + r'"/>(.*?)<w:commentRangeEnd w:id="' + cid + r'"/>', xml, re.S)
        if span and flat([para_text(span.group(1))]):
            anchored += 1
    authors = comments(n)
    fid, info = fidelity_checks(o, n, 999)
    return {
        "3+ comments by Claude": len([a for a in authors if a == "Claude"]) >= 3,
        "every comment anchored to text": anchored >= len(authors) >= 3,
        "no tracked changes": not revision_authors(xml),
        "text unchanged": doc_text(n) == doc_text(o),
        "other parts byte-identical": fid["other parts byte-identical"],
    }, info


def g_heading_fix(d):
    o, n = FIXTURES["handbook.docx"], d / "handbook-fixed.docx"
    if not n.exists():
        return {"output exists": False}, {}
    xml = read_part(n, "word/document.xml")
    target = [p for p in P.findall(xml) if para_text(p).strip().startswith("12. ")]
    checks = {
        "section 12 heading uses Heading1": len(target) == 1 and 'w:pStyle w:val="Heading1"' in target[0],
        "heading text unchanged": len(target) == 1 and para_text(target[0]).strip() == "12. Overtime",
        "no tracked changes": not revision_authors(xml),
        "text unchanged": doc_text(n) == doc_text(o),
    }
    fid, info = fidelity_checks(o, n, 1)
    return {**checks, **fid}, info


TASKS = {
    "redline3": (["contract.docx"], g_redline3, """contract.docx is a software services agreement. We represent the Customer (Cedar Vale). Prepare a redline for the other side:

1. Section 9.2 (Convenience): make the convenience termination right mutual, with ninety days' written notice.
2. Section 4.1 (Availability): raise the availability target to 99.9% and say that missed targets earn service credits.
3. Section 7.2 (Security incident): Provider must notify within 48 hours of becoming aware of an incident, and Provider bears the costs when it caused the incident.

Every change must be a tracked change (Word redline) authored by "Claude", and each of the three sections should get a short Word comment explaining the change. Do not change anything else in the document — formatting, numbering, headers, and all other text must stay exactly as they are. Save the result as contract-redlined.docx in this directory."""),
    "bulk-rename": (["contract.docx"], g_bulk_rename, """In contract.docx, rename the provider: every occurrence of the defined term "Provider" (including "Provider's") must become "Northstar". Make the edits as Word tracked changes authored by "Claude" and change nothing else. Save the result as contract-renamed.docx in this directory."""),
    "long-doc-fix": (["handbook.docx"], g_long_doc_fix, """handbook.docx is our employee handbook. The leave policy is changing: leave now accrues at 2 days per month (section 37.4), so the annual paid-leave total becomes 24 days. Update the accrual rate and every statement of the annual paid-leave total anywhere in the handbook. Make the edits as Word tracked changes authored by "Claude", change nothing else, and save the result as handbook-updated.docx in this directory."""),
    "compare": (["contract.docx", "contract-v2.docx"], g_compare, """contract-v2.docx is the other side's revision of contract.docx, sent back without tracked changes. Produce redline.docx: contract-v2.docx with every difference from contract.docx shown as a Word tracked change authored by "Northstar Legal", so that accepting all changes gives contract-v2.docx's text and rejecting all changes gives contract.docx's text. Save it in this directory."""),
    "to-pdf": (["contract.docx"], g_to_pdf, """Create contract.pdf from contract.docx in this directory, keeping its page layout, header, and footer."""),
    "table-row": (["contract.docx"], g_table_row, """In contract.docx, section 2's commercial table needs a new line item. Add a row "Premium support | Annual flat fee | $12,000" immediately above the "Year-one total" row, and change the year-one total from $162,000 to $174,000. Make both changes Word tracked changes authored by "Claude", change nothing else, and save the result as contract-support.docx in this directory."""),
    "generate-memo": ([], g_generate_memo, """Create memo.docx in this directory from scratch: a title "Q4 Planning"; a Heading 1 "Goals" followed by a bulleted list of three goals (Ship the CLI, Publish the benchmark, Sign two design partners); a Heading 1 "Budget" followed by a table with a header row (Item, Owner, Amount) and two data rows of your choice; and a footer reading "Page X of Y" with live page-number fields. US Letter paper, 1-inch margins."""),
    "resolve-tracked": (["contract-tracked.docx"], g_resolve_tracked, """contract-tracked.docx contains the counterparty's tracked changes. Accept all of them except those in section 8 (Allocation of risk), which we reject. Save the result as contract-resolved.docx in this directory, with no tracked changes remaining."""),
    "comment-review": (["nda.docx"], g_comment_review, """Review nda.docx for terms that are one-sided against the receiving party. Leave at least three Word comments, authored by "Claude", each anchored to the specific clause text it concerns. Do not change the document's text. Save the result as nda-reviewed.docx in this directory."""),
    "heading-fix": (["handbook.docx"], g_heading_fix, """In handbook.docx, the heading for section 12 was typed as bold body text instead of using the Heading 1 style like every other section heading. Fix it so it uses Heading 1 without changing its text, change nothing else, and save the result as handbook-fixed.docx in this directory. No tracked changes are needed."""),
}
