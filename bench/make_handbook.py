"""Generate fixtures/handbook.docx: a ~115-page employee handbook written as
raw OOXML (not by Revise, so the benchmark fixture carries no bias toward
Revise's own exporter). Deterministic.

  python3 make_handbook.py fixtures/handbook.docx

Planted content the tasks rely on:
  - Section 37.4 states leave accrues at "1.5 days per month".
  - Exactly three paragraphs state the annual total "18 days of paid leave"
    (sections 37.4, 52.1, and 88.2).
  - One distractor, "18 days' written notice" in section 61.3, must not change.
  - Section 12's heading is bold Normal text instead of Heading1.
"""
import random, sys, zipfile
from xml.sax.saxutils import escape

random.seed(20261008)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
TOPICS = ["Attendance", "Remote Work", "Equipment", "Expenses", "Travel", "Security", "Conduct", "Benefits",
          "Training", "Performance", "Compensation", "Overtime", "Holidays", "Health", "Safety", "Privacy",
          "Communications", "Social Media", "Conflicts of Interest", "Gifts", "Records", "Purchasing", "Facilities",
          "Visitors", "Parking", "Recruiting", "Onboarding", "Offboarding", "Grievances", "Discipline"]
VOCAB = ("employee manager department policy request approval written record period company time work "
         "schedule report review process standard practice document notice business member team role "
         "responsibility information system access account office location expectation guideline").split()

def sentence():
    n = random.randint(12, 24)
    words = [random.choice(VOCAB) for _ in range(n)]
    return " ".join(words).capitalize() + "."

def para(text, style=None, bold_runs=None):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    if bold_runs:
        runs = "".join(f'<w:r>{"<w:rPr><w:b/></w:rPr>" if b else ""}<w:t xml:space="preserve">{escape(t)}</w:t></w:r>' for t, b in bold_runs)
    else:
        runs = f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
    return f"<w:p>{ppr}{runs}</w:p>"

def table(rows):
    def cell(t, header):
        rpr = "<w:rPr><w:b/></w:rPr>" if header else ""
        return f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r>{rpr}<w:t>{escape(t)}</w:t></w:r></w:p></w:tc>'
    trs = "".join(f'<w:tr>{"<w:trPr><w:tblHeader/></w:trPr>" if i == 0 else ""}{"".join(cell(c, i == 0) for c in r)}</w:tr>' for i, r in enumerate(rows))
    return f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>{trs}</w:tbl>'

PLANTED = {
    (37, 4): "Paid leave accrues at 1.5 days per month of continuous service, for a total of 18 days of paid leave each calendar year. Accrual begins on the first day of employment.",
    (52, 1): "Part-time staff receive a prorated share of the 18 days of paid leave each calendar year, calculated from their contracted hours.",
    (88, 2): "On departure, unused leave is paid out up to the full annual entitlement of 18 days of paid leave each calendar year.",
    (61, 3): "A request to change a recurring schedule requires 18 days' written notice to the department manager.",
}

body = [para("Employee Handbook", "Title"), para("Effective 1 January 2027. This handbook applies to all employees of Meridian Labs.")]
for s in range(1, 141):
    topic = TOPICS[(s - 1) % len(TOPICS)]
    heading = f"{s}. {topic}" + ("" if s <= len(TOPICS) else f" ({(s - 1) // len(TOPICS) + 1})")
    if s == 12:
        body.append(para(None, None, [(heading, True)]))
    else:
        body.append(para(heading, "Heading1"))
    for sub in range(1, 6):
        text = PLANTED.get((s, sub)) or " ".join(sentence() for _ in range(random.randint(3, 5)))
        body.append(para(None, None, [(f"{s}.{sub} ", True), (text, False)]))
    if s % 10 == 5:
        body.append(table([["Item", "Owner", "Frequency"]] + [[random.choice(VOCAB).capitalize(), random.choice(VOCAB).capitalize(), random.choice(["Monthly", "Quarterly", "Annual"])] for _ in range(4)]))
        body.append(para(""))

sect = '<w:sectPr><w:headerReference w:type="default" r:id="rId8"/><w:footerReference w:type="default" r:id="rId9"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
document = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {W}><w:body>{"".join(body)}{sect}</w:body></w:document>'
styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:count="4"><w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/><w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="Title" w:uiPriority="10" w:qFormat="1"/><w:lsdException w:name="Table Grid" w:uiPriority="59"/></w:latentStyles><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="30"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>'''
header = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr {W}><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:t>MERIDIAN LABS — EMPLOYEE HANDBOOK</w:t></w:r></w:p></w:hdr>'
footer = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr {W}><w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'
ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'
rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
docrels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>'
with zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED) as z:
    for name, data in [("[Content_Types].xml", ct), ("_rels/.rels", rels), ("word/document.xml", document), ("word/_rels/document.xml.rels", docrels), ("word/styles.xml", styles), ("word/header1.xml", header), ("word/footer1.xml", footer)]:
        z.writestr(zipfile.ZipInfo(name, (2026, 10, 8, 0, 0, 0)), data, zipfile.ZIP_DEFLATED)
