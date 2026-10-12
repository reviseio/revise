#!/usr/bin/env python3
"""Summarize benchmark results as Markdown: for each results directory (one
agent and model), pass rates and medians of wall time, turns, cost, and
tokens for each arm, by task category and by task, with the ratios
baseline / revise.

  python3 summarize.py out/claude out/codex out/hermes > summary.md
  python3 summarize.py out/* >> "$GITHUB_STEP_SUMMARY"
  python3 summarize.py out/* --export results/2026-10-09-abc123 > results/2026-10-09-abc123/summary.md

--export copies each directory's meta.json and its records (records.json)
into DEST/<agent>/, without the agents' work directories, for publishing.

Each directory is a run.py --out. Its summary.json is used when the run has
finished, else the per-run record.json files written so far.
"""
import argparse, json, re, statistics
from pathlib import Path

CATEGORIES = {
    "Editing": ["redline3", "bulk-rename", "long-doc-fix", "table-row", "heading-fix"],
    "Review": ["compare", "resolve-tracked", "comment-review"],
    "Create & convert": ["to-pdf", "generate-memo"],
}
TASK_ORDER = [t for tasks in CATEGORIES.values() for t in tasks]


def all_records(out):
    out = Path(out)
    summary = out / "summary.json"
    if summary.exists():
        return json.loads(summary.read_text())
    return [json.loads(p.read_text()) for p in sorted(out.glob("*/*/*/record.json"))]


def records(out):
    """Runs both arms have finished: for each task, the first n runs of each
    arm, where n is the smaller arm's count (a run still in progress would
    otherwise compare unequal sets)."""
    rs = all_records(out)
    paired = []
    for task in sorted({r["task"] for r in rs}):
        arms = {arm: sorted((r for r in rs if r["task"] == task and r["arm"] == arm), key=lambda r: r["run"])
                for arm in ("revise", "baseline")}
        n = min(len(v) for v in arms.values())
        paired += arms["revise"][:n] + arms["baseline"][:n]
    return paired


def median(values):
    values = [v for v in values if v is not None]
    return statistics.median(values) if values else None


def tokens(r):
    return (r.get("usage") or {}).get("total_tokens")


def pair(rs, bs, get, fmt):
    a, b = median(get(r) for r in rs), median(get(r) for r in bs)
    if a is None or b is None:
        return "–"
    ratio = f" ({b / a:.1f}x)" if a else ""
    return f"{fmt(a)} / {fmt(b)}{ratio}"


def row(label, rs, bs):
    return "| " + " | ".join([
        label,
        f"{sum(r['passed'] for r in rs)}/{len(rs)}",
        f"{sum(r['passed'] for r in bs)}/{len(bs)}",
        pair(rs, bs, lambda r: r.get("wall_s"), lambda v: f"{v:.0f}s"),
        pair(rs, bs, lambda r: r.get("turns"), lambda v: f"{v:.0f}"),
        pair(rs, bs, lambda r: r.get("cost_usd") or None, lambda v: f"${v:.2f}"),
        pair(rs, bs, lambda r: tokens(r) or None, lambda v: f"{v / 1000:.0f}k"),
    ]) + " |"


HEADER = ("| {0} | revise pass | baseline pass | wall, median (revise / baseline) | turns | cost | tokens |\n"
          "|---|---|---|---|---|---|---|")


def split(rs, tasks=None):
    mine = [r for r in rs if tasks is None or r["task"] in tasks]
    return [r for r in mine if r["arm"] == "revise"], [r for r in mine if r["arm"] == "baseline"]


def label_of(out, meta):
    agent = meta.get("agent") or Path(out).name
    model = meta.get("model")
    provider = f" via {meta['provider']}" if meta.get("provider") else ""
    return f"{agent} · {model}{provider}" if model else agent


def section(out, rs, meta):
    lines = [f"## {label_of(out, meta)}", ""]
    facts = [
        ("Agent", meta.get("agent_version")),
        ("revise", meta.get("revise_version")),
        ("revise package", (meta.get("revise_package") or {}).get("spec")),
        ("Billing", meta.get("billing")),
        ("Runs per task per arm", meta.get("runs")),
        ("Started", meta.get("started_at")),
        ("Harness commit", ((meta.get("harness") or {}).get("commit") or "")[:12] or None),
        ("Host", (meta.get("host") or {}).get("platform")),
        ("LibreOffice", (meta.get("host") or {}).get("libreoffice")),
        ("pandoc", (meta.get("host") or {}).get("pandoc")),
        ("python-docx", (meta.get("host") or {}).get("python_docx")),
    ]
    bullets = [f"- {k}: {v}" for k, v in facts if v is not None]
    lines += bullets + [""] if bullets else []
    lines += [HEADER.format("Category")]
    for category, tasks in CATEGORIES.items():
        r, b = split(rs, tasks)
        if r and b:
            lines.append(row(category, r, b))
    r, b = split(rs)
    lines.append(row("**All tasks**", r, b))
    lines += ["", "<details><summary>By task</summary>", "", HEADER.format("Task")]
    for task in sorted({x["task"] for x in rs}, key=lambda t: (TASK_ORDER.index(t) if t in TASK_ORDER else 99, t)):
        r, b = split(rs, [task])
        lines.append(row(f"`{task}`", r, b))
    lines += ["", "</details>", ""]
    failed = [x for x in rs if not x["passed"]]
    if failed:
        lines += ["<details><summary>Failed checks</summary>", ""]
        for x in sorted(failed, key=lambda x: (x["arm"], x["task"], x["run"])):
            misses = ", ".join(k for k, v in x["checks"].items() if not v)
            lines.append(f"- {x['arm']} `{x['task']}` #{x['run']}: {misses}")
        lines += ["", "</details>", ""]
    return lines


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("dirs", nargs="+")
    ap.add_argument("--title", default="revise benchmark")
    ap.add_argument("--export", metavar="DEST", help="also write meta.json and records.json per directory under DEST")
    args = ap.parse_args()
    sections, overview = [], []
    for out in args.dirs:
        rs = records(out)
        if not rs:
            continue
        meta_path = Path(out) / "meta.json"
        meta = json.loads(meta_path.read_text()) if meta_path.exists() else {}
        sections += section(out, rs, meta)
        if args.export:
            name = re.sub(r"[^\w.-]+", "-", label_of(out, meta).replace(" · ", "-")).strip("-")
            dest = Path(args.export) / name
            dest.mkdir(parents=True, exist_ok=True)
            (dest / "meta.json").write_text(json.dumps(meta, indent=2) + "\n")
            (dest / "records.json").write_text(json.dumps(all_records(out), indent=2) + "\n")
        r, b = split(rs)
        overview.append(row(label_of(out, meta), r, b))
    print(f"# {args.title}\n")
    print("Each cell compares the revise arm with the baseline arm on the same tasks and runs. "
          "Pass = every scripted check passed, fidelity checks included. Ratios are baseline / revise; "
          "cost and tokens appear only where the agent reports them.\n")
    if overview:
        print(HEADER.format("Agent"))
        print("\n".join(overview) + "\n")
    print("\n".join(sections))


if __name__ == "__main__":
    main()
