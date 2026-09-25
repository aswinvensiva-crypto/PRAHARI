"""Build copy-paste-ready plain-text versions of the portal fields and check SIH character limits.

Usage: python3 source/build_portal_text.py   (run from SIH_Submission/)
"""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "portal_text"
LIMITS = {"title": 100, "description": 50_000, "abstract": 10_000}


def md_to_text(md: str) -> str:
    lines = []
    for line in md.splitlines():
        m = re.match(r"^(#{1,6})\s+(.*)$", line)
        if m:
            level, text = len(m.group(1)), m.group(2)
            text = text.upper() if level <= 2 else text
            if lines and lines[-1] != "":
                lines.append("")
            lines.append(text)
            continue
        line = re.sub(r"^(\s*)- ", r"\1• ", line)
        lines.append(line)
    text = "\n".join(lines)
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip() + "\n"


def main():
    OUT.mkdir(exist_ok=True)
    title_md = (ROOT / "01_Idea_Title.md").read_text()
    title = title_md.split("## FINAL TITLE (paste this into the portal)")[1].strip().splitlines()[0].strip()

    desc_md = (ROOT / "02_Idea_Description.md").read_text()
    # The portal already has a title field, so drop the H1 line from the description body.
    desc_md = desc_md.split("\n", 1)[1].lstrip()
    abstract_md = (ROOT / "03_Abstract.md").read_text().split("\n", 1)[1].lstrip()

    fields = {
        "title": title + "\n",
        "description": md_to_text(desc_md),
        "abstract": md_to_text(abstract_md),
    }
    names = {"title": "1_Title.txt", "description": "2_Description.txt", "abstract": "3_Abstract.txt"}
    report = []
    for key, text in fields.items():
        (OUT / names[key]).write_text(text, encoding="utf-8")
        n = len(text.rstrip("\n"))
        ok = n <= LIMITS[key]
        report.append(f"{names[key]:<18} {n:>6,} / {LIMITS[key]:>6,} chars  {'OK' if ok else 'OVER LIMIT'}")
        assert ok, f"{key} exceeds limit"
    (OUT / "CHARACTER_COUNTS.txt").write_text("\n".join(report) + "\n")
    print("\n".join(report))


if __name__ == "__main__":
    main()
