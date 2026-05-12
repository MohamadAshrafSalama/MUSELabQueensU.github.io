#!/usr/bin/env python3
"""
Fetch publications from Semantic Scholar for MUSE Lab members,
auto-classify topics, and update data/publications.json.

publication.html loads dynamically from the JSON, so no HTML generation needed.
"""

import argparse
import json
import os
import re
import subprocess
import sys
import time
import urllib.request
import urllib.error
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
MEMBERS_FILE = DATA_DIR / "members.json"
PUBS_FILE = DATA_DIR / "publications.json"

S2_API = "https://api.semanticscholar.org/graph/v1"
S2_FIELDS = "title,authors,year,venue,publicationTypes,externalIds,abstract"

TOPIC_KEYWORDS = {
    "oss": [
        "open source", "oss", "contributor", "community", "maintainer",
        "apache", "github", "governance", "sustainability", "newcomer",
        "retention", "turnover", "participation",
    ],
    "edi": [
        "diversity", "inclusion", "equity", "gender", "bias",
        "inclusivity", "underrepresented", "women", "minority",
        "fairness", "belonging",
    ],
    "msr": [
        "mining software", "repository mining", "empirical analysis",
        "commit", "pull request", "msr", "github mining", "temporal pattern",
        "cohort study", "large-scale empirical",
    ],
    "productivity": [
        "productivity", "developer experience", "llm", "copilot",
        "ai assistant", "developer tool", "ide", "workflow",
        "code review", "developer survey", "motivation",
    ],
    "ml4se": [
        "machine learning", "deep learning", "neural", "nlp",
        "language model", "prediction", "classify", "detect",
        "reverse engineer", "user stories from code",
    ],
    "tools": [
        "tool", "dashboard", "detector", "automated", "actionable tool",
        "visualization tool",
    ],
}
DEFAULT_TOPIC = "ese"

VENUE_ACRONYMS = {
    "icse": "ICSE",
    "international conference on software engineering": "ICSE",
    "software engineering in society": "ICSE",
    "ieee/acm 44th": "ICSE",
    "ieee/acm 43rd": "ICSE",
    "ieee/acm 45th": "ICSE",
    "chase": "CHASE",
    "cooperative and human aspects": "CHASE",
    "tosem": "TOSEM",
    "transactions on software engineering and methodology": "TOSEM",
    "tse": "TSE",
    "transactions on software engineering": "TSE",
    "information and software technology": "IST",
    "ieee software": "IEEE-SW",
    "cascon": "CASCON",
    "centre for advanced studies": "CASCON",
    "cscw": "CSCW",
    "human-computer interaction": "CSCW",
    "hum. comput. interact": "CSCW",
    "journal of internet services": "JISA",
    "arxiv": "arXiv",
    "ease": "EASE",
    "evaluation and assessment": "EASE",
}

LAB_MEMBER_NAMES = set()


def load_members():
    with open(MEMBERS_FILE) as f:
        members = json.load(f)
    ids = []
    for m in members:
        LAB_MEMBER_NAMES.add(m["name"].lower())
        for alias in m.get("aliases", []):
            LAB_MEMBER_NAMES.add(alias.lower())
        for sid in m.get("semantic_scholar_ids", []):
            ids.append(sid)
    return members, ids


def fetch_author_papers(author_id):
    url = f"{S2_API}/author/{author_id}/papers?fields={S2_FIELDS}&limit=500"
    req = urllib.request.Request(url, headers={"User-Agent": "MUSELab-PubFetcher/1.0"})
    try:
        resp = urllib.request.urlopen(req, timeout=30)
        data = json.loads(resp.read())
        return data.get("data", [])
    except urllib.error.HTTPError as e:
        print(f"  Warning: HTTP {e.code} fetching author {author_id}", file=sys.stderr)
        return []
    except Exception as e:
        print(f"  Warning: {e} fetching author {author_id}", file=sys.stderr)
        return []


def normalize_title(title):
    return re.sub(r"[^a-z0-9]", "", (title or "").lower())


def deduplicate_papers(all_papers):
    seen = {}
    for p in all_papers:
        key = normalize_title(p.get("title", ""))
        if not key:
            continue
        doi = (p.get("externalIds") or {}).get("DOI", "")
        if doi:
            key = doi.lower()
        existing = seen.get(key)
        if existing is None or (p.get("citationCount") or 0) > (existing.get("citationCount") or 0):
            seen[key] = p
    return list(seen.values())


def classify_topics(title, abstract):
    text = ((title or "") + " " + (abstract or "")).lower()
    matched = []
    for topic, keywords in TOPIC_KEYWORDS.items():
        score = sum(1 for kw in keywords if kw in text)
        if score > 0:
            matched.append((topic, score))
    if not matched:
        return [DEFAULT_TOPIC]
    matched.sort(key=lambda x: -x[1])
    return [t for t, _ in matched]


def extract_venue_acronym(venue):
    if not venue:
        return ""
    venue_lower = venue.lower()
    for pattern, acronym in VENUE_ACRONYMS.items():
        if pattern in venue_lower:
            return acronym
    return ""


def map_pub_type(types_list):
    if not types_list:
        return "conference"
    t = types_list[0].lower() if types_list else ""
    if "journal" in t:
        return "journal"
    if "conference" in t:
        return "conference"
    if "review" in t:
        return "book-chapter"
    return "conference"


def paper_to_record(paper):
    authors = [a.get("name", "") for a in paper.get("authors", [])]
    doi = (paper.get("externalIds") or {}).get("DOI", "")
    venue = paper.get("venue", "")
    return {
        "id": doi or normalize_title(paper.get("title", "")),
        "title": paper.get("title", ""),
        "authors": authors,
        "venue": venue,
        "venue_acronym": extract_venue_acronym(venue),
        "year": paper.get("year"),
        "type": map_pub_type(paper.get("publicationTypes")),
        "topics": classify_topics(paper.get("title"), paper.get("abstract")),
        "doi": doi,
        "links": {
            "doi": f"https://doi.org/{doi}" if doi else "",
            "pdf": "",
            "presentation": "",
        },
        "added_by": "auto",
        "added_on": datetime.now().strftime("%Y-%m-%d"),
    }


def load_existing_pubs():
    if PUBS_FILE.exists():
        with open(PUBS_FILE) as f:
            return json.load(f)
    return []


def find_new_papers(fetched, existing):
    existing_ids = set()
    for p in existing:
        existing_ids.add(p.get("id", ""))
        existing_ids.add(normalize_title(p.get("title", "")))
    new = []
    for p in fetched:
        pid = p.get("id", "")
        ntitle = normalize_title(p.get("title", ""))
        if pid not in existing_ids and ntitle not in existing_ids:
            new.append(p)
    return new


def interactive_classify(papers):
    topics_list = list(TOPIC_KEYWORDS.keys()) + [DEFAULT_TOPIC]
    for i, p in enumerate(papers):
        print(f"\n[{i+1}/{len(papers)}] {p['title']} ({p['year']})")
        print(f"  Auto-classified: {', '.join(p['topics'])}")
        print(f"  Topics: {', '.join(f'{j}={t}' for j, t in enumerate(topics_list))}")
        choice = input("  Enter topic numbers comma-separated (or press Enter to keep): ").strip()
        if choice:
            selected = []
            for c in choice.split(","):
                c = c.strip()
                if c.isdigit() and int(c) < len(topics_list):
                    selected.append(topics_list[int(c)])
            if selected:
                p["topics"] = selected


def open_pr(new_papers):
    branch = f"auto/update-publications-{datetime.now().strftime('%Y-%m')}"
    title = f"Update publications — {datetime.now().strftime('%B %Y')}"

    body_lines = ["## New publications found\n"]
    for p in new_papers:
        topics_str = ", ".join(p["topics"])
        body_lines.append(f"- **{p['title']}** ({p['year']}) [{topics_str}]")
    body_lines.append(f"\n_Auto-generated by `fetch_publications.py` on {datetime.now().strftime('%Y-%m-%d')}_")
    body = "\n".join(body_lines)

    cmds = [
        ["git", "checkout", "-b", branch],
        ["git", "add", "data/publications.json"],
        ["git", "commit", "-m", f"update publications {datetime.now().strftime('%Y-%m')}"],
        ["git", "push", "-u", "origin", branch],
    ]
    for cmd in cmds:
        result = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"  git error: {result.stderr.strip()}", file=sys.stderr)
            return

    result = subprocess.run(
        ["gh", "pr", "create", "--title", title, "--body", body],
        cwd=ROOT, capture_output=True, text=True,
    )
    if result.returncode == 0:
        print(f"PR created: {result.stdout.strip()}")
    else:
        print(f"PR error: {result.stderr.strip()}", file=sys.stderr)


def main():
    parser = argparse.ArgumentParser(description="Fetch and update MUSE Lab publications")
    parser.add_argument("--dry-run", action="store_true", help="Print new papers without modifying files")
    parser.add_argument("--no-pr", action="store_true", help="Update files but don't open a PR")
    parser.add_argument("--interactive", action="store_true", help="Prompt for topic classification")
    args = parser.parse_args()

    print("Loading members...")
    members, author_ids = load_members()
    print(f"  Found {len(author_ids)} Semantic Scholar author IDs")

    print("Fetching publications from Semantic Scholar...")
    all_papers = []
    for aid in author_ids:
        papers = fetch_author_papers(aid)
        print(f"  Author {aid}: {len(papers)} papers")
        all_papers.extend(papers)
        time.sleep(1)

    print("Deduplicating...")
    unique = deduplicate_papers(all_papers)
    print(f"  {len(unique)} unique papers")

    fetched_records = [paper_to_record(p) for p in unique]

    existing = load_existing_pubs()
    print(f"  {len(existing)} papers already in database")

    new_papers = find_new_papers(fetched_records, existing)
    print(f"  {len(new_papers)} new papers found")

    if not new_papers:
        print("No new publications. Done.")
        return

    if args.dry_run:
        print("\n--- DRY RUN: New papers ---")
        for p in new_papers:
            topics_str = ", ".join(p["topics"])
            print(f"  [{topics_str}] {p['title']} ({p['year']})")
        return

    if args.interactive:
        interactive_classify(new_papers)

    all_pubs = existing + new_papers
    all_pubs.sort(key=lambda p: (-(p.get("year") or 0), p.get("title", "")))

    print("Writing publications.json...")
    with open(PUBS_FILE, "w") as f:
        json.dump(all_pubs, f, indent=2, ensure_ascii=False)

    print(f"Updated {len(all_pubs)} publications ({len(new_papers)} new)")

    if not args.no_pr:
        print("Opening pull request...")
        open_pr(new_papers)
    else:
        print("Skipping PR (--no-pr flag)")


if __name__ == "__main__":
    main()
