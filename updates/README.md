# Updating the MUSE Lab website

You don't need to touch any HTML or CSS. Everything that changes regularly
(publications, news, members) is driven by Excel sheets in this folder.

The flow is always the same:

1. Edit the Excel sheets in this folder.
2. From the project root, run:
   ```
   python scripts/update_site.py
   ```
3. Look at `git status` to see what changed.
4. Commit the changes.

If you've never run the script before, install the dependencies first:

```
python3 -m venv .venv
.venv/bin/pip install -r scripts/requirements.txt
.venv/bin/python scripts/update_site.py
```

---

## Adding a publication

1. Open `updates/publications/publications.xlsx`.
2. Scroll to the first empty row and add the paper. The columns marked with a `*`
   are required.
3. For **Type**, click the cell and pick from the dropdown (`journal`,
   `conference`, `book-chapter`, or `preprint`).
4. For **Topics**, enter a comma-separated list. The valid codes are:
   - `oss` (Open Source)
   - `edi` (EDI)
   - `msr` (Mining Software Repos)
   - `productivity` (Developer Productivity)
   - `ml4se` (ML for SE)
   - `tools` (Tools)
   - `ese` (Empirical SE)

   Example: `oss, msr`
5. If you have a slides PDF, drop it in `updates/publications/files/` and enter
   just the filename (e.g. `icse-2026-slides.pdf`) in the **Slides URL** column.
6. Save and run `python scripts/update_site.py`.

Duplicate rows (same DOI or title) update existing entries instead of creating
new ones, so it's safe to edit a paper by changing its row.

## Adding or editing a member

Each member has a folder under `updates/members/`. The folder name is the
person's name in lowercase with hyphens, like `alice-smith`.

**To add a new member:**

1. Copy the `updates/members/_template/` folder.
2. Rename the copy to the person's name, e.g. `updates/members/alice-smith/`.
3. Open `info.xlsx` and fill in the **Value** column.
4. Drop their photo file into the same folder. Use the **Photo Filename** field
   to point to it (e.g. `alice.jpg`).
5. Save and run `python scripts/update_site.py`.

**To edit an existing member:** open their folder's `info.xlsx`, change the
values, and re-run the script.

**To remove a member:** delete their folder and re-run the script.

**To make someone alumni:** change their **Status** to `Alumni`. They'll move
from the active list to the alumni section automatically.

## Adding news

1. Open `updates/news/news.xlsx`.
2. Add a row. The columns are:
   - **Date Badge** — short label that appears in the badge, e.g. `May 2026`.
   - **Sort Date** — `YYYY-MM` format. Newest dates appear at the top.
   - **Text** — the news body. You can use `<strong>` and `<em>` for emphasis.
   - **Link URL / Link Label** — optional link (preprint, paper, etc.).
   - **Image Filename / Image Alt** — optional. Drop the image in
     `updates/news/images/` and enter just the filename.
3. Save and run `python scripts/update_site.py`.

To remove a news item, delete the row. To edit, change the row.

---

## What does the script change?

| File you edit                           | What gets updated                     |
| --------------------------------------- | ------------------------------------- |
| `updates/publications/publications.xlsx`| `data/publications.json`              |
| `updates/members/<name>/info.xlsx`      | `data/members.json` and `members.html`|
| `updates/news/news.xlsx`                | `data/news.json` and `index.html`     |

The script also copies any new images or files into the right folders
(`photos/`, `files/`).
