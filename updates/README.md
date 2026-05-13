# Updating the MUSE Lab website

You do not need to touch any HTML or CSS. Everything that changes often (publications, news, members) lives in Excel sheets inside this folder.

The flow is always the same:

1. Edit the Excel sheets in this folder.
2. From the project root, double-click one of these:
   - `update-site.command` on a Mac
   - `update-site.bat` on Windows
3. Look at `git status` or open GitHub Desktop to see what changed.
4. Commit and push.

The first time you double-click the launcher, it sets up its own Python environment inside a hidden `.venv/` folder. That step takes about 30 seconds. Every run after that is fast.

You need Python 3 installed on the machine. If it is missing, the launcher will tell you where to download it.

If you would rather run the script from the terminal, this is the same as the double-click flow:

```
python3 -m venv .venv
.venv/bin/pip install -r scripts/requirements.txt
.venv/bin/python scripts/update_site.py
```

You only need the first two lines the very first time.

---

## Publications

The file `publications.xlsx` holds every paper that appears on the site. Each paper is one row. Each topic has its own column. To mark a paper with a topic, put an `X` in that column.

- To add a paper, scroll to the first empty row and fill it in.
- To edit a paper, find its row and change the cells. The site picks up the change the next time you run the script.
- To delete a paper, delete its row.

Columns marked with `*` are required. The topic columns are Open Source, EDI, Mining Software Repos, Developer Productivity, ML for SE, Tools, and Empirical SE. Put an `X` in every topic that applies and leave the rest blank.

If you have a slides PDF on your computer, drop it into `updates/publications/files/` and put just the filename in the Slides URL column. For example, `icse-2026.pdf`.

Running the script twice in a row with no Excel changes does nothing. Each run fully rebuilds the JSON from the Excel sheet, so you cannot end up with duplicate papers.

## Adding or editing a member

Each member has their own folder under `updates/members/`. The folder name is the person's name in lowercase with hyphens, for example `alice-smith`.

**To add a new member:**

1. Copy the `updates/members/_template/` folder.
2. Rename the copy to the person's name, for example `updates/members/alice-smith/`.
3. Open `info.xlsx` and fill in the Value column.
4. Drop their photo into the same folder. Put the photo's filename in the Photo Filename field, for example `alice.jpg`.
5. Save the file and run the launcher.

**To edit an existing member:** open the `info.xlsx` in their folder, change the values, and run the launcher again.

**To remove a member:** delete their folder and run the launcher again.

**To move someone to alumni:** open their `info.xlsx` and change the Status field to `Alumni`. They will move out of the active list and into the alumni section the next time the launcher runs.

## Adding news

1. Open `updates/news/news.xlsx`.
2. Add a new row. The columns are:
   - **Date Badge** is the short label that appears on the news card. For example, `May 2026`.
   - **Sort Date** controls the order. Use the format `YYYY-MM`. The newest dates appear at the top.
   - **Text** is the news body. You can use `<strong>` and `<em>` for emphasis.
   - **Link URL** and **Link Label** are optional. Use them when you want to link out to a preprint or a paper.
   - **Image Filename** and **Image Alt** are optional. If you want an image, drop it into `updates/news/images/` and put the filename in Image Filename.
3. Save the file and run the launcher.

To remove a news item, delete its row. To edit one, change the values in its row.

---

## What the script changes

| File you edit                            | What gets updated                       |
| ---------------------------------------- | --------------------------------------- |
| `updates/publications/publications.xlsx` | `data/publications.json`                |
| `updates/members/<name>/info.xlsx`       | `data/members.json` and `members.html`  |
| `updates/news/news.xlsx`                 | `data/news.json` and `index.html`       |

The script also copies any new images or files into the right folders, like `photos/` and `files/`.
