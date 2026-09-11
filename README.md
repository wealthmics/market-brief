# MICS Weekly Market Brief automation

Every Saturday at 5:00 AM IST, GitHub takes the newsletter you put in the
`input/` folder, drops in the week's fresh index returns, sets the date, sets the
issue number, and saves the finished file into the `latest/` folder. When you wake
up, open `latest/` and download the file.

## The two folders you care about

- `input/` — put your newsletter here each week (with the story written). The
  file can have ANY name. If there is more than one, the newest is used.
- `latest/` — the finished file lands here. Download `Global_Market_Brief_latest.html`.
  A dated copy and the Excel file are kept here too as an archive.

## The rest of the files

- `weekly_market_snapshot_ExcelHTML.py` — your data script (unchanged).
- `stitch.py` — puts the fresh table in, sets the date, sets the issue number.
- `issue.txt` — created automatically; remembers the issue count so it goes up on
  its own. You never edit this.
- `.github/workflows/weekly.yml` — the schedule and steps.

## One-time setup

1. Create a new GitHub repository (Private).
2. Upload every file here, keeping the folder structure (the workflow must stay at
   `.github/workflows/weekly.yml`, and keep the `input/` and `latest/` folders).
3. Settings > Actions > General > Workflow permissions: choose
   "Read and write permissions", then Save.

## Your weekly job

Just one thing: put your finished newsletter (any name) into the `input/` folder
before Saturday 5:00 AM IST. The table, date, and issue number are handled for
you. Saturday morning, open `latest/` and download the file.

The issue number counts up by itself starting from whatever number is in your very
first uploaded file, so you do not have to set it each week.

## Test it now

Actions tab > "Weekly Market Brief" > "Run workflow". On a weekday it will label
the data as partial, which is fine for a test.

## Notes

- US markets close at 1:30 AM IST (summer) or 2:30 AM IST (winter), so the 5:00 AM
  run always has settled Friday prices.
- If a run fails, the usual cause is the market data source being rate-limited from
  GitHub's servers. It usually works; tell me if it starts failing often.
