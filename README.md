# Pokémon GBL Analyzer

Compares Pokémon GO Battle League rankings across two points in time and exports the results to a formatted Excel workbook. For each league (Great, Ultra, Master) it shows ranking changes, moveset updates, buffs/nerfs, newly available attacks, XL candy requirements, and a type-level summary.

---

## Requirements

**Python 3.8+**

Install dependencies:

```bash
pip install pandas openpyxl requests
```

---

## File Structure

The script expects the following files to be present alongside it:

```
your-folder/
├── RankingDifference.py
└── Game Data/
    └── xl_table.json
```

`xl_table.json` contains XL candy level mappings for shadow and non-shadow Pokémon. It must exist before running the script.

---

## Usage

Run the script from the command line:

```bash
python RankingDifference.py [--date YYYY-MM-DD] [--prev-date YYYY-MM-DD] [--branch BRANCH] [--prev-branch BRANCH] [--cup CUP] [--prev-cup CUP]
```

All arguments are optional. When run with no arguments, the script defaults to `--branch master`, `--cup all`, and both `--date` and `--prev-date` set to today — so both snapshots resolve to the same latest commit, giving a baseline no-change comparison.

---

## Arguments

| Argument          | Short   | Description                                                                          | Default                    |
| ----------------- | ------- | ------------------------------------------------------------------------------------ | -------------------------- |
| `--date`        | `-d`  | Cutoff date for the **current** snapshot in `YYYY-MM-DD` format                     | Today                      |
| `--prev-date`   | `-pd` | Cutoff date for the **previous** snapshot in `YYYY-MM-DD` format                    | Today                      |
| `--branch`      | `-b`  | Branch to use for the **current** (present) data                                    | `master`                 |
| `--prev-branch` | `-pb` | Branch to use for the **previous** snapshot                                          | Same as `--branch`       |
| `--cup`         | `-c`  | Cup folder name for the **current** rankings (e.g. `sunshine`, `remix`, `all`)      | `all`                    |
| `--prev-cup`    | `-pc` | Cup folder name for the **previous** rankings (e.g. `all`, `remix`)                 | Same as `--cup`          |

### How `--cup` works

Cup names correspond to folders under `src/data/rankings/` in the [pvpoke/pvpoke](https://github.com/pvpoke/pvpoke) repository. For example, `--cup sunshine` pulls from:

```
src/data/rankings/sunshine/overall/rankings-1500.json   → Great League
src/data/rankings/sunshine/overall/rankings-2500.json   → Ultra League
src/data/rankings/sunshine/overall/rankings-10000.json  → Master League
```

The script automatically detects which of the three ranking files exist in the cup folder and only processes those leagues. If a cup only has a `rankings-1500.json`, only the Great League sheet will be generated.

When two different cups are specified, only leagues present in **both** cups are compared.

---

## Examples

**Compare the latest master branch against a past date on master**

```bash
python RankingDifference.py --prev-date 2025-01-01
```

Loads the latest data from `master` as "current", and finds the most recent commit on `master` on or before January 1, 2025 as "previous".

---

**Pin both snapshots to specific dates**

```bash
python RankingDifference.py --prev-date 2025-01-01 --date 2025-04-01
```

Finds the most recent commit on or before January 1, 2025 as "previous", and the most recent commit on or before April 1, 2025 as "current".

---

**Compare two different branches (no date)**

```bash
python RankingDifference.py --branch my-feature-branch --prev-branch master
```

Compares the latest commit on `my-feature-branch` (current) against the latest commit on `master` (previous), both as of today.

---

**Compare a feature branch against an older date on a different branch**

```bash
python RankingDifference.py --branch my-feature --prev-branch master --prev-date 2025-03-15
```

Loads the HEAD of `my-feature` as current, and finds the most recent commit on `master` on or before March 15, 2025 as previous.

---

**Compare a specific cup against the overall rankings**

```bash
python RankingDifference.py --cup sunshine --prev-cup all
```

Compares the current `sunshine` cup rankings (current) against the standard `all` rankings (previous) on the `master` branch.

---

**Compare a specific cup against its own previous state**

```bash
python RankingDifference.py --cup remix --prev-date 2025-04-01
```

Compares the current `remix` cup rankings against where the `remix` cup stood on or before April 1, 2025.

---

**Combine cup and branch arguments**

```bash
python RankingDifference.py --cup sunshine --prev-cup all --branch my-feature --prev-branch master
```

Compares `sunshine` cup data from `my-feature` (current) against `all` cup data from `master` (previous).

---

**Run with short flags**

```bash
python RankingDifference.py -b my-feature -pb master -pd 2025-03-15
```

Equivalent to the feature branch example above, using short flags.

---

## Output

The script saves `Pokemon GBL Analyzer.xlsx` in the same directory as the script. It contains up to 6 sheets, depending on which leagues are available in the selected cup(s):

| Sheet                 | Contents                                                   |
| --------------------- | ---------------------------------------------------------- |
| `Great – Pokemon`   | Per-Pokémon ranking changes for Great League (1500 CP)    |
| `Great – Types`     | Average ranking changes by type for Great League           |
| `Ultra – Pokemon`   | Per-Pokémon ranking changes for Ultra League (2500 CP)    |
| `Ultra – Types`     | Average ranking changes by type for Ultra League           |
| `Master – Pokemon`  | Per-Pokémon ranking changes for Master League (no CP cap) |
| `Master – Types`    | Average ranking changes by type for Master League          |

The sheet header also includes the cup name(s) when comparing across cups, e.g. `Great – All → Sunshine – 05/01/2025 to 06/16/2025`.

### Columns (Pokémon sheets)

| Column                     | Description                                                  |
| -------------------------- | ------------------------------------------------------------ |
| Pokemon                    | Pokémon name                                                |
| Old Ranking                | Score in the previous snapshot                               |
| Ranking                    | Score in the current snapshot                                |
| Difference                 | Change in score (current − previous)                        |
| Fast Move                  | Best fast move in current rankings                           |
| Charged Move 1 / 2         | Best charged moves in current rankings                       |
| Moveset Change             | Any changes to the recommended moveset                       |
| Update                     | Summary label: Buff, Nerf, New Move, Rework, or combinations |
| Attack Availability        | Newly accessible moves for this Pokémon                     |
| Buffs / Nerfs / Rework     | Specific moves that were buffed, nerfed, or reworked         |
| XL                         | XL candy requirement (Great and Ultra leagues only)          |
| Level                      | Level at the league's CP cap                                 |
| Attack / Defense / Stamina | Base stats                                                   |
| Bulk                       | Defense × Stamina                                           |
| Stat Product               | Attack × Defense × Stamina                                 |
| Types                      | Pokémon typing                                              |

### Cell formatting

* **Green fill** — move was buffed this update
* **Red fill** — move was nerfed this update
* **Blue fill** — move was reworked (both buffed and nerfed in different ways)
* **Bold** — move is newly added to this Pokémon's recommended moveset
* **BOLD + ALL CAPS** — move is brand new to the game in this update
* **Color scale** — applied to Old Ranking, Ranking, and Difference columns (red → yellow → green)

---

## Notes

* Data is pulled live from the [pvpoke/pvpoke](https://github.com/pvpoke/pvpoke) GitHub repository. An internet connection is required.
* Cup folder names must match an existing folder under `src/data/rankings/` in the pvpoke repo. An error is raised if the folder is not found.
* The GitHub API has an unauthenticated rate limit of 60 requests per hour. If you hit this limit, wait an hour or [add a personal access token](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/managing-your-personal-access-tokens) to your requests.
* Each run overwrites the previous `Pokemon GBL Analyzer.xlsx` file.