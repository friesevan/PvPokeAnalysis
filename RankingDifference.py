import json
import os
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
import requests
from datetime import datetime

# ----------------------------
# Helpers
# ----------------------------

def get_commit_date(repo_owner, repo_name, commit_sha):
    """
    Returns the ISO date string of a specific commit.
    """
    url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/commits/{commit_sha}"
    response = requests.get(url)
    response.raise_for_status()
    data = response.json()
    return data["commit"]["committer"]["date"]  # ISO 8601 format, e.g., "2025-12-21T18:22:03Z"


def get_latest_commit_info(repo_owner, repo_name, path):
    """
    Returns the SHA and date of the latest commit for a given file path in ISO format.
    """
    url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/commits"
    params = {"path": path, "per_page": 1}  # Only need the latest commit
    response = requests.get(url, params=params)
    response.raise_for_status()
    data = response.json()[0]  # Latest commit is first
    sha = data["sha"]
    date_iso = data["commit"]["committer"]["date"]  # ISO 8601 string
    return sha, date_iso

def get_commit_sha(repo_owner, repo_name, until_date, path=None, branch="master"):
    """
    Returns the most recent commit SHA on or before until_date (ISO8601 string)
    """
    url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/commits"
    params = {
        "sha": branch,
        "per_page": 1,
        "until": until_date
    }
    if path:
        params["path"] = path

    response = requests.get(url, params=params)
    response.raise_for_status()
    commits = response.json()
    if not commits:
        raise ValueError("No commits found for that date")
    return commits[0]["sha"]

def load_json(path):
    with open(path, "r", encoding="utf-8-sig") as f:
        return json.load(f)
    
def load_json_from_github(url):
    """
    Load JSON from a raw GitHub URL.
    """
    response = requests.get(url)
    response.raise_for_status()
    return response.json()

def load_json_from_commit(path, commit_hash):
    """
    Load a JSON file from a specific commit.
    path: file path relative to repo root
    commit_hash: commit SHA string
    """
    url = f"https://raw.githubusercontent.com/pvpoke/pvpoke/{commit_hash}/{path}"
    return load_json_from_github(url)

def build_rankings_map(rankings):
    return {r["speciesId"]: r for r in rankings}

def build_pokemon_map(pokemon_data):
    return {p["speciesId"]: p for p in pokemon_data}

def stat_product(atk, defense, stamina):
    return atk * defense * stamina

def is_shadow(species_id, pokemon_entry):
    return "shadow" in species_id or ("tags" in pokemon_entry and "shadow" in pokemon_entry["tags"])

def calculate_xl(pokemon_entry, league, xl_table, shadow=False):
    default_ivs = pokemon_entry.get("defaultIVs", {})
    cp_key = {"Great": "cp1500", "Ultra": "cp2500"}.get(league)
    if not cp_key or cp_key not in default_ivs:
        return "", ""
    level = default_ivs[cp_key][0]
    if level <= 40:
        return "", level
    table_key = "shadow" if shadow else "non_shadow"
    return xl_table[table_key].get(str(level), ""), level

def generate_update_text(row):
    updates = []

    if row.get("Buffs"):
        updates.append("Buff")

    if row.get("Nerfs"):
        updates.append("Nerf")

    if row.get("Attack Availability"):
        new_moves = [m.strip() for m in row["Attack Availability"].split(",") if m.strip()]
        num_new = len(new_moves)

        if num_new == 1:
            updates.append("New Move")
        elif num_new > 1:
            updates.append(f"New Moves ({num_new})")

    if row.get("Rework"):
        updates.append("Rework")

    if not updates:
        return ""

    if len(updates) == 1:
        return updates[0]

    if len(updates) == 2:
        return f"{updates[0]} and {updates[1]}"

    return f"{updates[0]}, {updates[1]}, and {updates[2]}"

def format_moves(moves):
    return "\n".join([m.replace("_", " ").title() for m in moves])

def format_move_name(move_id):
    if not move_id:
        return ""

    move = move_lookup.get(move_id)
    if move:
        return move["name"]

    return move_id.replace("_", " ").title()

def clean_moveset_change(old_moves, new_moves):
    """
    Compute moveset change text for a Pokémon moveset.
    old_moves, new_moves: lists of 3 moves [fast, charged1, charged2]
    """
    changes = []

    # --- Fast move ---
    if old_moves[0] != new_moves[0]:
        changes.append(f"{format_move_name(old_moves[0])} → {format_move_name(new_moves[0])}")

    # --- Charged moves ---
    old_charged = [m for m in old_moves[1:] if m]
    new_charged = [m for m in new_moves[1:] if m]

    # Identify truly removed and added moves
    removed = [m for m in old_charged if m not in new_charged]
    added = [m for m in new_charged if m not in old_charged]

    # Pair removed and added moves (1:1)
    for rm, ad in zip(removed, added):
        changes.append(f"{format_move_name(rm)} → {format_move_name(ad)}")

    # Any leftover added moves
    for ad in added[len(removed):]:
        changes.append(f"(added) → {format_move_name(ad)}")

    # Any leftover removed moves
    for rm in removed[len(added):]:
        changes.append(f"{format_move_name(rm)} → (removed)")

    return "\n".join(changes)

def collect_move_updates(update_json):
    updates = []

    for move_id in update_json.get("buffs", []):
        updates.append((move_id, "Buff"))

    for move_id in update_json.get("nerfs", []):
        updates.append((move_id, "Nerf"))

    for move_id in update_json.get("rework", []):
        updates.append((move_id, "Rework"))

    return updates

def build_type_update_text(pokemon_type, move_updates):
    """
    pokemon_type: 'ice', 'fire', etc.
    """
    matched_updates = []

    for move_id, label in move_updates:
        move = move_lookup.get(move_id)
        if not move:
            continue

        if move["type"] == pokemon_type:
            matched_updates.append(f"{move['name']} {label}")

    return ", ".join(sorted(matched_updates))

# Base species key for updates
def get_base_species_key(pokemon_entry):
    species_id = pokemon_entry["speciesId"].upper()
    # Remove _SHADOW suffix if present
    if species_id.endswith("_SHADOW"):
        species_id = species_id.replace("_SHADOW", "")
    return species_id


# ----------------------------
# League processing
# ----------------------------
def process_league(league, seasons, updates, pokemon_map, xl_table,
                   rankings_current=None, rankings_previous=None):
    league_dir = os.path.join("Rankings", league)
    sorted_seasons = sorted(seasons.items(), key=lambda x: x[1])
    old_season, new_season = sorted_seasons[0][0], sorted_seasons[-1][0]

    old_rankings = load_json(os.path.join(league_dir, f"{league.lower()}-{old_season}.json"))
    new_rankings = load_json(os.path.join(league_dir, f"{league.lower()}-{new_season}.json"))

    old_map = build_rankings_map(rankings_previous[league])
    new_map = build_rankings_map(rankings_current[league])

    buffed_moves = set(updates["buffs"])
    nerfed_moves = set(updates["nerfs"])
    rework_moves = set(updates.get("rework", []))
    availability = updates["availability_update"]

    rows = []

    for species_id, new_entry in new_map.items():
        if species_id not in old_map or species_id not in pokemon_map:
            continue
        old_entry = old_map[species_id]
        pokemon = pokemon_map[species_id]

        old_score = old_entry.get("score")
        new_score = new_entry.get("score")
        diff = round(new_score - old_score, 2) if old_score is not None and new_score is not None else None

        atk, defense, stamina = pokemon["baseStats"]["atk"], pokemon["baseStats"]["def"], pokemon["baseStats"]["hp"]
        bulk, sp = defense * stamina, stat_product(atk, defense, stamina)

        old_moveset = old_entry.get("moveset", ["", "", ""])
        new_moveset = new_entry.get("moveset", ["", "", ""])
        while len(old_moveset) < 3:
            old_moveset.append("")
        while len(new_moveset) < 3:
            new_moveset.append("")

        # Moveset change
        moveset_change_text = clean_moveset_change(old_moveset, new_moveset)

        # Buffs/Nerfs/Rework
        all_moveset = set(old_moveset + new_moveset)
        base_species_key = get_base_species_key(pokemon)

        # Apply move updates from the base species key (handles shadow automatically)
        avail_moves = [m for m in availability.get(base_species_key, []) if m in new_moveset]
        buffs_used = sorted(m for m in buffed_moves if m in new_moveset)
        nerfs_used = sorted(m for m in nerfed_moves if m in old_moveset)
        reworks_used = sorted(m for m in all_moveset if m in rework_moves)

        # XL
        shadow = is_shadow(species_id, pokemon)
        xl_value, level = calculate_xl(pokemon, league, xl_table, shadow=shadow)

        types = [t for t in pokemon.get("types", []) if t.lower() != "none"]
        types_str = ", ".join([t.title() for t in types])

        rows.append({
            "Pokemon": pokemon["speciesName"],
            "Old Ranking": old_score,
            "Ranking": new_score,
            "Difference": diff,
            "XL": xl_value,
            "Level": level,
            "Attack": atk,
            "Defense": defense,
            "Stamina": stamina,
            "Bulk": bulk,
            "Stat Product": sp,
            "Shadow": shadow,
            "Types": types_str,
            "Attack Availability": ", ".join(avail_moves),
            "Buffs": ", ".join(buffs_used),
            "Nerfs": ", ".join(nerfs_used),
            "Rework": ", ".join(reworks_used),
            "Fast Move": format_move_name(new_moveset[0]),
            "Charged Move 1": format_move_name(new_moveset[1]),
            "Charged Move 2": format_move_name(new_moveset[2]),
            "Moveset Change": moveset_change_text
        })

    return pd.DataFrame(rows)

def append_average_row(ws, df, label_col=1, header_row=2):
    avg_row = ws.max_row + 1
    ws.cell(row=avg_row, column=label_col, value="Average")
    ws.cell(row=avg_row, column=label_col).font = Font(bold=True)
    headers = [c.value for c in ws[header_row]]
    for col_name in ["Old Ranking", "Ranking", "Difference"]:
        if col_name in headers:
            col_idx = headers.index(col_name) + 1
            avg_val = round(df[col_name].mean(), 2)
            ws.cell(row=avg_row, column=col_idx, value=avg_val)
            ws.cell(row=avg_row, column=col_idx).font = Font(bold=True)
            ws.cell(row=avg_row, column=col_idx).alignment = Alignment(horizontal="center")

def normalize_move(move):
    if not move:
        return ""
    return move.replace("_", " ").strip().title()


# ----------------------------
# Main
# ----------------------------

def main(previous_date=None):

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    # ----------------------------
    # GitHub URLs
    # ----------------------------
    REPO_BASE = "https://raw.githubusercontent.com/pvpoke/pvpoke/master"

    pokemon_url = f"{REPO_BASE}/src/data/gamemaster/pokemon.json"
    moves_url   = f"{REPO_BASE}/src/data/gamemaster/moves.json"
    rankings_urls = {
        "Great": f"{REPO_BASE}/src/data/rankings/all/overall/rankings-1500.json",
        "Ultra": f"{REPO_BASE}/src/data/rankings/all/overall/rankings-2500.json",
        "Master": f"{REPO_BASE}/src/data/rankings/all/overall/rankings-10000.json",
    }

    pokemon_data = load_json_from_github(pokemon_url) 
    moves_data = load_json_from_github(moves_url) 
    rankings_current = {league: load_json_from_github(url) for league, url in rankings_urls.items()}

    # Latest commit date (we assume master branch latest)
    latest_commit_sha, latest_commit_date_iso = get_latest_commit_info(
        repo_owner="pvpoke",
        repo_name="pvpoke",
        path="src/data/gamemaster/pokemon.json"
    )
    latest_commit_date = datetime.strptime(latest_commit_date_iso[:10], "%Y-%m-%d").strftime("%m/%d/%Y")

    previous_commit_date = None
    rankings_prev = {}
    if previous_date:
        previous_commit_sha = get_commit_sha(
            repo_owner="pvpoke",
            repo_name="pvpoke",
            until_date=previous_date,
            path="src/data/gamemaster/pokemon.json"
        )
        # Pull actual commit date for previous SHA
        previous_commit_date_iso = get_commit_date(
            repo_owner="pvpoke",
            repo_name="pvpoke",
            commit_sha=previous_commit_sha
        )
        previous_commit_date = datetime.strptime(previous_commit_date_iso[:10], "%Y-%m-%d").strftime("%m/%d/%Y")

        pokemon_prev = load_json_from_commit("src/data/gamemaster/pokemon.json", previous_commit_sha)
        moves_prev   = load_json_from_commit("src/data/gamemaster/moves.json", previous_commit_sha)
        rankings_prev = {
            league: load_json_from_commit(
                f"src/data/rankings/all/overall/rankings-{suffix}.json",
                previous_commit_sha
            )
            for league, suffix in [("Great", "1500"), ("Ultra", "2500"), ("Master", "10000")]
        }
    else:
        previous_commit_date = latest_commit_date  # if no previous, same as latest
        pokemon_prev = pokemon_data
        moves_prev   = moves_data
        rankings_prev = rankings_current


    # ----------------------------
    # Build lookup tables
    # ----------------------------
    global move_lookup
    move_lookup = {m["moveId"]: {"name": m["name"], "type": m["type"].lower()} for m in moves_data}
    pokemon_map = {p["speciesId"]: p for p in pokemon_data}

    # ----------------------------
    # Load updates JSON (local file)
    # ----------------------------
    updates_path = os.path.join(BASE_DIR, "Updates", "precious-paths-update.json")
    updates = load_json(updates_path)[0]

    # ----------------------------
    # Seasons (keep your existing seasons.json if needed)
    # ----------------------------
    seasons_path = os.path.join(BASE_DIR, "Game Data", "seasons.json")
    seasons = load_json(seasons_path)[0]

    latest_season = max(seasons, key=seasons.get)
    latest_season_name = latest_season.replace("-", " ").title() + " Update"
    output_path = os.path.join(BASE_DIR, f"{latest_season_name}.xlsx")
    xl_table = load_json(os.path.join(BASE_DIR, "Game Data", "xl_table.json"))
    leagues = ["Great", "Ultra", "Master"]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        move_updates = collect_move_updates(updates)
        for league in leagues:
            df = process_league(
                league,
                seasons,
                updates,
                pokemon_map,
                xl_table,
                rankings_current=rankings_current,
                rankings_previous=rankings_prev
            )
            if df.empty:
                continue

            # Update column
            df.insert(df.columns.get_loc("Difference") + 1, "Update", df.apply(generate_update_text, axis=1))

            # Remove XL for Master League
            if league == "Master" and "XL" in df.columns:
                df.drop(columns=["XL"], inplace=True)

            # Reorder columns
            cols_order = [
                "Pokemon", "Old Ranking", "Ranking", "Difference", 
                "Fast Move", "Charged Move 1", "Charged Move 2", "Moveset Change",
                "Update", "Attack Availability", "Buffs", "Nerfs", "Rework", "XL", "Level",
                "Attack", "Defense", "Stamina", "Bulk", "Stat Product", "Types"
            ]
            df = df[[c for c in cols_order if c in df.columns]]

            # Sort BEFORE writing to Excel so row indices match
            df_sorted = df.sort_values(by="Difference", ascending=False, na_position="last").reset_index(drop=True)

            # Helper to convert comma-separated moves to sentence case
            def title_case_moves(text):
                if not text:
                    return ""
                move_ids = [m.strip() for m in text.split(",") if m.strip()]
                formatted = []
                for move_id in move_ids:
                    if move_id in move_lookup:
                        formatted.append(move_lookup[move_id]["name"])
                    else:
                        formatted.append(move_id.replace("_", " ").title())
                return ", ".join(formatted)

            # Apply sentence case to relevant columns
            for col in ["Attack Availability", "Buffs", "Nerfs", "Rework"]:
                if col in df_sorted.columns:
                    df_sorted[col] = df_sorted[col].apply(title_case_moves)

            # Write Pokémon sheet
            sheet_pokemon = f"{league} – Pokemon"
            df_sorted.to_excel(writer, sheet_name=sheet_pokemon, index=False)
            ws = writer.sheets[sheet_pokemon]

            # --- Add date range header above the column headers ---
            ws.insert_rows(1)  # insert new first row
            header_range = f"A1:{get_column_letter(ws.max_column)}1"
            ws.merge_cells(header_range)
            ws["A1"].value = f"{league} – Pokemon – {previous_commit_date} to {latest_commit_date}"
            ws["A1"].font = Font(bold=True, size=14)
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

            # Format actual headers (row 2)
            for cell in ws[2]:
                cell.font = Font(bold=True, size=12)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # Bold Pokémon names (now in row 3+)
            for cell in ws["A"][2:]:
                cell.font = Font(bold=True, size=12)

            # Auto-fit columns
            for col_idx, col_cells in enumerate(ws.columns):
                max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col_cells)
                ws.column_dimensions[get_column_letter(col_idx + 1)].width = max_len + 5
                for cell in col_cells:
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # Filters on row 2
            ws.auto_filter.ref = f"A2:{get_column_letter(ws.max_column)}{ws.max_row}"

            # Conditional formatting: use row 3+ for data
            headers = [c.value for c in ws[2]]
            for col_name in ["Old Ranking", "Ranking", "Difference"]:
                if col_name in headers:
                    col_letter = get_column_letter(headers.index(col_name) + 1)
                    start_row = 3
                    if col_name == "Difference":
                        ws.conditional_formatting.add(
                            f"{col_letter}{start_row}:{col_letter}{ws.max_row}",
                            ColorScaleRule(start_type='num', start_value=-10, start_color='FF0000',
                                        mid_type='num', mid_value=0, mid_color='FFFF00',
                                        end_type='num', end_value=10, end_color='00FF00')
                        )
                    else:
                        ws.conditional_formatting.add(
                            f"{col_letter}{start_row}:{col_letter}{ws.max_row}",
                            ColorScaleRule(start_type='min', start_color='FF0000',
                                        mid_type='percentile', mid_value=50, mid_color='FFFF00',
                                        end_type='max', end_color='00FF00')
                        )

            # Append average row (after last row)
            append_average_row(ws, df_sorted, header_row=2)

            
            new_move_fill = PatternFill(start_color="A8D5BA", end_color="A8D5BA", fill_type="solid")   # soft green
            rework_fill   = PatternFill(start_color="C6A3C6", end_color="C6A3C6", fill_type="solid")   # muted purple
            buff_fill     = PatternFill(start_color="A3C4F3", end_color="A3C4F3", fill_type="solid")   # muted blue
            nerf_fill     = PatternFill(start_color="F4A3A3", end_color="F4A3A3", fill_type="solid")   # soft red

            # Highlight moves with debug prints (adjusted for date header)
            for row_idx, row in df_sorted.iterrows():
                styled_cols = set()

                # Normalize all move column values
                move_values = {col: normalize_move(row.get(col, "")) for col in ["Fast Move", "Charged Move 1", "Charged Move 2"]}

                # Data now starts at Excel row 3 because of date header + header row
                excel_row = row_idx + 3

                # 2. Buffed moves (blue)
                for move in [normalize_move(m) for m in row.get("Buffs", "").split(",") if m.strip()]:
                    for move_col, cell_value in move_values.items():
                        if cell_value == move:
                            ws.cell(row=excel_row, column=df_sorted.columns.get_loc(move_col) + 1).fill = buff_fill
                            styled_cols.add(move_col)

                # 1. Newly added moves (green + bold)
                for move in [normalize_move(m) for m in row.get("Attack Availability", "").split(",") if m.strip()]:
                    for move_col, cell_value in move_values.items():
                        if cell_value == move:
                            cell = ws.cell(row=excel_row, column=df_sorted.columns.get_loc(move_col) + 1)
                            cell.fill = new_move_fill
                            cell.font = Font(bold=True)
                            styled_cols.add(move_col)

                # 3. Nerfed moves (red)
                for move in [normalize_move(m) for m in row.get("Nerfs", "").split(",") if m.strip()]:
                    for move_col, cell_value in move_values.items():
                        if cell_value == move:
                            ws.cell(row=excel_row, column=df_sorted.columns.get_loc(move_col) + 1).fill = nerf_fill
                            styled_cols.add(move_col)

                # 4. Rework moves (purple)
                for move in [normalize_move(m) for m in row.get("Rework", "").split(",") if m.strip()]:
                    for move_col, cell_value in move_values.items():
                        if cell_value == move:
                            ws.cell(row=excel_row, column=df_sorted.columns.get_loc(move_col) + 1).fill = rework_fill
                            styled_cols.add(move_col)

                # 5. Moves new to Pokémon (bold only)
                for line in row.get("Moveset Change", "").split("\n"):
                    parts = line.split("→")
                    if len(parts) == 2:
                        new_move = normalize_move(parts[1])
                        if new_move not in ["(Removed)", ""]:
                            for move_col, cell_value in move_values.items():
                                if cell_value == new_move:
                                    ws.cell(row=excel_row, column=df_sorted.columns.get_loc(move_col) + 1).font = Font(bold=True)


            # --- Type analysis sheet ---
            type_stats = []
            all_types = sorted({t for ts in df_sorted["Types"].dropna() for t in ts.split(", ")})

            for type_name in all_types:
                type_df = df_sorted[df_sorted["Types"].str.contains(type_name)]
                update_text = build_type_update_text(type_name.lower(), move_updates)

                type_stats.append({
                    "Type": type_name,
                    "Old Ranking": round(type_df["Old Ranking"].mean(), 2),
                    "Ranking": round(type_df["Ranking"].mean(), 2),
                    "Difference": round(type_df["Difference"].mean(), 2),
                    "Update": update_text
                })

            if type_stats:
                df_types = pd.DataFrame(type_stats)
                sheet_types = f"{league} – Types"
                df_types.to_excel(writer, sheet_name=sheet_types, index=False)
                ws_types = writer.sheets[sheet_types]

                # Add date range header above types sheet
                ws_types.insert_rows(1)
                header_range_types = f"A1:{get_column_letter(ws_types.max_column)}1"
                ws_types.merge_cells(header_range_types)
                ws_types["A1"].value = f"{league} – Types – {previous_commit_date} to {latest_commit_date}"
                ws_types["A1"].font = Font(bold=True, size=14)
                ws_types["A1"].alignment = Alignment(horizontal="center", vertical="center")

                # Format actual headers (row 2)
                for cell in ws_types[2]:
                    cell.font = Font(bold=True, size=12)
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                # Auto-fit + alignment for all columns
                for col_idx, col_cells in enumerate(ws_types.columns):
                    max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col_cells)
                    ws_types.column_dimensions[get_column_letter(col_idx + 1)].width = max_len + 5
                    for cell in col_cells:
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                # Filters (start from row 2)
                ws_types.auto_filter.ref = f"A2:{get_column_letter(ws_types.max_column)}{ws_types.max_row}"

                # Make Type names bold (data starts at row 3)
                for row in range(3, ws_types.max_row + 1):
                    ws_types.cell(row=row, column=1).font = Font(bold=True)

                # Conditional formatting for numeric columns
                headers = [c.value for c in ws_types[2]]  # header row is now row 2
                for col_name in ["Old Ranking", "Ranking", "Difference"]:
                    if col_name in headers:
                        col_letter = get_column_letter(headers.index(col_name) + 1)
                        ws_types.conditional_formatting.add(
                            f"{col_letter}3:{col_letter}{ws_types.max_row}",  # start from data row
                            ColorScaleRule(
                                start_type='min', start_color='FF0000',
                                mid_type='percentile', mid_value=50, mid_color='FFFF00',
                                end_type='max', end_color='00FF00'
                            )
                        )


    print(f"Excel file saved as: {latest_season_name}.xlsx")

if __name__ == "__main__":
    import argparse
    from datetime import datetime

    parser = argparse.ArgumentParser(description="Process Pokémon ranking differences.")
    parser.add_argument(
        "--date", "-d",
        type=str,
        help='Date for previous commit in "YYYY-MM-DD" format (defaults to current data)',
        default=None
    )
    args = parser.parse_args()

    previous_date = None
    if args.date:
        try:
            # Convert YYYY-MM-DD to YYYY-MM-DDT23:59:59Z
            dt = datetime.strptime(args.date, "%Y-%m-%d")
            previous_date = dt.strftime("%Y-%m-%dT23:59:59Z")
        except ValueError:
            print("Error: Date must be in YYYY-MM-DD format.")
            exit(1)

    main(previous_date=previous_date)
