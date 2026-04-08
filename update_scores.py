#!/usr/bin/env python3
"""
Masters Pool 2026 - DraftKings Classic Golf Scoring Engine
Pulls hole-by-hole data from ESPN API, calculates DK Classic points,
and pushes scores to Google Sheet.
"""

import json
import os
import tempfile
import urllib.request
import gspread
from collections import defaultdict

# --- CONFIG ---
MASTERS_EVENT_ID = "401811941"
ESPN_URL = f"http://site.api.espn.com/apis/site/v2/sports/golf/pga/scoreboard/{MASTERS_EVENT_ID}"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1POSuW5FvJw39DYHumKHTIsoi3AmKbpMdRaHWEDkP_so/edit"
CREDS_FILE = "credentials.json"


def get_gspread_client():
    """Get gspread client, supporting both local file and env var credentials."""
    creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    if creds_json:
        # Running in CI/GitHub Actions - write temp file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            f.write(creds_json)
            return gspread.service_account(filename=f.name)
    else:
        # Running locally
        return gspread.service_account(filename=CREDS_FILE)

TEAMS = {
    'Depa': ['Cameron Young', 'Hideki Matsuyama', 'Adam Scott', 'Gary Woodland', 'Ryan Gerard', 'Nick Taylor', 'Johnny Keefer', 'Brian Campbell', 'Fred Couples'],
    'Bill': ['Ludvig Aberg', 'Robert MacIntyre', 'Justin Thomas', 'Nicolai Hojgaard', 'Max Homa', 'Wyndham Clark', 'Haotong Li', 'Danny Willett', 'Brandon Holtz'],
    'Nascar1776': ['Xander Schauffele', 'Justin Rose', 'Jason Day', 'Harris English', 'Marco Penge', 'Casey Jarvis', 'Aldrich Potgieter', 'Bubba Watson', 'Mateo Pulcini'],
    'KBM': ['Collin Morikawa', 'Russell Henley', 'Min Woo Lee', 'J.J. Spaun', 'Rasmus Hajgaard', 'Dustin Johnson', 'Michael Kim', 'Michael Brennan', 'Mike Weir'],
    'Jeff W.': ['Matt Fitzpatrick', 'Jordan Spieth', 'Shane Lowry', 'Sam Burns', 'Keegan Bradley', 'Brian Harman', 'Kristoffer Reitan', 'Pongsapak Laopakdee', 'Mason Howell'],
    'Primo': ['Rory McIlroy', 'Viktor Hovland', 'Jake Knapp', 'Jacob Bridgeman', 'Sungjae Im', 'Ryan Fox', 'Andrew Novak', 'Ethan Fang', 'Angel Cabrera'],
    'Wretchy': ['Tommy Fleetwood', 'Chris Gotterup', 'Akshay Bhatia', 'Kurt Kitayama', 'Daniel Berger', 'Aaron Rai', 'Rasmus Neergaard-Petersen', 'Sami Valimaki', 'Naoyuki Kataoka'],
    'Wass': ['Scottie Scheffler', 'Si Woo Kim', 'Tyrrell Hatton', 'Maverick McNealy', 'Ben Griffin', 'Max Greyserman', 'Nico Echavarria', 'Charl Schwartzel', 'Jackson Herrington'],
    'Benba': ['Bryson DeChambeau', 'Patrick Reed', 'Patrick Cantlay', 'Cameron Smith', 'Alex Noren', 'Sam Stevens', 'Tom McKibbin', 'Davis Riley', 'Vijay Singh'],
    'Goose': ['Jon Rahm', 'Brooks Koepka', 'Sepp Straka', 'Corey Conners', 'Harry Hall', 'Sergio Garcia', 'Carlos Ortiz', 'Zach Johnson', 'Jose Maria Olazabal'],
}

# --- DK CLASSIC GOLF SCORING ---
HOLE_SCORES = {
    # score_to_par: points
    -3: 13,   # double eagle or better
    -2: 8,    # eagle
    -1: 3,    # birdie
    0: 0.5,   # par
    1: -0.5,  # bogey
    2: -1,    # double bogey
}
# Worse than double bogey also -1 (handled in code for +3, +4, etc.)

FINISH_POINTS = {
    1: 30, 2: 20, 3: 18, 4: 16, 5: 14, 6: 12, 7: 10, 8: 9, 9: 8, 10: 7,
}
# Ranges
FINISH_RANGES = [
    (11, 15, 6), (16, 20, 5), (21, 25, 4), (26, 30, 3), (31, 40, 2), (41, 50, 1),
]


def parse_score_to_par(score_type_str):
    """Convert ESPN scoreType displayValue to integer score-to-par."""
    if score_type_str == "E":
        return 0
    return int(score_type_str)


def hole_points(score_to_par):
    """Calculate DK points for a single hole."""
    if score_to_par <= -3:
        return 13  # double eagle or better
    if score_to_par >= 2:
        return -1  # double bogey or worse
    return HOLE_SCORES[score_to_par]


def is_birdie_or_better(score_to_par):
    return score_to_par <= -1


def calc_round_bonuses(holes_data, round_strokes):
    """Calculate per-round bonuses. Returns (bonus_pts, details)."""
    bonus = 0
    details = []

    # Bogey-free round (must complete 18 holes)
    if len(holes_data) == 18:
        has_bogey = any(h['score_to_par'] >= 1 for h in holes_data)
        if not has_bogey:
            bonus += 3
            details.append("bogey-free +3")

    # Streak of 3+ birdies or better (max 1 per round)
    streak = 0
    has_streak = False
    for h in holes_data:
        if is_birdie_or_better(h['score_to_par']):
            streak += 1
            if streak >= 3 and not has_streak:
                bonus += 3
                has_streak = True
                details.append("birdie-streak +3")
        else:
            streak = 0

    # Hole in one (+5 each)
    for h in holes_data:
        if h['strokes'] == 1:
            bonus += 5
            details.append(f"hole-in-one hole {h['hole']} +5")

    return bonus, details


def get_finish_points(position):
    """Get DK finish position points."""
    if position in FINISH_POINTS:
        return FINISH_POINTS[position]
    for low, high, pts in FINISH_RANGES:
        if low <= position <= high:
            return pts
    return 0


def calculate_finish_positions(players_data):
    """
    Calculate finish positions handling ties per DK rules:
    Ties don't reduce points - all tied players get the same position's points.
    """
    # Only players who completed all 4 rounds and didn't WD/DQ get finish position
    finishers = []
    for name, data in players_data.items():
        if data['rounds_completed'] == 4 and not data.get('withdrawn'):
            total_strokes = sum(r['strokes'] for r in data['rounds'])
            finishers.append((total_strokes, name))

    finishers.sort()

    positions = {}
    i = 0
    while i < len(finishers):
        # Find all players tied at this stroke total
        current_strokes = finishers[i][0]
        tied = []
        while i < len(finishers) and finishers[i][0] == current_strokes:
            tied.append(finishers[i][1])
            i += 1
        # Position is rank (1-indexed)
        pos = i - len(tied) + 1
        for name in tied:
            positions[name] = pos

    return positions


def fetch_espn_data(event_id=None):
    """Fetch tournament data from ESPN API."""
    url = ESPN_URL if event_id is None else f"http://site.api.espn.com/apis/site/v2/sports/golf/pga/scoreboard/{event_id}"
    req = urllib.request.Request(url)
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read().decode())


def parse_players(espn_data):
    """Parse ESPN data into our player structure."""
    comps = espn_data.get('competitions', [])
    if not comps:
        print("No competition data found. Tournament may not have started.")
        return {}

    competitors = comps[0].get('competitors', [])
    players = {}

    for comp in competitors:
        name = comp['athlete']['fullName']
        linescores = comp.get('linescores', [])

        player = {
            'name': name,
            'espn_score': comp.get('score', ''),
            'rounds_completed': 0,
            'rounds': [],
            'total_hole_pts': 0,
            'total_bonus_pts': 0,
            'finish_pts': 0,
            'total_dk_pts': 0,
            'bonus_details': [],
        }

        for round_data in linescores:
            round_strokes = round_data.get('value')
            if round_strokes is None:
                continue

            round_strokes = int(round_strokes)
            holes = round_data.get('linescores', [])

            holes_parsed = []
            round_hole_pts = 0

            for hole in holes:
                score_type = hole.get('scoreType', {}).get('displayValue')
                strokes = hole.get('value')
                if score_type is None or strokes is None:
                    continue

                score_to_par = parse_score_to_par(score_type)
                strokes = int(strokes)
                pts = hole_points(score_to_par)
                round_hole_pts += pts

                holes_parsed.append({
                    'hole': hole.get('period', 0),
                    'strokes': strokes,
                    'score_to_par': score_to_par,
                    'pts': pts,
                })

            # Round bonuses
            round_bonus, bonus_details = calc_round_bonuses(holes_parsed, round_strokes)

            round_num = round_data.get('period', 0)
            player['rounds'].append({
                'round': round_num,
                'strokes': round_strokes,
                'holes': holes_parsed,
                'hole_pts': round_hole_pts,
                'bonus_pts': round_bonus,
            })
            player['total_hole_pts'] += round_hole_pts
            player['total_bonus_pts'] += round_bonus
            player['bonus_details'].extend([f"R{round_num}: {d}" for d in bonus_details])
            player['rounds_completed'] += 1

        # All rounds under 70 bonus (must complete all 4 rounds, all under 70)
        if player['rounds_completed'] == 4:
            all_under_70 = all(r['strokes'] < 70 for r in player['rounds'])
            if all_under_70:
                player['total_bonus_pts'] += 5
                player['bonus_details'].append("all-rounds-under-70 +5")

        players[name] = player

    return players


def score_tournament(players):
    """Calculate final DK scores for all players."""
    # Finish positions
    positions = calculate_finish_positions(players)

    for name, player in players.items():
        pos = positions.get(name)
        if pos:
            player['finish_position'] = pos
            player['finish_pts'] = get_finish_points(pos)
        else:
            player['finish_position'] = None
            player['finish_pts'] = 0

        player['total_dk_pts'] = (
            player['total_hole_pts'] +
            player['total_bonus_pts'] +
            player['finish_pts']
        )

    return players


def normalize_name(name):
    """Normalize player names for matching."""
    import unicodedata
    # Replace Nordic/special characters first
    replacements = {'ø': 'o', 'å': 'a', 'æ': 'ae', 'ö': 'o', 'ä': 'a', 'ü': 'u', 'é': 'e', 'í': 'i', 'á': 'a', 'ú': 'u', 'ñ': 'n'}
    lower = name.lower()
    for src, dst in replacements.items():
        lower = lower.replace(src, dst)
    # Remove remaining accents
    nfkd = unicodedata.normalize('NFKD', lower)
    ascii_name = ''.join(c for c in nfkd if not unicodedata.combining(c))
    return ascii_name.strip()

# Manual overrides for names that differ between sheet and ESPN
NAME_ALIASES = {
    'pongsapak laopakdee': 'fifa laopakdee',
    'rasmus hajgaard': 'rasmus hojgaard',
}


def build_name_lookup(players):
    """Build a lookup dict with normalized names."""
    lookup = {}
    for name in players:
        lookup[normalize_name(name)] = name
    return lookup


def match_player(golfer_name, name_lookup):
    """Find a player in ESPN data matching the golfer name from our roster."""
    norm = normalize_name(golfer_name)
    if norm in name_lookup:
        return name_lookup[norm]

    # Check aliases
    if norm in NAME_ALIASES:
        alias = NAME_ALIASES[norm]
        if alias in name_lookup:
            return name_lookup[alias]

    # Try partial matching (last name + first initial)
    parts = norm.split()
    for espn_norm, espn_name in name_lookup.items():
        espn_parts = espn_norm.split()
        if parts[-1] == espn_parts[-1] and parts[0][0] == espn_parts[0][0]:
            return espn_name

    # Try last name only (for unique last names)
    matches = [(en, ev) for en, ev in name_lookup.items() if en.split()[-1] == parts[-1]]
    if len(matches) == 1:
        return matches[0][1]

    return None


def update_google_sheet(players, name_lookup):
    """Update the Google Sheet with calculated DK scores."""
    gc = get_gspread_client()
    sh = gc.open_by_url(SHEET_URL)
    ws = sh.sheet1

    data = ws.get_all_values()

    # Find all golfer rows and update Player Total (column D) and +/- (column C)
    updates = []
    unmatched = []

    for row_idx, row in enumerate(data):
        if row_idx == 0:  # header row (row 2 is header based on current layout, row 1 is title)
            continue
        if row_idx == 1:  # column headers
            continue

        golfer_name = row[1].strip() if len(row) > 1 else ''
        if not golfer_name:
            continue

        espn_name = match_player(golfer_name, name_lookup)
        if espn_name and espn_name in players:
            player = players[espn_name]
            dk_pts = player['total_dk_pts']
            score = player.get('espn_score', '')

            # Column C = +/- (row is 1-indexed in sheets)
            sheet_row = row_idx + 1
            updates.append({'range': f'C{sheet_row}', 'values': [[score]]})
            # Column D = Player Total (DK points)
            updates.append({'range': f'D{sheet_row}', 'values': [[dk_pts]]})
        elif golfer_name and golfer_name not in ['Entrant', 'Golfers']:
            unmatched.append(golfer_name)

    if updates:
        ws.batch_update(updates, raw=False)
        print(f"Updated {len(updates) // 2} golfer scores in Google Sheet.")

    if unmatched:
        print(f"\nCould not match these golfers to ESPN data:")
        for name in unmatched:
            print(f"  - {name}")


def print_leaderboard(players):
    """Print a console leaderboard."""
    sorted_players = sorted(players.values(), key=lambda p: -p['total_dk_pts'])

    print(f"\n{'='*70}")
    print(f"{'PLAYER':<25} {'POS':>5} {'SCORE':>6} {'HOLE':>6} {'BONUS':>6} {'FINISH':>6} {'TOTAL':>7}")
    print(f"{'='*70}")

    for p in sorted_players[:30]:
        pos = p.get('finish_position', '-')
        pos_str = str(pos) if pos else '-'
        print(f"{p['name']:<25} {pos_str:>5} {p['espn_score']:>6} {p['total_hole_pts']:>6.1f} {p['total_bonus_pts']:>6.1f} {p['finish_pts']:>6.1f} {p['total_dk_pts']:>7.1f}")

    # Team totals
    print(f"\n{'='*70}")
    print("TEAM LEADERBOARD (Top 6 of 9)")
    print(f"{'='*70}")

    team_scores = []
    for team_name, golfers in TEAMS.items():
        golfer_scores = []
        for g in golfers:
            espn_name = match_player(g, build_name_lookup(players))
            if espn_name and espn_name in players:
                golfer_scores.append((g, players[espn_name]['total_dk_pts']))
            else:
                golfer_scores.append((g, 0))

        golfer_scores.sort(key=lambda x: -x[1])
        top6 = golfer_scores[:6]
        total = sum(s for _, s in top6)
        team_scores.append((team_name, total, top6, golfer_scores[6:]))

    team_scores.sort(key=lambda x: -x[1])

    for rank, (team, total, top6, bottom3) in enumerate(team_scores, 1):
        print(f"\n{rank}. {team}: {total:.1f} pts")
        for g, s in top6:
            print(f"   {g:<30} {s:>7.1f}  ✓")
        for g, s in bottom3:
            print(f"   {g:<30} {s:>7.1f}  ✗")


def main():
    import sys

    # Allow testing with a different event ID
    event_id = sys.argv[1] if len(sys.argv) > 1 else None

    print("Fetching ESPN data...")
    espn_data = fetch_espn_data(event_id)
    event_name = espn_data.get('name', 'Unknown')
    print(f"Tournament: {event_name}")

    print("Parsing player data...")
    players = parse_players(espn_data)
    if not players:
        print("No player data available yet.")
        return

    print(f"Found {len(players)} players")

    print("Calculating DK Classic scores...")
    players = score_tournament(players)

    name_lookup = build_name_lookup(players)

    print_leaderboard(players)

    # Update Google Sheet (skip if testing with different event)
    if event_id is None:
        print("\nUpdating Google Sheet...")
        update_google_sheet(players, name_lookup)
        print("Done!")
    else:
        print(f"\n(Test mode with event {event_id} - skipping sheet update)")
        # Offer to update anyway
        print("Run without arguments to update the Masters sheet.")


if __name__ == '__main__':
    main()
