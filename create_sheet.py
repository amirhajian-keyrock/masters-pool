import gspread

gc = gspread.service_account(filename='credentials.json')
sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/1POSuW5FvJw39DYHumKHTIsoi3AmKbpMdRaHWEDkP_so/edit')

teams = {
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

participants = ['Depa', 'Bill', 'Nascar1776', 'KBM', 'Jeff W.', 'Primo', 'Wretchy', 'Wass', 'Benba', 'Goose']
NUM_GOLFERS = 9

# Clean up sheets
all_sheets = sh.worksheets()
ws = all_sheets[0]
ws.update_title('2026 Masters Pool')
for extra in all_sheets[1:]:
    sh.del_worksheet(extra)
ws.clear()

# Calculate layout: 3 header rows + 10 blocks of (9 golfers + 1 blank) = 103 rows
# Columns: A-E (left) + F (spacer) + G-P (10 leaderboard cols) = 16 cols
total_rows = 3 + len(participants) * (NUM_GOLFERS + 1)
total_cols = 6 + len(participants)  # 16
ws.resize(rows=total_rows + 5, cols=total_cols)

# Build grid
grid = [[''] * total_cols for _ in range(total_rows)]

# Row 1: Title + Leaderboard header
grid[0][0] = '2026 Masters Pool'
grid[0][6] = 'Leaderboard'

# Row 3: Column headers
grid[2][0] = 'Entrant'
grid[2][1] = 'Golfers'
grid[2][2] = '"+/-"'
grid[2][3] = 'TOTAL'
grid[2][4] = 'Total'
for i, p in enumerate(participants):
    grid[2][6 + i] = p

# Row 4: Leaderboard scores (will be formulas referencing each entrant's Total)
# We'll set these after building the entrant blocks

# Entrant blocks
block_starts = []
for p_idx, p_name in enumerate(participants):
    start = 3 + p_idx * (NUM_GOLFERS + 1)
    block_starts.append(start)
    grid[start][0] = p_name
    golfers = teams[p_name]
    for g_idx, golfer in enumerate(golfers):
        row = start + g_idx
        grid[row][1] = golfer
    # Total formula in col E (sum of TOTAL/col D for this block)
    excel_first = start + 1
    excel_last = start + NUM_GOLFERS
    grid[start][4] = f'=SUM(D{excel_first}:D{excel_last})'

# Leaderboard row (row 4 = grid index 3): reference each entrant's Total cell
for i, start in enumerate(block_starts):
    excel_row = start + 1
    grid[3][6 + i] = f'=E{excel_row}'

# Leader section (in the leaderboard area, rows 7-8 like the original)
grid[6][8] = 'Leader'  # col I, row 7
grid[7][8] = f'=INDEX(G3:P3,MATCH(MAX(G4:P4),G4:P4,0))'  # winner name
grid[7][9] = f'=MAX(G4:P4)'  # winner score

# Write all data
ws.update(values=grid, range_name=f'A1:P{total_rows}', raw=False)

# --- FORMATTING (matching British Open exactly) ---
batch_requests = []
sid = ws.id

def rgb(r, g, b):
    return {'red': r, 'green': g, 'blue': b}

def cell_fmt(sheet_id, row, col, end_row=None, end_col=None, bg=None, bold=None, font_size=None, h_align=None):
    if end_row is None: end_row = row + 1
    if end_col is None: end_col = col + 1
    fmt = {}
    fields = []
    if bg:
        fmt['backgroundColor'] = bg
        fields.append('userEnteredFormat.backgroundColor')
    tf = {}
    if bold is not None:
        tf['bold'] = bold
        fields.append('userEnteredFormat.textFormat.bold')
    if font_size is not None:
        tf['fontSize'] = font_size
        fields.append('userEnteredFormat.textFormat.fontSize')
    if tf:
        fmt['textFormat'] = tf
    if h_align:
        fmt['horizontalAlignment'] = h_align
        fields.append('userEnteredFormat.horizontalAlignment')
    return {
        'repeatCell': {
            'range': {'sheetId': sheet_id, 'startRowIndex': row, 'endRowIndex': end_row, 'startColumnIndex': col, 'endColumnIndex': end_col},
            'cell': {'userEnteredFormat': fmt},
            'fields': ','.join(fields)
        }
    }

# Title row: A1 and G1 - light blue bg, bold, size 24
title_bg = rgb(0.79, 0.85, 0.97)
batch_requests.append(cell_fmt(sid, 0, 0, 1, 6, bg=title_bg, bold=True, font_size=24))
batch_requests.append(cell_fmt(sid, 0, 6, 1, total_cols, bg=title_bg, bold=True, font_size=23))

# Header row 3: different colors per column group
# A-B: white bg, bold, size 12
batch_requests.append(cell_fmt(sid, 2, 0, 3, 2, bold=True, font_size=12))
# C: light gray bg, bold, size 12
batch_requests.append(cell_fmt(sid, 2, 2, 3, 3, bg=rgb(0.95, 0.95, 0.95), bold=True, font_size=12))
# D: peach bg, bold, size 12
batch_requests.append(cell_fmt(sid, 2, 3, 3, 4, bg=rgb(0.99, 0.90, 0.80), bold=True, font_size=12))
# E: yellow bg, bold, size 12
batch_requests.append(cell_fmt(sid, 2, 4, 3, 5, bg=rgb(1.0, 1.0, 0.0), bold=True, font_size=12))
# G-P: warm tan bg, bold, size 13
batch_requests.append(cell_fmt(sid, 2, 6, 3, total_cols, bg=rgb(1.0, 0.95, 0.80), bold=True, font_size=13))

# Leaderboard scores row 4: bold, size 13
batch_requests.append(cell_fmt(sid, 3, 6, 4, total_cols, bold=True, font_size=13))

# Entrant names: bold, size 20
for start in block_starts:
    batch_requests.append(cell_fmt(sid, start, 0, start + 1, 1, bold=True, font_size=20))

# +/- column (C) for all data rows: light gray bg, bold, size 10
for start in block_starts:
    batch_requests.append(cell_fmt(sid, start, 2, start + NUM_GOLFERS, 3, bg=rgb(0.95, 0.95, 0.95), bold=True, font_size=10))

# TOTAL column (D) for all data rows: bold, size 12
for start in block_starts:
    batch_requests.append(cell_fmt(sid, start, 3, start + NUM_GOLFERS, 4, bold=True, font_size=12))

# Total column (E) for entrant first row: bold, size 12
for start in block_starts:
    batch_requests.append(cell_fmt(sid, start, 4, start + 1, 5, bold=True, font_size=12))

# Leader label: light blue bg, bold, size 15
batch_requests.append(cell_fmt(sid, 6, 8, 7, 9, bg=rgb(0.79, 0.85, 0.97), bold=True, font_size=15))
# Leader name: peach bg, bold, size 15
batch_requests.append(cell_fmt(sid, 7, 8, 8, 9, bg=rgb(0.99, 0.90, 0.80), bold=True, font_size=15))
# Leader score: bold, size 14
batch_requests.append(cell_fmt(sid, 7, 9, 8, 10, bold=True, font_size=14))

# Column widths matching original
col_widths = {0: 238, 1: 205, 2: 39, 3: 100, 4: 100, 5: 100}
for i in range(6, total_cols):
    col_widths[i] = 100

for col_idx, width in col_widths.items():
    batch_requests.append({
        'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': col_idx, 'endIndex': col_idx + 1},
            'properties': {'pixelSize': width},
            'fields': 'pixelSize'
        }
    })

# Center align the +/- and TOTAL columns
batch_requests.append(cell_fmt(sid, 3, 2, total_rows, 5, h_align='CENTER'))

sh.batch_update({'requests': batch_requests})

print('Done! Sheet matches British Open layout.')
