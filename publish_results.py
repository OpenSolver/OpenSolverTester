from secrets import *

import datetime
import gspread
import pandas as pd
import platform
import sys

SUMMARY_KEY = '1uHx6UECKT4gz-JImYJ2-byyTuBiI8JRRavLQnI-w_3Q'
RESULTS_KEY = '1YY8ws2sGgcmoXEwNg6V64b6fPKzXQpdpBGERwCJy16w'


def getofficestring(versionnumber):
  release = "Unknown"
  if versionnumber >= 15:
    release = "Office 2013"
  elif versionnumber >= 14.4:
    release = "Office 2011"
  elif versionnumber >= 14:
    release = "Office 2010"
  elif versionnumber >= 12:
    release = "Office 2007"
  return str(versionnumber), release


# https://stackoverflow.com/a/26885636/4492726
def wid_to_gid(wid):
  wid = wid[1:]
  return int(wid, 36) ^ 474

filename = sys.argv[1]

# Login to Google and open Sheet
gc = gspread.login(USERNAME, PASSWORD)
sum_sht = gc.open_by_key(SUMMARY_KEY)
summary_sheet = sum_sht.worksheet('All Results')

# Read in results
linesep = '\n' if platform.system() == 'Windows' else '\r'
results = pd.read_csv(filename, header=0, skiprows=1, na_values=[''],
                      keep_default_na=False, lineterminator=linesep)
results = results.dropna(axis=1, how='all')
row_end = len(results) + 1
col_end = len(results.columns)

# Add results sheet
time = datetime.datetime.now()
res_sht = gc.open_by_key(RESULTS_KEY)
results_sheet = res_sht.add_worksheet(title=str(time), rows=row_end,
                                      cols=col_end)

# Copy results into the new sheet
new_range = 'A1:%s%d' % (chr(col_end + ord('A') - 1), row_end)

cell_list = results_sheet.range(new_range)
for i in range(col_end):
  cell_list[i].value = list(results)[i]
for cell in cell_list[col_end:]:
  cell.value = results.iloc[cell.row - 2, cell.col - 1]
results_sheet.update_cells(cell_list)

# Get data for summary sheet
if platform.system() == "Windows":
  os_release = '%s %s' % (platform.system(), platform.release())
  os_version = platform.version()
else:
  os_release = 'OS X'
  os_version = platform.mac_ver()[0]
os_bitness = str(64 if platform.machine().endswith('64') else 32)

f = open(filename, 'rU')
values = f.readline().rstrip().split(',')
f.close()
office_version, office_release = getofficestring(float(values[0]))
office_bitness = str(values[1])
opensolver_version = str(values[2])

results_wid = results_sheet.id
results_gid = wid_to_gid(results_wid)
results_link = ('https://docs.google.com/spreadsheets/d/%s/edit#gid=%d'
                % (RESULTS_KEY, results_gid))

# Add line to summary sheet
summary_row = 2
while summary_sheet.cell(summary_row, 1).value:
  summary_row += 1

summary_range = 'A%d:R%d' % (summary_row, summary_row)
summary_cells = summary_sheet.range(summary_range)

summary_cells[0].value = time
summary_cells[1].value = os_release
summary_cells[2].value = os_version
summary_cells[3].value = os_bitness
summary_cells[4].value = office_release
summary_cells[5].value = '\'' + office_version
summary_cells[6].value = office_bitness
summary_cells[7].value = opensolver_version
summary_cells[8].value = '=HYPERLINK("%s","LINK")' % results_link

all_pass = 0
all_total = 0
for i, solver in enumerate(['CBC', 'Gurobi', 'NeosCBC', 'NeosBon', 'NeosCou',
                            'NOMAD', 'Bonmin', 'Couenne']):
  if solver in list(results):
    counts = results[solver].value_counts()
    num_pass = counts.get('PASS', 0) + counts.get('PASS*', 0)
    num_total = row_end - 1 - counts.get('NA', 0)
    summary_cells[i + 10].value = (float(num_pass) / num_total if num_total
                                   else '-')
    all_pass += num_pass
    all_total += num_total
  else:
    summary_cells[i + 10].value = '-'

summary_cells[9].value = float(all_pass) / all_total

summary_sheet.update_cells(summary_cells)
