from depatch import Depatcher
from datetime import datetime
import os
import re

user_dir = os.path.expanduser('~')
workdir = None
for e in os.walk(user_dir):
  if 'project-depatch-tools' in e[1]:
    workdir = os.path.join(e[0], 'project-depatch-tools')
    break
if workdir is None:
  raise Exception(f'Could not found project-depatch-tools in user home')

timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
tables_dir = os.path.abspath(os.path.join(workdir, 'tables'))
fiscal_years = os.listdir(tables_dir)

for y in fiscal_years:
  if not re.match(r'^\d+$', y):
    continue
  tables = [t for t in os.listdir(os.path.join(tables_dir, y)) if re.match(r'table-\d+', t)]
  if len(tables) < 1:
    raise Exception(f'not found source table in ${tables_dir}')
  tables.sort(reverse = True)
  table_latest = tables[0]
  table_new = f'table-{timestamp}.xlsx'
  depatch = Depatcher(os.path.join(tables_dir, y, table_latest), os.path.join(tables_dir, y, table_new), int(y))
  depatch.exec()
