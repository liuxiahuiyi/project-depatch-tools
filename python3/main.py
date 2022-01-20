from depatch import Depatcher
from datetime import datetime
import os
import re

workdir = os.path.abspath(os.path.dirname(__file__))
timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
tables_dir = os.path.abspath(os.path.join(workdir, '..', 'tables'))
tables = [t for t in os.listdir(tables_dir) if re.match(r'table-\d+', t)]
if len(tables) < 1:
  raise Exception(f'not found source table in ${tables_dir}')
tables.sort(reverse = True)
print(tables)
table_latest = tables[0]
table_new = f'table-{timestamp}.xlsx'
depatch = Depatcher(os.path.join(tables_dir, table_latest), os.path.join(tables_dir, table_new))
depatch.exec()
