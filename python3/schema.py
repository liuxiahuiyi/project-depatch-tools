from util import MonthConverter
from util import isNullVal
from datetime import datetime

class DimProject:
  def __init__(self, name, status, ma_or_project, category, budget, month_start, month_end, **budget_by_month):
    self.current_time = MonthConverter.month_to_int(int(datetime.now().strftime('%m')))
    self.name = name
    self.status = status
    self.ma_or_project = ma_or_project
    self.category = category
    self.budget = float(budget)
    self.time_start = MonthConverter.month_to_int(int(month_start))
    self.time_end = MonthConverter.month_to_int(int(month_end))
    self.allocateBudget(**budget_by_month)
  def allocateBudget(self, budget_by_month):
    count = self.time_end - self.time_start + 1
    if sum([float(budget_by_month[f'budget_by_month_{MonthConverter.int_to_month(i)}']) if f'budget_by_month_{MonthConverter.int_to_month(i)}' in budget_by_month and not isNullVal(budget_by_month[f'budget_by_month_{MonthConverter.int_to_month(i)}']) and float(budget_by_month[f'budget_by_month_{MonthConverter.int_to_month(i)}']) >= 0 else -1.0 for i in range(self.time_start, self.time_end + 1)]) == self.budget:
      self.budget_by_time = dict([(i, float(budget_by_month[f'budget_by_month_{MonthConverter.int_to_month(i)}']) if i >= self.time_start and i <= self.time_end else None) for i in range(4, 16)])
    else:
      if count == 1:
        percentages = [1]
      elif count > 1:
        percentages = [(1 / count) - 0.05 * (count - 2) / count if i == self.time_start or i == self.time_end else (1 / count) + 0.05 * 2 / count for i in range(self.time_start, self.time_end + 1)]
      else:
        raise Exception(f'illegal self time_start {self.time_start} and time_end {self.time_end}')
      self.budget_by_time = dict([(i, self.budget * percentages[i - self.time_start] if i >= self.time_start and i <= self.time_end else None) for i in range(4, 16)])
    self.calRemain()
  def calRemain(self):
    self.remain = self.budget - sum([v for k, v in self.budget_by_time.items() if k < self.current_time and v is not None])
  def updateBudgetByTime(self, time, budget_update):
    if time < self.time_start or time > self.time_end or time >= self.current_time:
      return
    diff = budget_update - self.budget_by_time[time]
    self.budget_by_time[time] = budget_update
    update_times = self.time_end - self.current_time + 1
    if update_times <= 0:
      return
    diff_per_time = diff / update_times
    for time in range(self.current_time, self.time_end + 1):
      self.budget_by_time[time] = self.budget_by_time[time] - diff_per_time
    self.calRemain()


class DimEmployee:
  def __init__(self, itcode, category, role, **rate_and_md):
    self.itcode = itcode
    self.category = category
    self.role = role
    self.readRate(**rate_and_md)
    self.readMd(**rate_and_md)
  def readRate(self, rate_and_md):
    self.rate = {}
    for i in range(4, 16):
      key = f'rate_{MonthConverter.int_to_month(i)}'
      if key in rate_and_md and not isNullVal(rate_and_md[key]):
        self.rate[i] = float(rate_and_md[key])
      else:
        self.rate[i] = None
  def readMd(self, rate_and_md):
    self.md = {}
    for i in range(4, 16):
      key = f'md_{MonthConverter.int_to_month(i)}'
      if key in rate_and_md and not isNullVal(rate_and_md[key]):
        self.md[i] = float(rate_and_md[key])
      else:
        self.md[i] = None
