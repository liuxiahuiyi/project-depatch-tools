class MonthConverter:
  def int_to_month(raw):
    if raw < 4 or raw > 15:
      raise Exception(f'illegal int {raw}')
    elif raw != 12:
      return raw % 12
    else:
      return 12
  def month_to_int(month):
    if month >= 4 and month <= 12:
      return month
    elif month < 4 and month >= 1:
      return month % 12 + 12
    else:
      raise Exception(f'illegal month {month}')
def isNullVal(e):
  return e in [None, '']

