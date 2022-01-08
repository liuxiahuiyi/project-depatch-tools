class MonthConverter:
  def int_to_month(raw):
    return raw % 12 if raw != 12 else 12
  def month_to_int(month):
    if month >= 4 and month <= 12:
      return month
    else:
      return month % 12 + 12

def isNullVal(e):
  return e in [None, '']

