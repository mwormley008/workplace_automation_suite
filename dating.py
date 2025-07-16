import calendar
from datetime import date

today = date.today()
print(today)
res = calendar.monthrange(today.year, today.month)[1]
payment_date = f"{today.month:02d}/{res:02d}/{today.year}"
