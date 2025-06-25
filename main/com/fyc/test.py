from datetime import datetime, timedelta

# 直接计算上周五
last_friday = datetime.today() - timedelta(days=(datetime.today().weekday() + 3) % 7)

# 直接计算上上周五
last_last_friday = last_friday - timedelta(7)

print(f"上周五: {last_friday:%Y-%m-%d}")
print(f"上上周五: {last_last_friday:%Y-%m-%d}")