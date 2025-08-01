from datetime import timedelta, datetime, time

def parse_time(value):
    if isinstance(value, str):
        try:
            h, m = map(int, value.strip().split(":"))
            return timedelta(hours=h, minutes=m)
        except:
            return timedelta(0)
    elif isinstance(value, (int, float)):
        return timedelta(hours=value)
    elif isinstance(value, time):
        return timedelta(hours=value.hour, minutes=value.minute)
    elif isinstance(value, datetime):
        return timedelta(hours=value.hour, minutes=value.minute)
    return timedelta(0)

def format_timedelta(td):
    total_seconds = int(td.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    return f"{hours}:{minutes:02d}"