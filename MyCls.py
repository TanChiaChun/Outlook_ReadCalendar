class Day:
    def __init__(self, busy_hours, all_day_events, due, do, start):
        self.busy_hours = busy_hours
        self.all_day_events = all_day_events
        self.due = due
        self.do = do
        self.start = start
    
    def __str__(self):
        return f"{self.busy_hours.total_seconds() / 3600},{self.all_day_events},{self.due},{self.do},{self.start}"

class Appointment:
    def __init__(self, start, end, is_all_day, cat):
        self.start = start
        self.end = end
        self.is_all_day = is_all_day
        self.cat = cat