class Day:
    def __init__(self, busy_hours, all_day_events):
        self.busy_hours = busy_hours
        self.all_day_events = all_day_events
    
    def __str__(self):
        return f"{self.busy_hours.total_seconds() / 3600} busy hours & {self.all_day_events} all day events"