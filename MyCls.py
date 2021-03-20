class Day:
    def __init__(self, busy_hours):
        self.busy_hours = busy_hours
    
    def __str__(self):
        return f"{self.busy_hours.total_seconds() / 3600}"