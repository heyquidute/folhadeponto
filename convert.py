from datetime import timedelta
import re

def to_time(valor):
    # Converte string "HH:MM" para objeto datetime.time
    if not valor:
        return None
    try:
        h, m = map(int, re.findall(r'\d+', valor))
        return timedelta(hours=h, minutes=m)
    except:
        return None  
    
def to_float(valor):
    try:
        return float(str(valor).replace(",", "."))
    except:
        return 0