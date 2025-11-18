from datetime import timedelta
import re

def str_para_tempo(valor):
    # Converte string "HH:MM" para objeto datetime.time
    if not valor:
        return None
    try:
        h, m = map(int, re.findall(r'\d+', valor))
        return timedelta(hours=h, minutes=m)
    except:
        return None  