import datetime


TH_MONTHS = (
    'มกราคม',
    'กุมภาพันธ์',
    'มีนาคม',
    'เมษายน',
    'พฤษภาคม',
    'มิถุนายน',
    'กรกฎาคม',
    'สิงหาคม',
    'กันยายน',
    'ตุลาคม',
    'พฤศจิกายน',
    'ธันวาคม',
    )

def get_full_month_th(month:int):
    return TH_MONTHS[month-1]

def get_full_month_th_with_year(month:int, year:int):
    return f"{TH_MONTHS[month-1]} {year+543}"

# month = get_full_month_th(13)
# print(month)
# assert get_full_month_th(2) == 'กุมภาพันธ์'