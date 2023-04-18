import os
from typing import Dict

def cell_value(cell):
    if (cv:=cell.value) is None:
        return ''
    elif type(cv) is str:
        return cv.strip()
    else:
        return cv

def readrow(row: tuple|list):
    return [cell_value(cell) for cell in row]

def check_file_exist(file:str):
    if not os.path.exists(file):
        raise Exception(f'Not found file {file}')
    

def get_header_by_row(ws:object, row_index:int = 1):
    header = []
    for row in ws.iter_rows(min_row=row_index, max_row=row_index):
        header = readrow(row)
    
    return header

def overwrite_excel(wb:object, save_file:str):
    if os.path.exists(save_file):
        os.remove(save_file)
    
    wb.save(save_file)

def delete_sheet(wb:object, sheet:str):
    if sheet in wb.sheetnames:
        wb.remove_sheet(wb[sheet])


def map_header_value(object_attributes: Dict[str, list|tuple], header_mapper:str):

    for _, array in object_attributes.items():
        value, header_name = array
        if header_mapper == header_name:
            return value
    else:
        return ''
{
    'product_code': ('SL36200BL', 'Product code'), 
    'product_name': ('เคเบิ้ลไทร์ Porlock SL36200BL ขนาด 8 นิ้ว สีดำ', 'Product name'), 
    'unit': ('Pack', 'Unit'), 
    'price_per_unit': [65, 'Price/unit'], 
    'qty': [1666, 'Qty'], 
    'vat': [7580.3, 'VAT'], 
    'excluding_vat': [108290, 'Excluding VAT'], 
    'including_vat': [115870.3, 'Total (including VAT)']
}