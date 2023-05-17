import os

class Path:
    BASE_DIR = os.path.dirname(os.path.dirname(__file__))

    MASTER_FILE = os.path.join(BASE_DIR, 'Master.xlsx')

    FOLDER_INPUT = os.path.join(BASE_DIR, 'data', 'input')
    FOLDER_OUTPUT = os.path.join(BASE_DIR, 'data', 'output')

class MasterField:
    # Fields
    PRODUCT_CODE = 'Product code'
    PRODUCT_NAME = 'Product name'
    UNIT = 'Unit'

    MASTER_SHEET = 'Product-Detail'

class TransactionOutField:
    PRODUCT_CODE = 'Product code'
    PRODUCT_NAME = 'Product name'
    UNIT = 'Unit'
    PRICE_PER_UNIT = 'Price/unit'
    VAT = 'VAT'
    EXCLUDING_VAT = 'Excluding VAT'
    INCLUDING_VAT = 'Total (including VAT)'
    QTY = 'Qty'


    MAP_HEADERS = {
        'รหัสสินค้า': PRODUCT_CODE, 
        'ชื่อสินค้า': PRODUCT_NAME, 
        'จำนวนสินค้า': QTY, 
        'ราคาต่อหน่วย': PRICE_PER_UNIT, 
        'VAT': VAT, 
        'Excluding VAT': EXCLUDING_VAT, 
        'Total (Including VAT)': INCLUDING_VAT
    }
  
# print(Path.BASE_DIR)
# print(Path.MASTER_FILE)