from openpyxl import load_workbook, Workbook
import os
from excel_utils import readrow, check_file_exist, get_header_by_row, overwrite_excel, map_header_value, delete_sheet
from typing import Dict
from config import Path, MasterField, TransactionOutField
import datetime
from date_utils import TH_MONTHS, get_full_month_th_with_year


class Master:

    def __init__(self, product_code: str, product_name: str, unit: str, price_per_unit: float | int, inventory_level: float | int) -> None:
        self.product_code = product_code
        self.product_name = product_name
        self.unit = unit
        self.price_per_unit = price_per_unit
        self.inventory_level = inventory_level


class TransactionOut:

    # product_code = 'product_code'
    # product_name = 'product_name'

    def __init__(self, product_code: str, product_name: str, unit: str, price_per_unit: float, qty: int, vat: float, excluding_vat: float, including_vat: float) -> None:
        self.product_code = (product_code, TransactionOutField.PRODUCT_CODE)
        self.product_name = (product_name, TransactionOutField.PRODUCT_NAME)
        self.unit = (unit, TransactionOutField.UNIT)
        self.price_per_unit = [price_per_unit, TransactionOutField.PRICE_PER_UNIT]
        self.qty = [qty, TransactionOutField.QTY]
        self.vat = [vat, TransactionOutField.VAT]
        self.excluding_vat = [excluding_vat, TransactionOutField.EXCLUDING_VAT]
        self.including_vat = [including_vat, TransactionOutField.INCLUDING_VAT]

    def __repr__(self) -> str:
        return f"code:{self.product_code} >>>> qty:{self.qty}"

    def __str__(self) -> str:
        return self.__repr__()


class YearlyReport:

    def __init__(self, save_file:str) -> None:
        self.save_file = save_file
        self.map_summary = {
            TransactionOutField.VAT: 0,
            TransactionOutField.EXCLUDING_VAT: 0,
            TransactionOutField.INCLUDING_VAT: 0,
        }
        self.headers = TransactionOutField.MAP_HEADERS
        self.last_row = 1


    def initial_workbook(self):
        self.wb = Workbook()

    def set_constant_cells(self, date: datetime.datetime):

        sheetname = date.strftime('%b %Y')
        if sheetname not in self.wb.sheetnames:
            self.ws = self.wb.create_sheet(title=sheetname)
        else:
            self.ws = self.wb.active
            self.ws.title = sheetname


        # set header
        self.ws['A2'] = f'รายงานสรุปรายการขายสินค้า ประจำเดือน {get_full_month_th_with_year(month=date.month, year=date.year)}'

        # set fields
        for at_column, header in enumerate(self.headers, start=1):
            self.ws.cell(row=4, column=at_column).value = header
        
        self.last_row = 4

    def set_data(self, map_master, map_transaction_out):
        sorted_product_code = sorted(list(map_master.keys()))
        # set data
        for at_row, product_code in enumerate(sorted_product_code, start=5):
            # get TransactionOut
            self.last_row += 1
            trx = map_transaction_out[product_code]
            for at_column, header_mapper in enumerate(self.headers.values(), start=1):
                value = map_header_value(object_attributes=trx.__dict__, header_mapper=header_mapper)
                self.ws.cell(row=at_row, column=at_column).value = value

                # increase value
                if header_mapper in self.map_summary:
                    self.map_summary[header_mapper] += value
    
    def set_summary(self):
        # set summary
        summary_row = self.last_row + 2
        
        self.ws.cell(row=summary_row, column=2).value = 'รวม'
        self.ws.cell(row=summary_row, column=5).value = f'=SUM(E5:E{self.last_row})'
        self.ws.cell(row=summary_row, column=6).value = f'=SUM(F5:F{self.last_row})'
        self.ws.cell(row=summary_row, column=7).value = f'=SUM(G5:G{self.last_row})'

    def remove_sheet(self):
        # sheet = 'Sheet'
        # if sheet in self.wb.sheetnames:
        #     self.wb.remove_sheet(sheet)
        delete_sheet(wb=self.wb, sheet='Sheet')

    def save(self):
        overwrite_excel(wb=self.wb, save_file=self.save_file )
        self.wb.close()
        
        

def read_master(file_master: str = Path.MASTER_FILE):
    """
    read master excel in Path.MASTER_FILE

    params: 
      file_master: str = absolute path file 

    return: map_master: Dict[str, Master]
    """

    map_master: Dict[str, Master] = {}

    # check file exist

    check_file_exist(file=file_master)

    # if not os.path.exists(file_master):
    #   raise Exception(f'Not found file {file_master}')

    # open workbook
    wb = load_workbook(filename=file_master, data_only=True)
    ws = wb[MasterField.MASTER_SHEET]

    # read header
    header = get_header_by_row(ws=ws, row_index=1)
    # header = []
    # for row in ws.iter_rows(min_row=1, max_row=1):
    #     header = readrow(row)

    # print(header)

    for row in ws.iter_rows(min_row=2):
        row = readrow(row)

        product_code = row[header.index(MasterField.PRODUCT_CODE)]
        master = Master(
            product_code=row[header.index(MasterField.PRODUCT_CODE)],
            product_name=row[header.index(MasterField.PRODUCT_CODE)],
            unit=row[header.index(MasterField.UNIT)],
            price_per_unit=row[header.index('Price/unit')],  # hard code
            inventory_level=row[4]
        )
        if product_code not in map_master:
            map_master[product_code] = master

    wb.close()

    return map_master


def read_transaction_out(map_master: dict, date: datetime.datetime, input_folder: str = Path.FOLDER_INPUT):
    """
    read transaction excel in Path.FOLDER_INPUT

    params: 
        map_master: Dict[str, Master] 
        input_folder: str = path of folder

    return: map_transaction_out: Dict[str, TransactionOut]
    """

    map_transaction_out: Dict[str, TransactionOut] = {}

    file_trx_out = os.path.join(input_folder, f"transaction out - {date.strftime('%Y %m')}.xlsx")

    check_file_exist(file=file_trx_out)

    # open workbook
    wb = load_workbook(filename=file_trx_out)
    ws = wb[wb.sheetnames[0]]

    # read header
    header = get_header_by_row(ws=ws, row_index=1)

    for row in ws.iter_rows(min_row=2):
        row = readrow(row)

        product_code = row[header.index(TransactionOutField.PRODUCT_CODE)]
        product_name = row[header.index(TransactionOutField.PRODUCT_NAME)]
        unit = row[header.index(TransactionOutField.UNIT)]
        price_per_unit = row[header.index(TransactionOutField.PRICE_PER_UNIT)]
        qty = row[header.index(TransactionOutField.QTY)]
        vat = row[header.index(TransactionOutField.VAT)]
        excluding_vat = row[header.index(TransactionOutField.EXCLUDING_VAT)]
        including_vat = row[header.index(TransactionOutField.INCLUDING_VAT)]

        if product_code not in map_master:
            continue

        if product_code not in map_transaction_out:
            map_transaction_out[product_code] = TransactionOut(
                product_code=product_code,
                product_name=product_name,
                unit=unit,
                price_per_unit=price_per_unit,
                qty=qty,
                vat=vat,
                excluding_vat=excluding_vat,
                including_vat=including_vat
            )

        else:
            map_transaction_out[product_code].qty[0] += qty
            map_transaction_out[product_code].vat[0] += vat
            map_transaction_out[product_code].excluding_vat[0] += excluding_vat
            map_transaction_out[product_code].including_vat[0] += including_vat

    # print('map_transaction_out', map_transaction_out)
    wb.close()
    return map_transaction_out


def write_yearly_report(map_transaction_out: Dict[str, TransactionOut], map_master: Dict[str, Master], date:datetime.datetime):

    wb = Workbook()
    ws = wb.active
    # ws.title = "Jan 2022"
    ws.title = date.strftime('%b %Y')

    # set header
    ws['A2'] = f'รายงานสรุปรายการขายสินค้า ประจำเดือน {get_full_month_th_with_year(month=date.month, year=date.year)}'

    headers = TransactionOutField.MAP_HEADERS

    map_summary = {
        TransactionOutField.VAT: 0,
        TransactionOutField.EXCLUDING_VAT: 0,
        TransactionOutField.INCLUDING_VAT: 0,
    }

    for at_column, header in enumerate(headers, start=1):
        ws.cell(row=4, column=at_column).value = header

    sorted_product_code = sorted(list(map_master.keys()))
    last_row = 4
    # set data
    for at_row, product_code in enumerate(sorted_product_code, start=5):
        # get TransactionOut
        last_row += 1
        trx = map_transaction_out[product_code]
        for at_column, header_mapper in enumerate(headers.values(), start=1):
            value = map_header_value(object_attributes=trx.__dict__, header_mapper=header_mapper)

            ws.cell(row=at_row, column=at_column).value = value
            

            # increase value
            if header_mapper in map_summary:
                map_summary[header_mapper] += value


    # set summary
    summary_row = last_row + 2
    
    ws.cell(row=summary_row, column=2).value = 'รวม'
    ws.cell(row=summary_row, column=5).value = f'=SUM(E5:E{last_row})'
    ws.cell(row=summary_row, column=6).value = f'=SUM(F5:F{last_row})'
    ws.cell(row=summary_row, column=7).value = f'=SUM(G5:G{last_row})'


    # save
    save_file=os.path.join(Path.FOLDER_OUTPUT, f'รายงานสรุปรายการขายสินค้าแต่ละเดือน ปี 2022.xlsx')
    overwrite_excel(wb=wb, save_file=save_file )

    wb.close()




def main():

    # Read master
    map_master = read_master()

    # Read transaction out per month

    # write excel on sheet [month name]
    # write_yearly_report(map_transaction_out=map_transaction_out, map_master=map_master)

    save_file=os.path.join(Path.FOLDER_OUTPUT, f'รายงานสรุปรายการขายสินค้าแต่ละเดือน ปี 2022.xlsx')
    
    yearly_report = YearlyReport(save_file=save_file)
    yearly_report.initial_workbook()

    for month in range(1,13):

        date = datetime.datetime(2022,month, 1)
        print('date', date)

        map_transaction_out = read_transaction_out(map_master=map_master, date=date)
        
        yearly_report.set_constant_cells(date=date)
        yearly_report.set_data(map_master=map_master, map_transaction_out=map_transaction_out)
        yearly_report.set_summary()

        # if month == 3 :
        #     break

    yearly_report.remove_sheet()
    yearly_report.save()


if __name__ == '__main__':

    main()
    
    # for month in TH_MONTHS:
    #     print(month)
    # pass
