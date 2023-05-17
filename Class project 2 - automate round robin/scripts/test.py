# from main import validate_income


# result = validate_income(product_tier=4, customer_tier=2)


# result = validate_income(product_tier=4, customer_tier=4)



map_sale = {
    'S': ['S1', 'S2', 'S3'],
    'A': ['A1', 'A2', 'A3'],
    'B': ['B1', 'B2', 'B3'],
    'C': ['C1', 'C2'],
    'D': ['D1', 'D2', 'D3'],
}

map_valid_customer = {
    'S': ['1111', '2222', '3333', '4444', '5555', '6666', '7777'],
    'A': ['aaaa', 'ssss', 'Ad3' , 'Aaaaa4'],
    'B': ['Bsdf1', 'Bsdf2', 'Basdf3'],
    'D': ['DDD111', 'dddd2', 'dd3'],
}




def round_robin(sales: list, valid_customer: list):
    sale_dummy = sales.copy()

    output = []
    # [ [S1, Regis], ... ]
    for data in valid_customer:
        # Customer '1111'
        # sale[0] = S1  >>>>  ['S1', 'S2', 'S3']

        if len(sales) == 0:
            sales = sale_dummy.copy()

        # print('Customer ',data, sales[0])
        output.append([sales[0], data ]) 
            
        if len(sales) > 0:
            sales.pop(0)
        # print()

    return output

if __name__ == '__main__':
    # group = 'A'
    # sales = map_sale[group].copy()
    # valid_data = map_valid_data[group].copy()
    # final_a = round_robin(sales, valid_data)

    group = 'S'
    sales = map_sale[group].copy()
    valid_customer = map_valid_customer[group]

    final_s = round_robin(sales=sales , valid_customer=valid_customer)

    # all_output = []

    # for group in map_sale:

    #     if group not in map_valid_customer:
    #         all_output.append(['C', 'no customer'])
    #         continue
    #     # sales = map_sale[group].copy()
    #     # valid_data = map_valid_data[group].copy()
    #     final_data = round_robin(sales=map_sale[group].copy() , valid_customer=map_valid_customer[group] )
    #     print('final_data',group, final_data)
    #     all_output += final_data


    # print('all_output', all_output)
