from importlib.resources import path
from sys import argv
from sys import exit
import os
from datetime import date
from this import s
import pandas as pd
import re
import xlsxwriter

def get_sales():
    
    if len(argv) >= 2: ##checks for command line peramiter
        get_sales_csv = argv[1]
        if os.path.isfile(get_sales_csv):
            return get_sales_csv
        else:
            print("Error: No file at that location")
            exit()
    else: ##tells if thare was no command line peramiter 
        print("Error: No CSV Path has been provided")
        exit("Plese restart script with path")

def get_order_dir(sales_csv):
    
    ## get directory path of sales data csv
    sales_dir = os.path.dirname(sales_csv)
    
    ## Determine Orders Directory name (Orders_yyy.mm.dd)
    todays_date = date.today().isoformat()
    order_dir_name = "Orders_" + todays_date
    
    ## build the full path of the orders directory
    order_dir = os.path.join(sales_dir, order_dir_name )
   
    ## make the orders directory id it dosent exist
    if not os.path.exists(order_dir):
        os.makedirs(order_dir)
    
    return order_dir

def split_sales_into_orders(sales_csv, order_dir):
    
    ##read data form sales csv to Dataframe
    sales_df = pd.read_csv(sales_csv)

    ##insert new total price column
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df["ITEM PRICE"])

    ##drop unwanted columns
    sales_df.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)

    

    for order_id, order_df in sales_df.groupby('ORDER ID'):
        
        ##drop the order id
        order_df.drop(columns=["ORDER ID"], inplace=True)

        ##sort order by number
        order_df.sort_values(by="ITEM NUMBER", inplace=True)
       
        ##Add grand total at bottom
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE': ['GRAND TOTAL:'], 'TOTAL PRICE':[grand_total]})
        order_df = pd.concat([order_df,grand_total_df])

        ##Determine The Path of the order file
        customer_name = order_df['CUSTOMER NAME'].values[0]
        customer_name = re.sub(r'\W','',customer_name)
        order_file_name = 'Order' + str(order_id) + '_' + customer_name + '.xlsx'
        order_file_path = os.path.join(order_dir, order_file_name)
        sheet_name = 'order#' + str(order_id)
       
        ##format the order and save it
        writer = pd.ExcelWriter(order_file_path, engine='xlsxwriter')
        order_df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        money_fmt = workbook.add_format({'num_format': '$#,##0'})
        worksheet.set_column('A:A', 11)
        worksheet.set_column('B:B', 13)
        worksheet.set_column('C:E', 20)
        worksheet.set_column('F:G', 12,money_fmt)
        worksheet.set_column('H:H', 10)
        worksheet.set_column('I:I', 30)
        writer.save()
        

sales_csv = get_sales()
order_dir = get_order_dir(sales_csv)
split_sales_into_orders (sales_csv,order_dir)