import requests
import xlwt
from xlwt.Utils import rowcol_to_cell

PRICING_URL = "http://aws.amazon.com/ec2/pricing/pricing-on-demand-instances.json"

def prices(region_name="us-east", os="linux", currency="USD"):
    "Retrieve and prase AWS pricing information."
    prices = requests.get(PRICING_URL).json
    region = next((x for x in prices['config']['regions'] if x['region'] == region_name))
    price_tuples = []
    for instance_type in region['instanceTypes']:
        instance_name = instance_type['type']
        for size in instance_type['sizes']:
            size_name = size['size']
            size_price = next((x for x in size['valueColumns'] if x['name'] == os))['prices'][currency]
            price_tuples.append((instance_name, size_name, size_price))
    return price_tuples

def write_prices(ws, aws_prices):
    "Write AWS prices to an Excel worksheet."
    headers = ('Type', 'Size', 'Price/hour (USD)', 'Price/month (USD)')
    for i, header in enumerate(headers):
        ws.row(0).write(i, header)

    # format prices in dollars
    style = xlwt.XFStyle()
    style.num_format_str = '"$"#,##0.00_);("$"#,##'

    for i, (instance, size, price) in enumerate(aws_prices, start=1):
        ws.row(i).write(0, instance)
        ws.row(i).write(1, size)
        ws.row(i).write(2, float(price), style=style)
        hourly_cell = rowcol_to_cell(i, 2)
        hours_in_month = 24 * 30
        ws.row(i).write(3, xlwt.Formula("%s * %s" % (hourly_cell, hours_in_month)), style)

def write_instances(ws, aws_prices):
    "Write columns for recording your AWS instances to Excel worksheet."
    headers = ('Type', 'Size', 'Num Instances', 'Cost/month (USD)')
    for i, header in enumerate(headers):
        ws.row(0).write(i, header)

    # format prices in dollars
    style = xlwt.XFStyle()
    style.num_format_str = '"$"#,##0.00_);("$"#,##'

    for i, (instance, size, _) in enumerate(aws_prices, start=1):
        ws.row(i).write(0, instance)
        ws.row(i).write(1, size)
        ws.row(i).write(2, 0)

        instances_cell = rowcol_to_cell(i, 2)
        monthly_price_cell = rowcol_to_cell(i, 3)
        formula = "'EC2 Pricing'!%s * %s" % (monthly_price_cell, instances_cell)
        ws.row(i).write(3, xlwt.Formula(formula), style)

    # calculate total infrastructure cost
    num_rows = len(ws.rows)
    ws.row(num_rows).write(2, "Total")

    first_price_cell = rowcol_to_cell(1, 3)
    last_price_cell = rowcol_to_cell(num_rows-1, 3)
    total_formula = "SUM(%s:%s)" % (first_price_cell, last_price_cell)
    ws.row(num_rows).write(3, xlwt.Formula(total_formula), style)


wb = xlwt.Workbook()
ws2 = wb.add_sheet('Instances')
ws = wb.add_sheet('EC2 Pricing')

aws_prices = prices()
write_prices(ws, aws_prices)
write_instances(ws2, aws_prices)

wb.save('aws.xls')



