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

    # format prices as prices
    style = xlwt.XFStyle()
    style.num_format_str = '"$"#,##0.00_);("$"#,##'

    for i, (instance, size, price) in enumerate(aws_prices, start=1):
        ws.row(i).write(0, instance)
        ws.row(i).write(1, size)
        ws.row(i).write(2, float(price), style=style)
        hourly_cell = rowcol_to_cell(i, 2)
        hours_in_month = 24 * 30
        ws.row(i).write(3, xlwt.Formula("%s * %s" % (hourly_cell, hours_in_month)), style)

wb = xlwt.Workbook()
ws = wb.add_sheet('EC2 Pricing')
aws_prices = prices()
write_prices(ws, aws_prices)
wb.save('aws.xls')



