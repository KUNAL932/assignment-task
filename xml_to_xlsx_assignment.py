import re
import os
from bs4 import BeautifulSoup
from datetime import datetime
from xlsxwriter import Workbook
from flask import Flask, request, render_template


def get_transaction_type(transaction_type_tag):
    transaction_type = ''
    if not transaction_type_tag:
        return 'NA'
    if re.search(r"Bank|GST", transaction_type_tag.text, re.I):
        transaction_type = 'Other'
    if re.search(r"agst|ref|new\s*ref", transaction_type_tag.text, re.I):
        transaction_type = 'Child'
    else:
        transaction_type = 'Parent'
    return transaction_type


def validate_date(date):
    dated = datetime.strptime(date.text, "%Y%m%d").strftime('%d-%m-%Y')
    if not dated:
        return 'NA'
    return dated


def main(content, filepath):
    final_output = []
    soup = BeautifulSoup(content, 'xml')
    ordered_list = ['Date', 'Transaction Type',	'Vch No.', 'Ref No', 'Ref Type', 'Ref Date',
                    'Debtor', 'Ref Amount', 'Amount', 'Particulars', 'Vch Type', 'Amount Verified']
    wb = Workbook("output.xlsx")
    if not soup:
        return {"info": "No Soup Data Found."}
    request_data = soup.findAll('REQUESTDATA')
    if not request_data:
        return {"info": "No Request Data Found."}
    for data in request_data:
        voucher_data = data.findAll('VOUCHER')
        if not voucher_data:
            continue
        for vouch in voucher_data:
            matches = {}
            date = vouch.find("DATE")
            if date:
                dated = validate_date(date)
            transaction_type_tag = vouch.find("TRANSACTIONTYPE")

            transaction_type = get_transaction_type(transaction_type_tag)
            voucher_number = vouch.find("VOUCHERNUMBER")
            reference_parent = vouch.find("BILLALLOCATIONS.LIST")
            if reference_parent:
                reference_number = reference_parent.find("NAME")
            reference_type = vouch.find("BILLTYPE")
            reference_date = vouch.find("REFERENCEDATE")
            if reference_date:
                reference_dated = validate_date(reference_date)
            debtor = vouch.find("PARTYNAME")
            amount_value = vouch.find("AMOUNT")
            if transaction_type == 'Parent' or transaction_type == 'Other':
                ref_amount = 'NA'
                amount = amount_value.text if amount_value else 'NA'
                amount_verified = 'Yes'
            else:
                ref_amount = amount_value.text if amount_value else 'NA'
                amount = 'NA'
                amount_verified = 'NA'
            particulars = debtor.text if debtor else 'NA'
            vouch_type = vouch['VCHTYPE'] if 'VCHTYPE' in vouch.attrs else 'NA'
            if not re.search(r"Receipt", vouch_type, re.I):
                continue
            matches['Date'] = dated if dated else 'NA'
            matches['Transaction Type'] = transaction_type if transaction_type else 'NA'
            matches['Vch No.'] = voucher_number.text if voucher_number else 'NA'
            matches['Ref No'] = reference_number.text if reference_number else 'NA'
            matches['Ref Type'] = reference_type.text if reference_type else 'NA'
            matches['Ref Date'] = reference_dated if reference_dated else 'NA'
            matches['Debtor'] = debtor.text if debtor else 'NA'
            matches['Amount'] = amount if amount else 'NA'
            matches['Ref Amount'] = ref_amount if ref_amount else 'NA'
            matches['Amount Verified'] = amount_verified if amount_verified else 'NA'
            matches['Vch Type'] = vouch_type if vouch_type else 'NA'
            matches['Particulars'] = particulars if particulars else 'NA'
            final_output.append(matches)
    ws = wb.add_worksheet()
    first_row = 0
    for header in ordered_list:
        col = ordered_list.index(header)
        ws.write(first_row, col, header)
    row = 1
    for output in final_output:
        for _key, _value in output.items():
            col = ordered_list.index(_key)
            ws.write(row, col, _value)
        row += 1
    wb.close()


api = Flask(__name__)

wsgi_app = api.wsgi_app


@api.route('/', methods=['GET', 'POST'])
def download_xlsx_file():
    if request.method == 'POST':
        file = request.files['file']
        file.save(os.path.join('templates', file.filename))
        filepath = file.filename
        infile = open(filepath, "r")
        contents = infile.read()
        main(contents, filepath)
        return render_template('index.html', message={'success'})
    return render_template('index.html', message={'upload'})


if __name__ == '__main__':
    api.run()
