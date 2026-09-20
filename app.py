# -*- coding: utf-8 -*-
"""亚马逊报表自动化处理网页。"""
from __future__ import annotations

import datetime
import io
import os

from flask import Flask, render_template, request, send_file

from processor import process_reports
from settlement import summarize_settlement

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024


def io_bytes(data: bytes):
    buffer = io.BytesIO(data)
    buffer.seek(0)
    return buffer


def format_amount(value: float) -> str:
    return '{0:,.2f}'.format(float(value))


@app.route('/')
def home():
    return render_template('home.html', nav='home')


@app.route('/weekly', methods=['GET', 'POST'])
def weekly():
    error = None
    if request.method == 'POST':
        business = request.files.get('business')
        ads = request.files.get('ads')
        template = request.files.get('template')
        if business is None or business.filename in (None, ''):
            error = '请上传第一份业务报表 CSV'
        elif ads is None or ads.filename in (None, ''):
            error = '请上传第二份广告报表 CSV'
        elif template is None or template.filename in (None, ''):
            error = '请上传第三份模板 Excel'
        else:
            try:
                output_bytes, _summary = process_reports(
                    business.read(),
                    business.filename or 'business.csv',
                    ads.read(),
                    ads.filename or 'ads.csv',
                    template.read(),
                    template.filename or 'template.xlsx',
                )
                stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = 'amazon报表_{0}.xlsx'.format(stamp)
                return send_file(
                    io_bytes(output_bytes),
                    as_attachment=True,
                    download_name=filename,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )
            except Exception as exc:
                error = str(exc)
    return render_template('weekly.html', error=error, nav='weekly')


@app.route('/settlement', methods=['GET', 'POST'])
def settlement():
    error = None
    result = None
    if request.method == 'POST':
        transaction = request.files.get('transaction')
        ads = request.files.get('ads')
        tx_name = '' if transaction is None else (transaction.filename or '')
        ads_name = '' if ads is None else (ads.filename or '')
        tx_raw = b'' if transaction is None or tx_name == '' else transaction.read()
        ads_raw = b'' if ads is None or ads_name == '' else ads.read()
        if tx_name == '' and ads_name == '':
            error = '请至少上传一份客户交易表或广告报表'
        else:
            try:
                summary = summarize_settlement(tx_raw, tx_name, ads_raw, ads_name)
                result = {
                    'has_transaction': summary['has_transaction'],
                    'has_ads': summary['has_ads'],
                    'transaction_total': None if summary['transaction_total'] is None else format_amount(summary['transaction_total']),
                    'ads_total': None if summary['ads_total'] is None else format_amount(summary['ads_total']),
                    'net': None if summary['net'] is None else format_amount(summary['net']),
                    'ads_currency': summary['ads_currency'],
                }
            except Exception as exc:
                error = str(exc)
    return render_template('settlement.html', error=error, result=result, nav='settlement')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', '5055'))
    app.run(host='0.0.0.0', port=port, debug=False)
