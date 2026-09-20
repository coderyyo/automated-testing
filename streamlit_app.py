# -*- coding: utf-8 -*-
"""亚马逊报表自动化 — Streamlit 入口（可用于 Streamlit Community Cloud）。"""
from __future__ import annotations

import datetime

import streamlit as st

from processor import process_reports
from settlement import summarize_settlement

st.set_page_config(page_title="亚马逊报表自动化", page_icon="📦", layout="centered")


def format_amount(value: float) -> str:
    return '{0:,.2f}'.format(float(value))


def read_upload(upload) -> tuple:
    return upload.getvalue(), upload.name


def render_weekly() -> None:
    st.subheader("周报处理")
    st.write("上传业务报表、广告报表和模板表，清洗 ASIN 后写入模板里的 US 工作表。")
    business = st.file_uploader("1. 业务报表 CSV", type=["csv", "xlsx", "xls"], key="weekly_business")
    ads = st.file_uploader("2. 广告报表 CSV", type=["csv", "xlsx", "xls"], key="weekly_ads")
    template = st.file_uploader("3. 模板 Excel", type=["xlsx", "xlsm"], key="weekly_template")
    if st.button("处理并生成 Excel", type="primary", key="weekly_run"):
        if business is None or ads is None or template is None:
            st.error("请上传全部三个文件")
            return
        try:
            output_bytes, _summary = process_reports(
                *read_upload(business),
                *read_upload(ads),
                *read_upload(template),
            )
        except Exception as exc:
            st.error(str(exc))
            return
        stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        st.session_state['weekly_file'] = {
            'name': 'amazon报表_{0}.xlsx'.format(stamp),
            'data': output_bytes,
        }
        st.success("处理完成，请下载 Excel。")
    weekly_file = st.session_state.get('weekly_file')
    if weekly_file is not None:
        st.download_button(
            "下载 Excel",
            data=weekly_file['data'],
            file_name=weekly_file['name'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="weekly_download",
        )
    st.caption("只写入 US 表。广告组合例如「B0F5WJBX67 7件套」只会留下 B0F5WJBX67。J 列或 O 列里多出来的 ASIN 会黄底标记。")


def render_settlement() -> None:
    st.subheader("交易对账")
    st.write("客户交易表和广告报表可以只传一份，也可以两份一起传。只传一份就只算该份总和；两份都传再算出差额。")
    transaction = st.file_uploader("1. 客户交易 CSV（可选）", type=["csv", "xlsx", "xls"], key="settle_tx")
    ads = st.file_uploader("2. 广告报表 CSV（可选）", type=["csv", "xlsx", "xls"], key="settle_ads")
    if st.button("计算总和", type="primary", key="settle_run"):
        if transaction is None and ads is None:
            st.error("请至少上传一份客户交易表或广告报表")
            return
        try:
            tx_raw, tx_name = (b'', '') if transaction is None else read_upload(transaction)
            ads_raw, ads_name = (b'', '') if ads is None else read_upload(ads)
            summary = summarize_settlement(tx_raw, tx_name, ads_raw, ads_name)
        except Exception as exc:
            st.error(str(exc))
            return
        st.session_state['settle_result'] = summary
    result = st.session_state.get('settle_result')
    if result is not None:
        cols = st.columns(3)
        if result.get('has_transaction'):
            cols[0].metric("交易总和", format_amount(result['transaction_total']))
        else:
            cols[0].metric("交易总和", "—")
        if result.get('has_ads'):
            cols[1].metric(
                "广告总和（{0}）".format(result['ads_currency']),
                format_amount(result['ads_total']),
            )
        else:
            cols[1].metric("广告总和", "—")
        if result.get('net') is not None:
            cols[2].metric("交易 − 广告", format_amount(result['net']))
        else:
            cols[2].metric("交易 − 广告", "—")
    st.caption(
        "交易表默认第 10 行作表头。美/加/法：type + total（Transfer / Transfert）；"
        "德：Typ + Gesamt（Übertrag）；西：tipo + total（Transferir）；"
        "意：Tipo + totale（Trasferimento）。广告支出支持 USD/CAD/GBP/EUR。"
    )


st.title("亚马逊报表自动化")
st.write("两个工具共用同一套处理规则。部署到 Streamlit Community Cloud 后，外网可以直接打开。")
weekly_tab, settle_tab = st.tabs(["周报处理", "交易对账"])
with weekly_tab:
    render_weekly()
with settle_tab:
    render_settlement()
