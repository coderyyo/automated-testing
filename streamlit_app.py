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
    st.write("上传同一站点的客户交易表和广告报表，展示交易总和、广告总和，以及两者差额。")
    transaction = st.file_uploader("1. 客户交易 CSV", type=["csv", "xlsx", "xls"], key="settle_tx")
    ads = st.file_uploader("2. 广告报表 CSV", type=["csv", "xlsx", "xls"], key="settle_ads")
    if st.button("计算总和", type="primary", key="settle_run"):
        if transaction is None or ads is None:
            st.error("请上传客户交易表和广告报表")
            return
        try:
            summary = summarize_settlement(
                *read_upload(transaction),
                *read_upload(ads),
            )
        except Exception as exc:
            st.error(str(exc))
            return
        st.session_state['settle_result'] = summary
    result = st.session_state.get('settle_result')
    if result is not None:
        col1, col2, col3 = st.columns(3)
        col1.metric("交易总和", format_amount(result['transaction_total']))
        col2.metric("广告总和（{0}）".format(result['ads_currency']), format_amount(result['ads_total']))
        col3.metric("交易 − 广告", format_amount(result['net']))
    st.caption("交易表第 10 行作为表头，只保留 type/Typ 和 total/Gesamt；含 Transfer 或 Übertrag 的行会去掉。请不要把美元、英镑、欧元混在一次计算里。")


st.title("亚马逊报表自动化")
st.write("两个工具共用同一套处理规则。部署到 Streamlit Community Cloud 后，外网可以直接打开。")
weekly_tab, settle_tab = st.tabs(["周报处理", "交易对账"])
with weekly_tab:
    render_weekly()
with settle_tab:
    render_settlement()
