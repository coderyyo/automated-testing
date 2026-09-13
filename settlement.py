# -*- coding: utf-8 -*-
"""客户交易表与广告支出对账。"""
from __future__ import annotations

import io
from typing import Any, Dict, List, Optional, Sequence

import pandas as pd

from processor import (
    _decode_bytes,
    extract_asin,
    find_column,
    is_zero,
    load_table,
    normalize_header,
    parse_number,
)

TRANSACTION_TYPE_ALIASES = ['type', 'typ']
TRANSACTION_TOTAL_ALIASES = ['total', 'gesamt']
TRANSFER_MARKERS = ('transfer', 'übertrag', 'uebertrag')
ADS_PORTFOLIO_ALIASES = ['广告组合', 'portfolio', 'portfolioname']
ADS_SPEND_BY_CURRENCY = (
    ('USD', ['支出(usd)', '花费(usd)', 'spend(usd)']),
    ('GBP', ['支出(gbp)', '花费(gbp)', 'spend(gbp)']),
    ('EUR', ['支出(eur)', '花费(eur)', 'spend(eur)']),
)


def is_transfer_type(value: Any) -> bool:
    text = str(value or '').casefold()
    return any(marker in text for marker in TRANSFER_MARKERS)


def _is_transaction_header(line: str) -> bool:
    compact = normalize_header(line.replace(',', ' ').replace('\t', ' ').replace(';', ' '))
    has_type = any(token in compact for token in TRANSACTION_TYPE_ALIASES)
    has_total = any(token in compact for token in TRANSACTION_TOTAL_ALIASES)
    return has_type and has_total


def _transaction_header_index(lines: List[str]) -> int:
    if len(lines) >= 10 and _is_transaction_header(lines[9]):
        return 9
    for index, line in enumerate(lines[:40]):
        if _is_transaction_header(line):
            return index
    if len(lines) >= 10:
        return 9
    return 0


def _read_csv_from_row(text: str, start: int) -> pd.DataFrame:
    lines = text.splitlines()
    if not lines:
        return pd.DataFrame()
    sliced = '\n'.join(lines[start:])
    buffer = io.StringIO(sliced)
    try:
        frame = pd.read_csv(buffer, sep=None, engine='python')
    except Exception:
        buffer = io.StringIO(sliced)
        frame = pd.read_csv(buffer)
    if frame.shape[1] == 1:
        buffer = io.StringIO(sliced)
        frame = pd.read_csv(buffer, sep='\t')
    return frame


def load_transaction_table(raw: bytes, filename: str) -> pd.DataFrame:
    name = (filename or '').lower()
    if name.endswith(('.xlsx', '.xlsm', '.xls')):
        raw_frame = pd.read_excel(io.BytesIO(raw), header=None)
        header_index = 9 if len(raw_frame) >= 10 else 0
        header_text = ' '.join(str(item) for item in raw_frame.iloc[header_index].tolist())
        if not _is_transaction_header(header_text):
            for index in range(min(40, len(raw_frame))):
                row_text = ' '.join(str(item) for item in raw_frame.iloc[index].tolist())
                if _is_transaction_header(row_text):
                    header_index = index
                    break
        header = raw_frame.iloc[header_index].tolist()
        body = raw_frame.iloc[header_index + 1:].copy()
        body.columns = header
        return body.reset_index(drop=True)
    text = _decode_bytes(raw)
    start = _transaction_header_index(text.splitlines())
    return _read_csv_from_row(text, start)


def process_transaction_report(raw: bytes, filename: str) -> Dict[str, Any]:
    frame = load_transaction_table(raw, filename)
    type_col = find_column(frame.columns, TRANSACTION_TYPE_ALIASES)
    total_col = find_column(frame.columns, TRANSACTION_TOTAL_ALIASES)
    missing = []
    if type_col is None:
        missing.append('type/Typ')
    if total_col is None:
        missing.append('total/Gesamt')
    if missing:
        available = [str(col) for col in frame.columns if str(col).strip() and not str(col).lower().startswith('unnamed')]
        raise ValueError(
            '客户交易表缺少列：{0}。当前识别到的表头：{1}'.format(
                '、'.join(missing),
                '、'.join(available) if available else '无',
            )
        )

    work = pd.DataFrame({
        'type': frame[type_col],
        'total': frame[total_col].map(parse_number),
    })
    work = work[~work['type'].map(is_transfer_type)].copy()
    total = round(float(work['total'].sum()), 2)
    return {
        'rows': int(len(work)),
        'total': total,
        'type_column': str(type_col),
        'total_column': str(total_col),
    }


def _find_spend_column(columns: Sequence[Any]) -> Optional[tuple]:
    for currency, aliases in ADS_SPEND_BY_CURRENCY:
        column = find_column(columns, aliases)
        if column is not None:
            return currency, column
    return None


def process_ads_spend_report(raw: bytes, filename: str) -> Dict[str, Any]:
    hints = ['广告组合'] + [aliases[0] for _currency, aliases in ADS_SPEND_BY_CURRENCY]
    frame = load_table(raw, filename, hints)
    portfolio_col = find_column(frame.columns, ADS_PORTFOLIO_ALIASES)
    spend_found = _find_spend_column(frame.columns)
    missing = []
    if portfolio_col is None:
        missing.append('广告组合')
    if spend_found is None:
        missing.append('支出(USD)/支出(GBP)/支出(EUR)')
    if missing:
        available = [str(col) for col in frame.columns if str(col).strip() and not str(col).lower().startswith('unnamed')]
        raise ValueError(
            '广告报表缺少列：{0}。当前识别到的表头：{1}'.format(
                '、'.join(missing),
                '、'.join(available) if available else '无',
            )
        )
    currency, spend_col = spend_found

    work = pd.DataFrame({
        'portfolio': frame[portfolio_col],
        'spend': frame[spend_col].map(parse_number),
    })
    work['asin'] = work['portfolio'].map(extract_asin)
    work = work[work['asin'].notna()].copy()
    work = work[~work['spend'].map(is_zero)].copy()
    if work.empty:
        return {
            'rows': 0,
            'total': 0.0,
            'currency': currency,
            'spend_column': str(spend_col),
        }

    grouped = work.groupby('asin', as_index=False).agg({'spend': 'sum'})
    total = round(float(grouped['spend'].sum()), 2)
    return {
        'rows': int(len(grouped)),
        'total': total,
        'currency': currency,
        'spend_column': str(spend_col),
    }


def summarize_settlement(
    transaction_raw: bytes,
    transaction_name: str,
    ads_raw: bytes,
    ads_name: str,
) -> Dict[str, Any]:
    transaction = process_transaction_report(transaction_raw, transaction_name)
    ads = process_ads_spend_report(ads_raw, ads_name)
    net = round(transaction['total'] - ads['total'], 2)
    return {
        'transaction_total': transaction['total'],
        'transaction_rows': transaction['rows'],
        'transaction_type_column': transaction['type_column'],
        'transaction_total_column': transaction['total_column'],
        'ads_total': ads['total'],
        'ads_rows': ads['rows'],
        'ads_currency': ads['currency'],
        'ads_spend_column': ads['spend_column'],
        'net': net,
    }
