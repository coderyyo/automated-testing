# -*- coding: utf-8 -*-
from __future__ import annotations

import unittest

from processor import parse_number
from settlement import (
    is_transfer_type,
    process_ads_spend_report,
    process_transaction_report,
    summarize_settlement,
)


def _transaction_csv(header, rows, encoding='utf-8-sig'):
    preamble = '\n'.join(['meta line {0}'.format(index) for index in range(1, 10)])
    body = '\n'.join([header] + rows)
    return (preamble + '\n' + body + '\n').encode(encoding)


class TransferFilterTests(unittest.TestCase):
    def test_transfer_and_uebertrag(self):
        self.assertTrue(is_transfer_type('Transfer'))
        self.assertTrue(is_transfer_type('Fund Transfer'))
        self.assertTrue(is_transfer_type('Übertrag'))
        self.assertTrue(is_transfer_type('Transferir'))
        self.assertTrue(is_transfer_type('Trasferimento'))
        self.assertTrue(is_transfer_type('Transfert'))
        self.assertFalse(is_transfer_type('Order'))
        self.assertFalse(is_transfer_type('Bestellung'))
        self.assertFalse(is_transfer_type('Pedido'))
        self.assertFalse(is_transfer_type('Ordine'))


class ParseNumberLocaleTests(unittest.TestCase):
    def test_german_and_english_money(self):
        self.assertEqual(parse_number('1.234,56'), 1234.56)
        self.assertEqual(parse_number('€12,50'), 12.5)
        self.assertEqual(parse_number('£1,234.50'), 1234.50)


class TransactionReportTests(unittest.TestCase):
    def test_us_header_on_row_10_drops_transfer_and_sums_total(self):
        raw = _transaction_csv(
            'date/time,settlement id,type,order id,total',
            [
                'd,1,Order,A,10.00',
                'd,1,Transfer,B,99.00',
                'd,1,Refund,C,-2.50',
            ],
        )
        result = process_transaction_report(raw, 'us.csv')
        self.assertEqual(result['rows'], 2)
        self.assertEqual(result['total'], 7.5)

    def test_german_typ_gesamt_drops_uebertrag(self):
        raw = _transaction_csv(
            'Datum,Abrechnungsnummer,Typ,Bestellnummer,Gesamt',
            [
                'd,1,Bestellung,A,"1.234,56"',
                'd,1,Übertrag,B,"100,00"',
                'd,1,Erstattung,C,"-10,00"',
            ],
        )
        result = process_transaction_report(raw, 'de.csv')
        self.assertEqual(result['rows'], 2)
        self.assertEqual(result['total'], 1224.56)

    def test_spain_tipo_total_drops_transferir(self):
        raw = _transaction_csv(
            'fecha,tipo,pedido,total',
            [
                'd,Pedido,A,10.00',
                'd,Transferir,B,99.00',
                'd,Reembolso,C,-1.50',
            ],
        )
        result = process_transaction_report(raw, 'es.csv')
        self.assertEqual(result['rows'], 2)
        self.assertEqual(result['total'], 8.5)

    def test_italy_tipo_totale_drops_trasferimento(self):
        raw = _transaction_csv(
            'data,Tipo,ordine,totale',
            [
                'd,Ordine,A,10.00',
                'd,Trasferimento,B,99.00',
                'd,Rimborso,C,-1.50',
            ],
        )
        result = process_transaction_report(raw, 'it.csv')
        self.assertEqual(result['rows'], 2)
        self.assertEqual(result['total'], 8.5)

    def test_france_type_total_drops_transfert(self):
        raw = _transaction_csv(
            'date,type,commande,total',
            [
                'd,Commande,A,10.00',
                'd,Transfert,B,99.00',
                'd,Remboursement,C,-1.50',
            ],
        )
        result = process_transaction_report(raw, 'fr.csv')
        self.assertEqual(result['rows'], 2)
        self.assertEqual(result['total'], 8.5)

    def test_canada_type_total_drops_transfer(self):
        raw = _transaction_csv(
            'date/time,type,order id,total',
            [
                'd,Order,A,10.00',
                'd,Transfer,B,99.00',
                'd,Refund,C,-1.50',
            ],
        )
        result = process_transaction_report(raw, 'ca.csv')
        self.assertEqual(result['rows'], 2)
        self.assertEqual(result['total'], 8.5)


class AdsSpendReportTests(unittest.TestCase):
    def test_gbp_extract_filter_merge_and_sum(self):
        csv_text = (
            '广告组合,支出(GBP)\n'
            'B0F5WJBX67 7件套,2.5\n'
            'B0F5WJBX67 备用,1.5\n'
            '春季促销无编码,9\n'
            'B0CCC00003,0\n'
            'B0DDD00004 套装,4\n'
        ).encode('utf-8-sig')
        result = process_ads_spend_report(csv_text, 'ads-uk.csv')
        self.assertEqual(result['currency'], 'GBP')
        self.assertEqual(result['rows'], 2)
        self.assertEqual(result['total'], 8.0)


class SettlementSummaryTests(unittest.TestCase):
    def test_net_is_transaction_minus_ads(self):
        transaction = _transaction_csv(
            'type,total',
            ['Order,20', 'Transfer,8'],
        )
        ads = '广告组合,支出(USD)\nB0F5WJBX67 7件套,5\n'.encode('utf-8-sig')
        result = summarize_settlement(transaction, 'tx.csv', ads, 'ads.csv')
        self.assertEqual(result['transaction_total'], 20.0)
        self.assertEqual(result['ads_total'], 5.0)
        self.assertEqual(result['net'], 15.0)
        self.assertEqual(result['ads_currency'], 'USD')

    def test_transaction_only(self):
        transaction = _transaction_csv('type,total', ['Order,20', 'Transfer,8'])
        result = summarize_settlement(transaction, 'tx.csv', None, '')
        self.assertTrue(result['has_transaction'])
        self.assertFalse(result['has_ads'])
        self.assertEqual(result['transaction_total'], 20.0)
        self.assertIsNone(result['ads_total'])
        self.assertIsNone(result['net'])

    def test_ads_only(self):
        ads = '广告组合,支出(CAD)\nB0F5WJBX67 7件套,5\n'.encode('utf-8-sig')
        result = summarize_settlement(None, '', ads, 'ads.csv')
        self.assertFalse(result['has_transaction'])
        self.assertTrue(result['has_ads'])
        self.assertEqual(result['ads_total'], 5.0)
        self.assertEqual(result['ads_currency'], 'CAD')
        self.assertIsNone(result['net'])

    def test_requires_at_least_one_file(self):
        with self.assertRaisesRegex(ValueError, '至少'):
            summarize_settlement(None, '', None, '')


if __name__ == '__main__':
    unittest.main()
