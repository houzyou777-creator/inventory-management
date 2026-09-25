# -*- coding: utf-8 -*-
"""pricetar_cost_diff.py のテスト

・分類/紐付け/原価確認記録の判定はダミーデータで確かめる
・E2E は実ファイルを読み取り専用で実行し、正本(商品マスター・原価確認記録・KPI)が
  1バイトも変わらないことをハッシュで確かめる。出力は一時フォルダへ

実行: python3 -m unittest discover -s MomijiStore_OS/02_Analytics/tests -v
"""
import glob
import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'Python'))
import pricetar_cost_diff as D  # noqa: E402


def L(lid, pid, asin='', asku='', ch='Amazon', pack=None, unit=''):
    return {'lid': lid, 'pid': pid, 'ch': ch, 'asin': asin, 'asku': asku,
            'pack': pack, 'unit': unit, 'proof': ''}


def index(listings):
    by_sku, by_asin = {}, {}
    for l in listings:
        if l['asku']:
            by_sku.setdefault(l['asku'], []).append(l)
        if l['asin']:
            by_asin.setdefault(l['asin'], []).append(l)
    return by_sku, by_asin


class Classify(unittest.TestCase):
    def test_match(self):
        self.assertEqual(D.classify(500, 500, None, '')[0], D.CAT_MATCH)

    def test_single_cost_times_pack_is_match(self):
        c, exp, diff, ratio, _ = D.classify(1500, 500, 3, D.UNIT_SINGLE)
        self.assertEqual((c, exp, diff), (D.CAT_MATCH, 1500, 0))

    def test_set_cost_not_multiplied(self):
        # セット原価に入数を掛けると二重になる → 掛けずに比較する
        c, exp, _, _, _ = D.classify(1500, 1500, 3, D.UNIT_SET)
        self.assertEqual((c, exp), (D.CAT_MATCH, 1500))

    def test_integer_multiple_when_unit_unknown(self):
        c, _, diff, ratio, why = D.classify(2082, 694, None, '')
        self.assertEqual(c, D.CAT_MULTI)
        self.assertEqual(ratio, 3)
        self.assertIn('3倍', why)

    def test_inverse_multiple_when_unit_unknown(self):
        c, _, _, _, why = D.classify(761, 2283, None, '')
        self.assertEqual(c, D.CAT_MULTI)
        self.assertIn('1/3', why)

    def test_multiple_is_diff_when_unit_confirmed(self):
        # 単位が確認済みなら倍率で言い逃れしない(確認済み入数と合わなければ差異)
        c, exp, diff, _, _ = D.classify(2082, 694, 2, D.UNIT_SINGLE)
        self.assertEqual((c, exp, diff), (D.CAT_DIFF, 1388, 694))

    def test_single_unit_without_pack_is_not_converted(self):
        c, exp, _, _, why = D.classify(1388, 694, None, D.UNIT_SINGLE)
        self.assertEqual((c, exp), (D.CAT_MULTI, 694))
        self.assertIn('未確認', why)

    def test_small_and_large_diff(self):
        self.assertIn('少額差', D.classify(537, 536, None, '')[4])
        self.assertIn('±15%', D.classify(3940, 3546, None, '')[4])
        self.assertIn('大きな乖離', D.classify(963, 525, None, '')[4])

    def test_missing_costs(self):
        self.assertEqual(D.classify(None, 500, None, '')[0], D.CAT_NO_PC)
        self.assertEqual(D.classify(0, 500, None, '')[0], D.CAT_NO_PC)
        self.assertEqual(D.classify(500, None, None, '')[0], D.CAT_NO_MC)

    def test_zero_master_cost_does_not_crash(self):
        self.assertEqual(D.classify(500, 0, None, '')[0], D.CAT_DIFF)


class Link(unittest.TestCase):
    def test_sku_unique(self):
        ls = [L('C1', 'P1', 'B01', 'SKU1')]
        l, m, c, _ = D.link({'SKU': 'sku1', 'ASIN': 'B01'}, *index(ls), {'SKU1'})
        self.assertEqual((l['lid'], m, c), ('C1', 'SKU', None))

    def test_sku_duplicate_is_ambiguous(self):
        ls = [L('C1', 'P1', 'B01', 'SKU1'), L('C2', 'P2', 'B01', 'SKU1')]
        self.assertEqual(D.link({'SKU': 'SKU1', 'ASIN': 'B01'}, *index(ls), {'SKU1'})[2], D.CAT_AMBIG)

    def test_asin_only_single(self):
        ls = [L('C1', 'P1', 'B01', '')]
        l, m, c, note = D.link({'SKU': 'NEW', 'ASIN': 'B01'}, *index(ls), {'NEW'})
        self.assertEqual((l['lid'], m, c), ('C1', 'ASINのみ', None))
        self.assertIn('SKU変更の可能性', note)

    def test_asin_multiple_is_ambiguous(self):
        ls = [L('C1', 'P1', 'B01', 'A'), L('C2', 'P1', 'B01', 'B')]
        self.assertEqual(D.link({'SKU': 'NEW', 'ASIN': 'B01'}, *index(ls), {'NEW'})[2], D.CAT_AMBIG)

    def test_asin_listing_owned_by_other_csv_row_is_ambiguous(self):
        # FBA と自己発送で同じASIN。出品テーブルにはFBA側のSKUしか無い → 自己発送行を寄せない
        ls = [L('C1', 'P1', 'B01', 'FBA1')]
        r = D.link({'SKU': 'SELF1', 'ASIN': 'B01'}, *index(ls), {'SELF1', 'FBA1'})
        self.assertEqual(r[2], D.CAT_AMBIG)

    def test_no_link(self):
        self.assertEqual(D.link({'SKU': 'X', 'ASIN': 'B09'}, {}, {}, {'X'})[2], D.CAT_NOLINK)

    def test_rakuten_listing_not_used_for_link(self):
        rows = D.build_rows([{'SKU': 'X', 'ASIN': 'B01', 'cost': '100', 'title': '', 'number': '1',
                              'price': '1', '_file': 'a.csv', '_line': 2}],
                            {'P1': {'jan': '', 'name': 'n', 'type': '単品', 'cost': 100}},
                            [L('C1', 'P1', 'B01', '', ch='楽天')], [], '2026-09')
        self.assertEqual(rows[0]['cat'], D.CAT_NOLINK)
        self.assertIn('楽天出品にのみ同ASINあり', rows[0]['reason'])


class Confirmation(unittest.TestCase):
    REC = {'記録ID': 'K1', 'チャネル': 'Amazon', 'SKU/ASIN': 'B01', '適用開始月': '2026-09',
           '適用終了月': '2026-10', '1販売分の原価': 1388, '販売入数': 2, '原価単位': '単品原価'}

    def test_period(self):
        self.assertTrue(D.confirmation_for([self.REC], 'S', 'B01', '2026-10'))
        self.assertFalse(D.confirmation_for([self.REC], 'S', 'B01', '2026-11'))
        self.assertFalse(D.confirmation_for([self.REC], 'S', 'B01', '2026-08'))

    def test_invalidated_or_other_channel(self):
        self.assertFalse(D.confirmation_for([dict(self.REC, 無効化日='2026-09-20')], 'S', 'B01', '2026-09'))
        self.assertFalse(D.confirmation_for([dict(self.REC, チャネル='楽天')], 'S', 'B01', '2026-09'))

    def test_reason_mentions_support(self):
        rows = D.build_rows([{'SKU': 'S', 'ASIN': 'B01', 'cost': '1388', 'title': '', 'number': '1',
                              'price': '1', '_file': 'a.csv', '_line': 2}],
                            {'P1': {'jan': '', 'name': 'n', 'type': '単品', 'cost': 694}},
                            [L('C1', 'P1', 'B01', 'S', pack=2)], [self.REC], '2026-09')
        self.assertIn('裏付けあり', rows[0]['reason'])
        self.assertIn('K1', rows[0]['confirm'])


class CsvRead(unittest.TestCase):
    def test_cp932_and_equal_quote(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, 'x_FBA.csv')
            with open(p, 'wb') as f:
                f.write('SKU,ASIN,number,price,point,cost,condition,title\n'
                        '="pr_1",="B01","5","820","0","273","11","テスト"\n'.encode('cp932'))
            r = D.read_pricetar_csv(p)[0]
            self.assertEqual((r['SKU'], r['ASIN'], r['cost'], r['title']), ('pr_1', 'B01', '273', 'テスト'))

    def test_missing_column_stops(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, 'x.csv')
            with open(p, 'wb') as f:
                f.write('SKU,ASIN,number,price\n'.encode('cp932'))
            with self.assertRaises(SystemExit):
                D.read_pricetar_csv(p)


class EndToEndReadOnly(unittest.TestCase):
    """実ファイルで実行し、正本が変わらないことを確かめる(出力は一時フォルダ)。"""

    def test_real_files_unchanged(self):
        csvs = sorted(glob.glob(D.MP.INV_SD.replace('SourceData', 'import') + '/Amazon/*inventoryList*.csv'))
        if not csvs or not os.path.exists(D.MASTER_FILE):
            self.skipTest('実データが無い環境')
        kpis = glob.glob(D.MP.ANA_SD + '/*KPI管理シート.xlsx')
        guarded = [D.MASTER_FILE, D.CONFIRM_FILE] + kpis + csvs
        before = {p: D.sha256(p) for p in guarded if os.path.exists(p)}
        with tempfile.TemporaryDirectory() as d:
            out = os.path.join(d, 'r.xlsx')
            rows = D.main(csvs + ['--out', out, '--month', '2026-09'])
            self.assertTrue(os.path.exists(out))
        n_csv = sum(len(D.read_pricetar_csv(p)) for p in csvs)
        self.assertEqual(len(rows), n_csv)                     # 行の取りこぼしが無い
        self.assertTrue({r['cat'] for r in rows} <= set(D.CAT_ORDER))
        self.assertEqual({p: D.sha256(p) for p in before}, before)

    def test_refuses_to_overwrite_input(self):
        with self.assertRaises(SystemExit):
            D.main(['dummy.csv', '--out', D.MASTER_FILE])


if __name__ == '__main__':
    unittest.main()
