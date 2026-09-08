#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银行对账核对工具
基于《银行对账任务规范.md》—— 当前唯一生效标准
既往所有临时方案、被淘汰规则全部作废。

【图形界面用法】(双击运行 / 直接 python main.py)
  打开简易 GUI，选择"对账单文件"和"日记账文件"两个 Excel，
  点击[开始核对]，自动生成 7-Sheet 对账结果文件。

【命令行用法】(保留，便于批量/调试)
  python main.py <对账单文件> <日记账文件> [输出文件]        # 单对核对
  python main.py 对账单1 日记账1 对账单2 日记账2 ...         # 一次多对(文件两两一组)
  python main.py <合并文件(含Sheet1+银行明细账)> [输出文件]

"""

import sys
import os
import re
import threading
import queue
import subprocess
from datetime import datetime, timedelta, date
from collections import defaultdict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Border, Side, Alignment

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from tkinter.scrolledtext import ScrolledText
    _TK_OK = True
except Exception:  # 无图形环境（如服务器）时仍可正常使用命令行模式
    _TK_OK = False

# ============================================================
# 样式常量 (Section 8)
# 说明: 输出结果一律不加底色(表头/自核对/未核对均无填充)；
#       组标注红框为细框，仅多笔组使用。
# ============================================================
HEADER_FONT = Font(name='微软雅黑', size=10, bold=True)
DATA_FONT   = Font(name='微软雅黑', size=10)
THIN_SIDE  = Side(style='thin', color='B0B0B0')
GROUP_RED  = Side(style='medium', color='FF0000')   # 多笔核对组的红框标注(中粗)
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
NUM_FMT = '#,##0.00'
CENTER_AL = Alignment(horizontal='center', vertical='center')
RIGHT_AL  = Alignment(horizontal='right', vertical='center')
LEFT_AL   = Alignment(horizontal='left', vertical='center')


# ============================================================
# Section 2: 数据读取
# ============================================================

def safe_float(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (datetime, date)):
        return 0.0
    try:
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip()
        s = s.replace(',', '').replace('$', '')
        s = s.replace('O', '0').replace('o', '0')
        s = s.replace('(', '-').replace(')', '')
        if s == '' or s == '-' or s == '--':
            return 0.0
        return float(s)
    except Exception:
        return 0.0


def parse_date(v):
    if v is None or v == '':
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    try:
        s = str(v).strip()
        # 处理 float 类型日期 (如 20220101.0)
        if isinstance(v, float):
            s = str(int(v))
        for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y%m%d', '%m/%d/%Y', '%m-%d-%Y'):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                pass
        try:
            n = float(s)
            if 1 < n < 100000:
                return datetime(1899, 12, 30) + timedelta(days=n)
        except ValueError:
            pass
        return None
    except Exception:
        return None


def remove_spaces(s):
    if s is None:
        return ''
    return str(s).replace(' ', '').replace('\u3000', '').strip()


def normalize_cp(s):
    return remove_spaces(s).lower()


def read_statements(ws):
    records = []
    # 用 iter_rows 顺序流式读取（read_only 模式下 ws.cell() 随机访问极慢，禁止使用）
    for row_idx, row in enumerate(ws.iter_rows(min_row=4, max_row=ws.max_row,
                                               max_col=7, values_only=True), start=4):
        d = parse_date(row[0] if len(row) > 0 else None)
        if d is None:
            continue
        debit  = safe_float(row[1] if len(row) > 1 else None)
        credit = safe_float(row[2] if len(row) > 2 else None)
        balance = safe_float(row[3] if len(row) > 3 else None)
        cp  = row[4] if len(row) > 4 else None
        cpa = row[5] if len(row) > 5 else None
        sm  = row[6] if len(row) > 6 else None
        rec = {
            'idx': len(records), 'date': d,
            'debit': debit, 'credit': credit, 'balance': balance,
            'counterparty': str(cp) if cp else '',
            'counterparty_acct': str(cpa) if cpa else '',
            'summary': str(sm) if sm else '',
            'source_row': row_idx,
            'matched': False, 'match_id': None, 'match_type': None,
            'amount': max(debit, credit),
        }
        records.append(rec)
    return records


def read_ledger(ws):
    records = []
    # 用 iter_rows 顺序流式读取（read_only 模式下 ws.cell() 随机访问极慢，禁止使用）
    for row_idx, row in enumerate(ws.iter_rows(min_row=4, max_row=ws.max_row,
                                               max_col=8, values_only=True), start=4):
        d = parse_date(row[0] if len(row) > 0 else None)
        if d is None:
            continue
        vtype = row[1] if len(row) > 1 else None
        vno   = row[2] if len(row) > 2 else None
        debit  = safe_float(row[3] if len(row) > 3 else None)
        credit = safe_float(row[4] if len(row) > 4 else None)
        balance = safe_float(row[5] if len(row) > 5 else None)
        cp = row[6] if len(row) > 6 else None
        sm = row[7] if len(row) > 7 else None
        rec = {
            'idx': len(records), 'date': d,
            'voucher_type': str(vtype) if vtype else '',
            'voucher_no': str(vno) if vno else '',
            'debit': debit, 'credit': credit, 'balance': balance,
            'counterparty': str(cp) if cp else '',
            'summary': str(sm) if sm else '',
            'source_row': row_idx,
            'matched': False, 'match_id': None, 'match_type': None,
            'amount': max(debit, credit),
        }
        records.append(rec)
    return records


def read_header_rows(path):
    """只读 Excel 前 3 行（用于提取银行名称/账号），流式读取避免整表加载。"""
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(min_row=1, max_row=3, max_col=12, values_only=True))
    wb.close()
    return rows


def extract_bank_info(header_rows):
    """从表头行（前3行）提取银行名称和账号。"""
    bank_name, bank_account = '', ''
    for row in header_rows:
        for val in row:
            if val is None:
                continue
            s = str(val)
            acct_m = re.search(r'(\d{12,})', s)
            if acct_m and not bank_account:
                bank_account = acct_m.group(1)
            if '银行' in s and not bank_name:
                bank_name = s.strip()
            if ('账号' in s or '账户' in s) and not bank_account:
                parts = re.findall(r'[\d-]+', s)
                if parts:
                    bank_account = ''.join(parts).replace('-', '')
    return bank_name, bank_account


# ============================================================
# Section 3: 交叉核对 — 索引加速版
# ============================================================

def _mark_cross(stmts_list, ledgers_list, mtype, gid, matches):
    for r in stmts_list:
        r['matched'] = True
        r['match_id'] = gid
        r['match_type'] = mtype
    for r in ledgers_list:
        r['matched'] = True
        r['match_id'] = gid
        r['match_type'] = mtype
    matches.append({
        'group_id': gid, 'match_type': mtype,
        'stmts': list(stmts_list), 'ledgers': list(ledgers_list),
    })


def _flow_dir(r, is_ledger):
    """资金流向: 'in'=收款, 'out'=付款。
    方向约定(两边记账视角相反):
      对账单(银行视角): 贷方=收款(in), 借方=付款(out)
      日记账(企业视角): 借方=收款(in), 贷方=付款(out)
    借贷同记的异常行返回 None(不参与交叉核对)。
    """
    debit = r['debit'] or 0
    credit = r['credit'] or 0
    if debit > 0 and credit > 0:
        return None
    if is_ledger:
        return 'in' if debit > 0 else 'out'
    return 'in' if credit > 0 else 'out'


def _dir_ok(s, l):
    """交叉匹配方向校验: 只允许 对账单借↔日记账贷 / 对账单贷↔日记账借。"""
    sf = _flow_dir(s, False)
    lf = _flow_dir(l, True)
    return sf is not None and lf is not None and sf == lf


def cross_reconcile(stmts, ledgers):
    matches = []
    gc = [1]

    def next_gid():
        gid = f'GROUP{gc[0]:06d}'
        gc[0] += 1
        return gid

    # ---- 构建索引 (一次构建, 运行时通过 matched 标志过滤) ----
    # ledger 索引
    l_by_da = defaultdict(list)   # (date, round(amount,2)) -> [ledger]
    l_by_amt = defaultdict(list)  # round(amount,2) -> [ledger]
    for l in ledgers:
        if l['amount'] > 0:
            a = round(l['amount'], 2)
            l_by_da[(l['date'], a)].append(l)
            l_by_amt[a].append(l)
    # stmt 索引
    s_by_da = defaultdict(list)
    s_by_amt = defaultdict(list)
    for s in stmts:
        if s['amount'] > 0:
            a = round(s['amount'], 2)
            s_by_da[(s['date'], a)].append(s)
            s_by_amt[a].append(s)

    # --- Pass 1: 金额 + 同日 + 对方 ---
    for s in stmts:
        if s['matched'] or s['amount'] == 0:
            continue
        key = (s['date'], round(s['amount'], 2))
        for l in l_by_da.get(key, []):
            if l['matched'] or not _dir_ok(s, l):
                continue
            s_cp = remove_spaces(s['counterparty'])
            l_cp = remove_spaces(l['counterparty'])
            if not s_cp or not l_cp:
                continue
            if s_cp in l_cp or l_cp in s_cp:
                _mark_cross([s], [l], '金额+日期+对方', next_gid(), matches)
                break

    # --- Pass 2: 金额 + 同日 (唯一候选) ---
    for s in stmts:
        if s['matched'] or s['amount'] == 0:
            continue
        key = (s['date'], round(s['amount'], 2))
        candidates = [l for l in l_by_da.get(key, [])
                      if not l['matched'] and _dir_ok(s, l)]
        if len(candidates) == 1:
            _mark_cross([s], candidates, '金额+日期', next_gid(), matches)

    # --- Pass 3: 金额 + 近日 ±3天 (唯一候选) ---
    for s in stmts:
        if s['matched'] or s['amount'] == 0:
            continue
        key = round(s['amount'], 2)
        candidates = [l for l in l_by_amt.get(key, [])
                      if not l['matched'] and _dir_ok(s, l)
                      and abs((s['date'] - l['date']).days) <= 3]
        if len(candidates) == 1:
            _mark_cross([s], candidates, '金额+近日期', next_gid(), matches)

    # --- Pass 4: 金额 + 宽日期 ±7天 (唯一候选) ---
    for s in stmts:
        if s['matched'] or s['amount'] == 0:
            continue
        key = round(s['amount'], 2)
        candidates = [l for l in l_by_amt.get(key, [])
                      if not l['matched'] and _dir_ok(s, l)
                      and abs((s['date'] - l['date']).days) <= 7]
        if len(candidates) == 1:
            _mark_cross([s], candidates, '金额+宽日期', next_gid(), matches)

    # --- Pass 5: 多对一 (凭证汇总) ---
    ledger_groups = defaultdict(list)
    for l in ledgers:
        if not l['matched'] and l['amount'] > 0:
            ledger_groups[(l['date'], l['voucher_type'], l['voucher_no'])].append(l)

    for key, grp in ledger_groups.items():
        active = [r for r in grp if not r['matched']]
        if len(active) < 2:
            continue
        flows = {_flow_dir(r, True) for r in active}
        if None in flows or len(flows) != 1:
            continue  # 组内方向不一致/异常行不参与交叉汇总
        g_flow = flows.pop()
        total = round(sum(r['amount'] for r in active), 2)
        d = key[0]
        target = None
        for s in s_by_da.get((d, total), []):
            if not s['matched'] and _flow_dir(s, False) == g_flow:
                target = s
                break
        if target is None:
            best_delta = 999
            for s in s_by_amt.get(total, []):
                if s['matched'] or _flow_dir(s, False) != g_flow:
                    continue
                dd = abs((s['date'] - d).days)
                if dd <= 3 and dd < best_delta:
                    target = s
                    best_delta = dd
        if target:
            _mark_cross([target], active, '多对一(凭证汇总)', next_gid(), matches)

    # --- Pass 6: 多对一 (对账单汇总) ---
    stmt_groups = defaultdict(list)
    for s in stmts:
        if not s['matched'] and s['amount'] > 0:
            stmt_groups[(s['date'], normalize_cp(s['counterparty']))].append(s)

    for key, grp in stmt_groups.items():
        active = [r for r in grp if not r['matched']]
        if len(active) < 2:
            continue
        flows = {_flow_dir(r, False) for r in active}
        if None in flows or len(flows) != 1:
            continue  # 组内方向不一致/异常行不参与交叉汇总
        g_flow = flows.pop()
        total = round(sum(r['amount'] for r in active), 2)
        d = key[0]
        target = None
        best_delta = 999
        for l in l_by_amt.get(total, []):
            if l['matched'] or _flow_dir(l, True) != g_flow:
                continue
            dd = abs((l['date'] - d).days)
            if dd <= 3 and dd < best_delta:
                target = l
                best_delta = dd
        if target:
            _mark_cross(active, [target], '多对一(对账单汇总)', next_gid(), matches)

    # --- Pass 7: 一借一贷自动抵消 ---
    for s in stmts:
        if s['matched'] or s['amount'] == 0:
            continue
        key = (s['date'], round(s['amount'], 2))
        for l in l_by_da.get(key, []):
            if not l['matched'] and _dir_ok(s, l):
                _mark_cross([s], [l], '一借一贷抵消', next_gid(), matches)
                break

    # --- Pass 8: 多对多 ---
    def _do_m2m(stmt_key_fn, ledger_key_fn, mtype):
        # 组键附带资金流向: 方向不一致的记录分组即隔离，永不可能配对
        sg = defaultdict(list)
        for s in stmts:
            if not s['matched'] and s['amount'] > 0:
                f = _flow_dir(s, False)
                if f is None:
                    continue
                sg[(stmt_key_fn(s), f)].append(s)
        lg = defaultdict(list)
        for l in ledgers:
            if not l['matched'] and l['amount'] > 0:
                f = _flow_dir(l, True)
                if f is None:
                    continue
                lg[(ledger_key_fn(l), f)].append(l)

        # 按组总额建索引: total -> [组key]，把 O(组数×组数) 全量配对降为 O(组数) 查找。
        # 注: Pass8 匹配总是整组消费，组内成员一旦匹配即整组出局，
        #     因此建索引时的总额在本次 _do_m2m 内不会变化，语义与全量遍历完全一致。
        lg_by_total = defaultdict(list)
        for lk, lgrp in lg.items():
            total = round(sum(r['amount'] for r in lgrp), 2)
            lg_by_total[total].append(lk)

        for sk, sgrp in sg.items():
            s_active = [r for r in sgrp if not r['matched']]
            if not s_active:
                continue
            s_total = round(sum(r['amount'] for r in s_active), 2)
            s_date = sk[0][0]
            for lk in lg_by_total.get(s_total, []):
                if lk[1] != sk[1]:
                    continue  # 资金流向不一致，不配对
                l_active = [r for r in lg[lk] if not r['matched']]
                if not l_active:
                    continue
                l_date = lk[0][0]
                if abs((s_date - l_date).days) <= 3:
                    if len(s_active) >= 2 or len(l_active) >= 2:
                        _mark_cross(s_active, l_active, mtype, next_gid(), matches)

    _do_m2m(lambda s: (s['date'], normalize_cp(s['counterparty'])),
            lambda l: (l['date'], normalize_cp(l['counterparty'])),
            '多对多（交叉核对）')
    _do_m2m(lambda s: (s['date'], normalize_cp(s['counterparty'])),
            lambda l: (l['date'], remove_spaces(l['summary'][:10])),
            '多对多（交叉核对）')

    return matches


# ============================================================
# Section 4 & 5: 借贷自核对 — 索引加速版
# ============================================================

def _split_debit_credit(records):
    debits, credits = [], []
    for r in records:
        if r['matched'] or r['amount'] == 0:
            continue
        if r['debit'] > 0 and r['credit'] == 0:
            debits.append(r)
        elif r['credit'] > 0 and r['debit'] == 0:
            credits.append(r)
        elif r['debit'] > r['credit']:
            debits.append(r)
        elif r['credit'] > 0:
            credits.append(r)
    return debits, credits


def _mark_self(debits_list, credits_list, mtype, gid, matches):
    for r in debits_list:
        r['matched'] = True
        r['match_id'] = gid
        r['match_type'] = mtype
    for r in credits_list:
        r['matched'] = True
        r['match_id'] = gid
        r['match_type'] = mtype
    matches.append({
        'group_id': gid, 'match_type': mtype,
        'debits': list(debits_list), 'credits': list(credits_list),
    })


def _build_self_indexes(credits):
    """为贷方记录构建索引"""
    c_by_da = defaultdict(list)
    c_by_amt = defaultdict(list)
    for c in credits:
        a = round(c['amount'], 2)
        c_by_da[(c['date'], a)].append(c)
        c_by_amt[a].append(c)
    return c_by_da, c_by_amt


def ledger_self_reconcile(records):
    matches = []
    gc = [1]

    def next_gid():
        gid = f'LS{gc[0]:05d}'
        gc[0] += 1
        return gid

    debits, credits = _split_debit_credit(records)
    c_by_da, c_by_amt = _build_self_indexes(credits)

    # --- LS Pass 1: 同日同金额 ---
    for d in debits:
        if d['matched']:
            continue
        key = (d['date'], round(d['amount'], 2))
        for c in c_by_da.get(key, []):
            if not c['matched']:
                _mark_self([d], [c], '金额+日期（自核对）', next_gid(), matches)
                break

    # --- LS Pass 2: 近日期 ±3天 ---
    for d in debits:
        if d['matched']:
            continue
        key = round(d['amount'], 2)
        for c in c_by_amt.get(key, []):
            if not c['matched'] and abs((d['date'] - c['date']).days) <= 3:
                _mark_self([d], [c], '金额+近日期（自核对）', next_gid(), matches)
                break

    # --- LS Pass 3: 同日多借一贷 ---
    d_groups = defaultdict(list)
    for d in debits:
        if not d['matched']:
            d_groups[d['date']].append(d)
    for d_date, d_grp in d_groups.items():
        active = [r for r in d_grp if not r['matched']]
        if len(active) < 2:
            continue
        total = round(sum(r['amount'] for r in active), 2)
        for c in c_by_amt.get(total, []):
            if not c['matched'] and c['date'] == d_date:
                _mark_self(active, [c], '同日多借一贷（自核对）', next_gid(), matches)
                break

    # --- LS Pass 4: 同日多贷一借 ---
    c_groups = defaultdict(list)
    for c in credits:
        if not c['matched']:
            c_groups[c['date']].append(c)
    for c_date, c_grp in c_groups.items():
        active = [r for r in c_grp if not r['matched']]
        if len(active) < 2:
            continue
        total = round(sum(r['amount'] for r in active), 2)
        for d in debits:
            if d['matched']:
                continue
            if d['date'] == c_date and abs(d['amount'] - total) < 0.01:
                _mark_self([d], active, '同日多贷一借（自核对）', next_gid(), matches)
                break

    # --- LS Pass 5: 宽日期一对一 (0~±3天逐日尝试) ---
    for d in debits:
        if d['matched']:
            continue
        key = round(d['amount'], 2)
        for delta in range(0, 4):
            for sign in (1, -1) if delta > 0 else (1,):
                target_date = d['date'] + timedelta(days=delta * sign)
                for c in c_by_amt.get(key, []):
                    if not c['matched'] and c['date'] == target_date:
                        _mark_self([d], [c], '金额+宽日期（自核对）', next_gid(), matches)
                        break
                else:
                    continue
                break
            else:
                continue
            break

    # --- LS Pass 6: 多对多 (同日 + 近日±1) ---
    d_groups2 = defaultdict(list)
    for d in debits:
        if not d['matched']:
            d_groups2[d['date']].append(d)
    c_groups2 = defaultdict(list)
    for c in credits:
        if not c['matched']:
            c_groups2[c['date']].append(c)

    for d_date, d_grp in d_groups2.items():
        d_active = [r for r in d_grp if not r['matched']]
        if len(d_active) < 2:
            continue
        d_total = round(sum(r['amount'] for r in d_active), 2)
        if d_date in c_groups2:
            c_active = [r for r in c_groups2[d_date] if not r['matched']]
            if len(c_active) >= 2:
                c_total = round(sum(r['amount'] for r in c_active), 2)
                if abs(d_total - c_total) < 0.01:
                    _mark_self(d_active, c_active, '同日多对多（自核对）', next_gid(), matches)
                    continue
        for delta in (1, -1):
            near_date = d_date + timedelta(days=delta)
            if near_date in c_groups2:
                c_active = [r for r in c_groups2[near_date] if not r['matched']]
                if len(c_active) >= 2:
                    c_total = round(sum(r['amount'] for r in c_active), 2)
                    if abs(d_total - c_total) < 0.01:
                        _mark_self(d_active, c_active, '近日多对多（自核对）', next_gid(), matches)
                        break

    return matches


def stmt_self_reconcile(records):
    matches = []
    gc = [1]

    def next_gid():
        gid = f'SS{gc[0]:05d}'
        gc[0] += 1
        return gid

    debits, credits = _split_debit_credit(records)
    c_by_da, c_by_amt = _build_self_indexes(credits)

    # --- SS Pass 1: 同日同金额 ---
    for d in debits:
        if d['matched']:
            continue
        key = (d['date'], round(d['amount'], 2))
        for c in c_by_da.get(key, []):
            if not c['matched']:
                _mark_self([d], [c], '金额+日期（自核对）', next_gid(), matches)
                break

    # --- SS Pass 2: 同日多借一贷 ---
    d_groups = defaultdict(list)
    for d in debits:
        if not d['matched']:
            d_groups[d['date']].append(d)
    for d_date, d_grp in d_groups.items():
        active = [r for r in d_grp if not r['matched']]
        if len(active) < 2:
            continue
        total = round(sum(r['amount'] for r in active), 2)
        for c in c_by_amt.get(total, []):
            if not c['matched'] and c['date'] == d_date:
                _mark_self(active, [c], '同日多借一贷（自核对）', next_gid(), matches)
                break

    # --- SS Pass 3: 同日多贷一借 ---
    c_groups = defaultdict(list)
    for c in credits:
        if not c['matched']:
            c_groups[c['date']].append(c)
    for c_date, c_grp in c_groups.items():
        active = [r for r in c_grp if not r['matched']]
        if len(active) < 2:
            continue
        total = round(sum(r['amount'] for r in active), 2)
        for d in debits:
            if d['matched']:
                continue
            if d['date'] == c_date and abs(d['amount'] - total) < 0.01:
                _mark_self([d], active, '同日多贷一借（自核对）', next_gid(), matches)
                break

    # --- SS Pass 4: 近日期一对一 ---
    for d in debits:
        if d['matched']:
            continue
        key = round(d['amount'], 2)
        for c in c_by_amt.get(key, []):
            if not c['matched'] and abs((d['date'] - c['date']).days) <= 3:
                _mark_self([d], [c], '金额+近日期（自核对）', next_gid(), matches)
                break

    # --- SS Pass 5: 多对多 (同日 + 近日±1) ---
    d_groups2 = defaultdict(list)
    for d in debits:
        if not d['matched']:
            d_groups2[d['date']].append(d)
    c_groups2 = defaultdict(list)
    for c in credits:
        if not c['matched']:
            c_groups2[c['date']].append(c)

    for d_date, d_grp in d_groups2.items():
        d_active = [r for r in d_grp if not r['matched']]
        if len(d_active) < 2:
            continue
        d_total = round(sum(r['amount'] for r in d_active), 2)
        if d_date in c_groups2:
            c_active = [r for r in c_groups2[d_date] if not r['matched']]
            if len(c_active) >= 2:
                c_total = round(sum(r['amount'] for r in c_active), 2)
                if abs(d_total - c_total) < 0.01:
                    _mark_self(d_active, c_active, '同日多对多（自核对）', next_gid(), matches)
                    continue
        for delta in (1, -1):
            near_date = d_date + timedelta(days=delta)
            if near_date in c_groups2:
                c_active = [r for r in c_groups2[near_date] if not r['matched']]
                if len(c_active) >= 2:
                    c_total = round(sum(r['amount'] for r in c_active), 2)
                    if abs(d_total - c_total) < 0.01:
                        _mark_self(d_active, c_active, '近日多对多（自核对）', next_gid(), matches)
                        break

    return matches


# ============================================================
# Pass 9: 日汇总核对（兜底，流程最后执行）
# 业务场景: 发工资/收货款等——流水(对账单)一天发出去很多笔/收到很多笔，
#           账上(日记账)只记一笔汇总(或少数几笔经手人汇总)。
# 规则: 在所有交叉 Pass1-8 和两边自核对完成之后，对仍然未匹配的记录，
#       按【交易日期】+【收支方向】分组：
#         对账单贷方(收款)  <->  日记账借方(收款)
#         对账单借方(付款)  <->  日记账贷方(付款)
#       同一天、同方向两边的合计金额相等，且至少一边 >= 2 条 -> 整组合计为一组"日汇总"核对。
# ============================================================

def _daily_bucket(records, is_ledger):
    """按(日期, 收支方向)归类未匹配记录。
    方向语义: 'in' 收款 / 'out' 付款。
    注意两边记账方向约定相反: 日记账借方=收款(资产借增)，对账单贷方=收款(银行视角)。
    """
    buckets = defaultdict(list)  # (date, 'in'|'out') -> [record]
    for r in records:
        if r['matched'] or r['amount'] <= 0:
            continue
        debit = r['debit'] or 0
        credit = r['credit'] or 0
        if debit > 0 and credit > 0:
            continue  # 借贷同时存在的异常行不参与日汇总
        if is_ledger:
            # 日记账: 借方=收款(in), 贷方=付款(out)
            direction = 'in' if debit > 0 else 'out'
        else:
            # 对账单: 贷方=收款(in), 借方=付款(out)
            direction = 'in' if credit > 0 else 'out'
        d = r['date']
        dd = d.date() if isinstance(d, datetime) else d
        buckets[(dd, direction)].append(r)
    return buckets


def daily_summary_cross(stmts, ledgers, gid_start=1):
    """日汇总核对，返回 (matches, 下一个组号)。只处理双方都未匹配的记录。"""
    matches = []
    gc = [gid_start]

    def next_gid():
        gid = f'GROUP{gc[0]:06d}'
        gc[0] += 1
        return gid

    stmt_b = _daily_bucket(stmts, is_ledger=False)
    led_b = _daily_bucket(ledgers, is_ledger=True)

    # 同方向才配对: (date,'in') 与 (date,'in') 比，(date,'out') 与 (date,'out') 比
    for direction in ('in', 'out'):
        for (dd, _dir) in list(stmt_b.keys()):
            if _dir != direction:
                continue
            key = (dd, direction)
            s_grp = [r for r in stmt_b[key] if not r['matched']]
            l_grp = [r for r in led_b.get(key, []) if not r['matched']]
            if not s_grp or not l_grp:
                continue
            if len(s_grp) < 2 and len(l_grp) < 2:
                continue  # 单对单且同额同日应早已被前面 Pass 匹配，这里不兜 1:1
            s_total = round(sum(r['amount'] for r in s_grp), 2)
            l_total = round(sum(r['amount'] for r in l_grp), 2)
            if abs(s_total - l_total) < 0.005:
                _mark_cross(s_grp, l_grp, '日汇总', next_gid(), matches)

    return matches, gc[0]


# ============================================================
# Section 6: 统计
# ============================================================
def compute_stats(stmts, ledgers, cross_matches, stmt_self, ledger_self):
    cross_stmt_idx = set()
    cross_ledger_idx = set()
    for m in cross_matches:
        for r in m['stmts']:
            cross_stmt_idx.add(r['idx'])
        for r in m['ledgers']:
            cross_ledger_idx.add(r['idx'])

    stmt_self_idx = set()
    for m in stmt_self:
        for r in m['debits'] + m['credits']:
            stmt_self_idx.add(r['idx'])
    ledger_self_idx = set()
    for m in ledger_self:
        for r in m['debits'] + m['credits']:
            ledger_self_idx.add(r['idx'])

    def _stats(records, cross_idx, self_idx):
        cross_debit = sum(r['debit'] for r in records if r['idx'] in cross_idx and r['debit'] > 0)
        cross_credit = sum(r['credit'] for r in records if r['idx'] in cross_idx and r['credit'] > 0)
        self_amt = sum(r['debit'] for r in records if r['idx'] in self_idx and r['debit'] > 0)
        total_debit = sum(r['debit'] for r in records)
        total_credit = sum(r['credit'] for r in records)
        denom_d = total_debit - self_amt
        denom_c = total_credit - self_amt
        ratio_d = (cross_debit / denom_d * 100) if denom_d > 0.01 else 0.0
        ratio_c = (cross_credit / denom_c * 100) if denom_c > 0.01 else 0.0
        dates = [r['date'] for r in records if r['date']]
        period = f'{min(dates).strftime("%Y-%m-%d")} ~ {max(dates).strftime("%Y-%m-%d")}' if dates else ''
        return {
            'period': period,
            'cross_debit': round(cross_debit, 2),
            'cross_credit': round(cross_credit, 2),
            'self_amount': round(self_amt, 2),
            'total_debit': round(total_debit, 2),
            'total_credit': round(total_credit, 2),
            'ratio_debit': round(ratio_d, 2),
            'ratio_credit': round(ratio_c, 2),
        }

    stmt_stats = _stats(stmts, cross_stmt_idx, stmt_self_idx)
    ledger_stats = _stats(ledgers, cross_ledger_idx, ledger_self_idx)
    return stmt_stats, ledger_stats


# ============================================================
# Section 7: 输出 (7 Sheet)
# ============================================================

def _fmt_date(d):
    if d is None:
        return ''
    if isinstance(d, datetime):
        return d.strftime('%Y-%m-%d')
    return str(d)


def _write_header(ws, headers, fill=None):
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = HEADER_FONT
        if fill:
            cell.fill = fill
        cell.alignment = CENTER_AL
        cell.border = THIN_BORDER
    ws.freeze_panes = 'A2'


def _set_cell(ws, r, c, val, font=None, fill=None, align=None, border=None, num_fmt=None):
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = font or DATA_FONT
    if fill:
        cell.fill = fill
    cell.alignment = align or LEFT_AL
    cell.border = border or THIN_BORDER
    if num_fmt:
        cell.number_format = num_fmt
    return cell


def _apply_group_border(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            ex = cell.border
            left = ex.left or THIN_SIDE
            right = ex.right or THIN_SIDE
            top = ex.top or THIN_SIDE
            bottom = ex.bottom or THIN_SIDE
            if r == r1:
                top = GROUP_RED
            if r == r2:
                bottom = GROUP_RED
            if c == c1:
                left = GROUP_RED
            if c == c2:
                right = GROUP_RED
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)


def _set_col_widths(ws, widths):
    from openpyxl.utils import get_column_letter
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_sheet1(wb, cross_matches):
    ws = wb.create_sheet('交叉核对-对账单出发')
    headers = ['核对方式', '对账单日期', '借方金额', '贷方金额', '对方户名', '对账单摘要',
               '日记账日期', '凭证类型', '凭证序号', '借方金额', '贷方金额',
               '对方名称', '日记账摘要', '组ID']
    _write_header(ws, headers)
    _set_col_widths(ws, [18, 12, 12, 12, 18, 22, 12, 10, 10, 12, 12, 18, 22, 14])
    r = 2
    for m in cross_matches:
        stmts, ledgers = m['stmts'], m['ledgers']
        nrows = max(len(stmts), len(ledgers))
        is_multi = len(stmts) >= 2 or len(ledgers) >= 2
        r1 = r
        for i in range(nrows):
            _set_cell(ws, r, 1, m['match_type'] if i == 0 else '', align=CENTER_AL)
            if i < len(stmts):
                s = stmts[i]
                _set_cell(ws, r, 2, _fmt_date(s['date']), align=CENTER_AL)
                _set_cell(ws, r, 3, s['debit'] if s['debit'] > 0 else '', align=RIGHT_AL, num_fmt=NUM_FMT)
                _set_cell(ws, r, 4, s['credit'] if s['credit'] > 0 else '', align=RIGHT_AL, num_fmt=NUM_FMT)
                _set_cell(ws, r, 5, s['counterparty'], align=CENTER_AL)
                _set_cell(ws, r, 6, s['summary'], align=LEFT_AL)
            if i < len(ledgers):
                l = ledgers[i]
                _set_cell(ws, r, 7, _fmt_date(l['date']), align=CENTER_AL)
                _set_cell(ws, r, 8, l['voucher_type'], align=CENTER_AL)
                _set_cell(ws, r, 9, l['voucher_no'], align=CENTER_AL)
                _set_cell(ws, r, 10, l['debit'] if l['debit'] > 0 else '', align=RIGHT_AL, num_fmt=NUM_FMT)
                _set_cell(ws, r, 11, l['credit'] if l['credit'] > 0 else '', align=RIGHT_AL, num_fmt=NUM_FMT)
                _set_cell(ws, r, 12, l['counterparty'], align=CENTER_AL)
                _set_cell(ws, r, 13, l['summary'], align=LEFT_AL)
            _set_cell(ws, r, 14, m['group_id'] if i == 0 else '', align=CENTER_AL)
            r += 1
        if is_multi:
            _apply_group_border(ws, r1, r - 1, 1, 14)
    return ws


def write_sheet2(wb, cross_matches):
    ws = wb.create_sheet('交叉核对-日记账出发')
    headers = ['核对方式', '日记账日期', '凭证类型', '凭证序号', '借方金额', '贷方金额',
               '对方名称', '日记账摘要', '对账单日期', '借方金额', '贷方金额', '对方户名',
               '对账单摘要', '组ID']
    _write_header(ws, headers)
    _set_col_widths(ws, [18, 12, 10, 10, 12, 12, 18, 22, 12, 12, 12, 18, 22, 14])
    r = 2
    for m in cross_matches:
        stmts, ledgers = m['stmts'], m['ledgers']
        nrows = max(len(stmts), len(ledgers))
        is_multi = len(stmts) >= 2 or len(ledgers) >= 2
        r1 = r
        for i in range(nrows):
            _set_cell(ws, r, 1, m['match_type'] if i == 0 else '', align=CENTER_AL)
            if i < len(ledgers):
                l = ledgers[i]
                _set_cell(ws, r, 2, _fmt_date(l['date']), align=CENTER_AL)
                _set_cell(ws, r, 3, l['voucher_type'], align=CENTER_AL)
                _set_cell(ws, r, 4, l['voucher_no'], align=CENTER_AL)
                _set_cell(ws, r, 5, l['debit'] if l['debit'] > 0 else '', align=RIGHT_AL, num_fmt=NUM_FMT)
                _set_cell(ws, r, 6, l['credit'] if l['credit'] > 0 else '', align=RIGHT_AL, num_fmt=NUM_FMT)
                _set_cell(ws, r, 7, l['counterparty'], align=CENTER_AL)
                _set_cell(ws, r, 8, l['summary'], align=LEFT_AL)
            if i < len(stmts):
                s = stmts[i]
                _set_cell(ws, r, 9, _fmt_date(s['date']), align=CENTER_AL)
                _set_cell(ws, r, 10, s['debit'] if s['debit'] > 0 else '', align=RIGHT_AL, num_fmt=NUM_FMT)
                _set_cell(ws, r, 11, s['credit'] if s['credit'] > 0 else '', align=RIGHT_AL, num_fmt=NUM_FMT)
                _set_cell(ws, r, 12, s['counterparty'], align=CENTER_AL)
                _set_cell(ws, r, 13, s['summary'], align=LEFT_AL)
            _set_cell(ws, r, 14, m['group_id'] if i == 0 else '', align=CENTER_AL)
            r += 1
        if is_multi:
            _apply_group_border(ws, r1, r - 1, 1, 14)
    return ws


def write_sheet3(wb, ledger_self_matches):
    ws = wb.create_sheet('日记账借贷自核对')
    headers = ['核对方式', '借方日期', '凭证类型', '凭证序号', '借方金额', '对方名称', '摘要',
               '贷方日期', '凭证类型', '凭证序号', '贷方金额', '对方名称', '摘要', '组ID']
    _write_header(ws, headers)
    _set_col_widths(ws, [20, 12, 10, 10, 12, 18, 22, 12, 10, 10, 12, 18, 22, 14])
    r = 2
    for m in ledger_self_matches:
        debits, credits = m['debits'], m['credits']
        nrows = max(len(debits), len(credits))
        r1 = r
        is_multi = len(debits) >= 2 or len(credits) >= 2
        for i in range(nrows):
            _set_cell(ws, r, 1, m['match_type'] if i == 0 else '', align=CENTER_AL)
            if i < len(debits):
                d = debits[i]
                _set_cell(ws, r, 2, _fmt_date(d['date']), align=CENTER_AL)
                _set_cell(ws, r, 3, d['voucher_type'], align=CENTER_AL)
                _set_cell(ws, r, 4, d['voucher_no'], align=CENTER_AL)
                _set_cell(ws, r, 5, d['amount'], align=RIGHT_AL, num_fmt=NUM_FMT)
                _set_cell(ws, r, 6, d['counterparty'], align=CENTER_AL)
                _set_cell(ws, r, 7, d['summary'], align=LEFT_AL)
            else:
                for c in range(2, 8):
                    _set_cell(ws, r, c, '')
            if i < len(credits):
                c = credits[i]
                _set_cell(ws, r, 8, _fmt_date(c['date']), align=CENTER_AL)
                _set_cell(ws, r, 9, c['voucher_type'], align=CENTER_AL)
                _set_cell(ws, r, 10, c['voucher_no'], align=CENTER_AL)
                _set_cell(ws, r, 11, c['amount'], align=RIGHT_AL, num_fmt=NUM_FMT)
                _set_cell(ws, r, 12, c['counterparty'], align=CENTER_AL)
                _set_cell(ws, r, 13, c['summary'], align=LEFT_AL)
            else:
                for c in range(8, 14):
                    _set_cell(ws, r, c, '')
            _set_cell(ws, r, 14, m['group_id'] if i == 0 else '', align=CENTER_AL)
            r += 1
        if is_multi:
            _apply_group_border(ws, r1, r - 1, 1, 14)
    return ws


def write_sheet4(wb, stmt_self_matches):
    ws = wb.create_sheet('对账单借贷自核对')
    headers = ['核对方式', '借方日期', '金额', '对方户名', '借方摘要',
               '贷方日期', '金额', '对方户名', '贷方摘要', '组ID']
    _write_header(ws, headers)
    _set_col_widths(ws, [20, 12, 12, 18, 22, 12, 12, 18, 22, 14])
    r = 2
    for m in stmt_self_matches:
        debits, credits = m['debits'], m['credits']
        nrows = max(len(debits), len(credits))
        r1 = r
        is_multi = len(debits) >= 2 or len(credits) >= 2
        for i in range(nrows):
            _set_cell(ws, r, 1, m['match_type'] if i == 0 else '', align=CENTER_AL)
            if i < len(debits):
                d = debits[i]
                _set_cell(ws, r, 2, _fmt_date(d['date']), align=CENTER_AL)
                _set_cell(ws, r, 3, d['amount'], align=RIGHT_AL, num_fmt=NUM_FMT)
                _set_cell(ws, r, 4, d['counterparty'], align=CENTER_AL)
                _set_cell(ws, r, 5, d['summary'], align=LEFT_AL)
            else:
                for c in range(2, 6):
                    _set_cell(ws, r, c, '')
            if i < len(credits):
                c = credits[i]
                _set_cell(ws, r, 6, _fmt_date(c['date']), align=CENTER_AL)
                _set_cell(ws, r, 7, c['amount'], align=RIGHT_AL, num_fmt=NUM_FMT)
                _set_cell(ws, r, 8, c['counterparty'], align=CENTER_AL)
                _set_cell(ws, r, 9, c['summary'], align=LEFT_AL)
            else:
                for c in range(6, 10):
                    _set_cell(ws, r, c, '')
            _set_cell(ws, r, 10, m['group_id'] if i == 0 else '', align=CENTER_AL)
            r += 1
        if is_multi:
            _apply_group_border(ws, r1, r - 1, 1, 10)
    return ws


def write_sheet5(wb, stmts):
    ws = wb.create_sheet('未核对对账单')
    headers = ['交易日期', '借方金额', '贷方金额', '余额', '对方户名', '对方账号', '摘要', '方向']
    _write_header(ws, headers)
    _set_col_widths(ws, [12, 12, 12, 12, 18, 18, 22, 10])
    r = 2
    for s in stmts:
        if s['matched']:
            continue
        _set_cell(ws, r, 1, _fmt_date(s['date']), align=CENTER_AL)
        _set_cell(ws, r, 2, s['debit'] if s['debit'] > 0 else '', align=RIGHT_AL, num_fmt=NUM_FMT)
        _set_cell(ws, r, 3, s['credit'] if s['credit'] > 0 else '', align=RIGHT_AL, num_fmt=NUM_FMT)
        _set_cell(ws, r, 4, s['balance'], align=RIGHT_AL, num_fmt=NUM_FMT)
        _set_cell(ws, r, 5, s['counterparty'], align=CENTER_AL)
        _set_cell(ws, r, 6, s['counterparty_acct'], align=CENTER_AL)
        _set_cell(ws, r, 7, s['summary'], align=LEFT_AL)
        direction = '借方(钱出)' if s['debit'] > 0 else ('贷方(钱入)' if s['credit'] > 0 else '')
        _set_cell(ws, r, 8, direction, align=CENTER_AL)
        r += 1
    return ws


def write_sheet6(wb, ledgers):
    ws = wb.create_sheet('未核对日记账')
    headers = ['记账日期', '凭证类型', '凭证序号', '借方金额', '贷方金额', '余额', '对方名称', '摘要', '方向']
    _write_header(ws, headers)
    _set_col_widths(ws, [12, 10, 10, 12, 12, 12, 18, 22, 10])
    r = 2
    for l in ledgers:
        if l['matched']:
            continue
        _set_cell(ws, r, 1, _fmt_date(l['date']), align=CENTER_AL)
        _set_cell(ws, r, 2, l['voucher_type'], align=CENTER_AL)
        _set_cell(ws, r, 3, l['voucher_no'], align=CENTER_AL)
        _set_cell(ws, r, 4, l['debit'] if l['debit'] > 0 else '', align=RIGHT_AL, num_fmt=NUM_FMT)
        _set_cell(ws, r, 5, l['credit'] if l['credit'] > 0 else '', align=RIGHT_AL, num_fmt=NUM_FMT)
        _set_cell(ws, r, 6, l['balance'], align=RIGHT_AL, num_fmt=NUM_FMT)
        _set_cell(ws, r, 7, l['counterparty'], align=CENTER_AL)
        _set_cell(ws, r, 8, l['summary'], align=LEFT_AL)
        direction = '借方(钱入)' if l['debit'] > 0 else ('贷方(钱出)' if l['credit'] > 0 else '')
        _set_cell(ws, r, 9, direction, align=CENTER_AL)
        r += 1
    return ws


def write_sheet7(wb, stmt_stats, ledger_stats, bank_name, bank_account):
    ws = wb.create_sheet('统计')
    headers = ['核对期间', '银行名称', '银行账号', '借方已核对', '贷方已核对',
               '自核对金额', '借方合计', '贷方合计', '借方核对比例', '贷方核对比例', '核对口径']
    _write_header(ws, headers)
    _set_col_widths(ws, [24, 18, 20, 14, 14, 14, 14, 14, 14, 14, 24])
    scope = '交叉核对+自核对'
    for row_idx, st in enumerate([stmt_stats, ledger_stats], 2):
        _set_cell(ws, row_idx, 1, st['period'], align=CENTER_AL)
        _set_cell(ws, row_idx, 2, bank_name, align=CENTER_AL)
        _set_cell(ws, row_idx, 3, bank_account, align=CENTER_AL)
        _set_cell(ws, row_idx, 4, st['cross_debit'], align=RIGHT_AL, num_fmt=NUM_FMT)
        _set_cell(ws, row_idx, 5, st['cross_credit'], align=RIGHT_AL, num_fmt=NUM_FMT)
        _set_cell(ws, row_idx, 6, st['self_amount'], align=RIGHT_AL, num_fmt=NUM_FMT)
        _set_cell(ws, row_idx, 7, st['total_debit'], align=RIGHT_AL, num_fmt=NUM_FMT)
        _set_cell(ws, row_idx, 8, st['total_credit'], align=RIGHT_AL, num_fmt=NUM_FMT)
        _set_cell(ws, row_idx, 9, f'{st["ratio_debit"]:.2f}%', align=CENTER_AL)
        _set_cell(ws, row_idx, 10, f'{st["ratio_credit"]:.2f}%', align=CENTER_AL)
        _set_cell(ws, row_idx, 11, scope, align=CENTER_AL)
    return ws


# ============================================================
# 核心流程 — GUI 与命令行共用
# ============================================================

def _find_sheet(wb, names):
    """按候选 Sheet 名查找工作表，找不到则取第一个。"""
    for name in names:
        if name in wb.sheetnames:
            return wb[name]
    return wb.worksheets[0]


def run_reconcile(stmt_file, ledger_file=None, output_file=None, log=None):
    """执行完整对账流程（读取 → 交叉核对 → 自核对 → 统计 → 7-Sheet 输出）。

    参数:
        stmt_file:   对账单 Excel 路径（Sheet1，第 4 行起数据）
        ledger_file: 日记账 Excel 路径（银行明细账）；为 None 时视为单文件合并模式
        output_file: 输出 Excel 路径；None 时自动生成
        log:         日志回调 log(str)；None 时使用 print
    返回:
        dict 汇总信息（输出路径、笔数、匹配组数等），供界面展示
    """
    if log is None:
        log = print

    # ---- 确定输出文件名 ----
    if output_file is None:
        if ledger_file:
            base = os.path.splitext(os.path.basename(stmt_file))[0]
            output_file = os.path.join(os.path.dirname(stmt_file) or '.', f'{base}_对账结果.xlsx')
        else:
            base, ext = os.path.splitext(stmt_file)
            output_file = f'{base}_对账结果.xlsx'
    os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)

    log('[1/7] 读取输入文件...')
    if ledger_file:
        log(f'  对账单: {stmt_file}')
        log(f'  日记账: {ledger_file}')
        wb_stmt = load_workbook(stmt_file, data_only=True, read_only=True)
        stmt_ws = _find_sheet(wb_stmt, ('Sheet1', 'sheet1', '对账单'))
        wb_ledger = load_workbook(ledger_file, data_only=True, read_only=True)
        ledger_ws = _find_sheet(wb_ledger, ('银行明细账', '日记账', '明细账'))
    else:
        log(f'  合并文件: {stmt_file}')
        wb_in = load_workbook(stmt_file, data_only=True, read_only=True)
        stmt_ws = _find_sheet(wb_in, ('Sheet1', 'sheet1', '对账单'))
        ledger_ws = _find_sheet(wb_in, ('银行明细账', '日记账', '明细账'))
        if ledger_ws is None:
            ledger_ws = wb_in.worksheets[1] if len(wb_in.worksheets) >= 2 else stmt_ws

    log(f'  对账单 Sheet: {stmt_ws.title}')
    log(f'  日记账 Sheet: {ledger_ws.title}')

    # 提取银行信息（流式只读前3行，不再整表加载）
    header_rows = read_header_rows(stmt_file)
    bank_name, bank_account = extract_bank_info(header_rows)
    log(f'  银行名称: {bank_name or "(未识别)"}')
    log(f'  银行账号: {bank_account or "(未识别)"}')

    # 读取数据
    stmts = read_statements(stmt_ws)
    ledgers = read_ledger(ledger_ws)
    log(f'  对账单记录: {len(stmts)} 条')
    log(f'  日记账记录: {len(ledgers)} 条')

    log('[2/7] 交叉核对 (对账单 <-> 日记账)...')
    cross_matches = cross_reconcile(stmts, ledgers)
    cross_s = sum(len(m['stmts']) for m in cross_matches)
    cross_l = sum(len(m['ledgers']) for m in cross_matches)
    log(f'  交叉核对匹配: {len(cross_matches)} 组 (对账单 {cross_s} 条, 日记账 {cross_l} 条)')

    log('[3/7] 日记账借贷自核对...')
    ledger_self = ledger_self_reconcile(ledgers)
    ls_count = sum(len(m['debits']) + len(m['credits']) for m in ledger_self)
    log(f'  日记账自核对: {len(ledger_self)} 组 ({ls_count} 条)')

    log('[4/7] 对账单借贷自核对...')
    stmt_self = stmt_self_reconcile(stmts)
    ss_count = sum(len(m['debits']) + len(m['credits']) for m in stmt_self)
    log(f'  对账单自核对: {len(stmt_self)} 组 ({ss_count} 条)')

    log('[5/7] 日汇总核对 (流水多笔 <-> 账上汇总, 兜底)...')
    daily_matches, _next_gid = daily_summary_cross(stmts, ledgers, gid_start=len(cross_matches) + 1)
    d_s = sum(len(m['stmts']) for m in daily_matches)
    d_l = sum(len(m['ledgers']) for m in daily_matches)
    cross_matches.extend(daily_matches)
    log(f'  日汇总匹配: {len(daily_matches)} 组 (对账单 {d_s} 条, 日记账 {d_l} 条)')

    log('[6/7] 计算统计...')
    stmt_stats, ledger_stats = compute_stats(stmts, ledgers, cross_matches, stmt_self, ledger_self)
    log(f'  对账单 — 借方核对比例: {stmt_stats["ratio_debit"]:.1f}%, 贷方核对比例: {stmt_stats["ratio_credit"]:.1f}%')
    log(f'  日记账 — 借方核对比例: {ledger_stats["ratio_debit"]:.1f}%, 贷方核对比例: {ledger_stats["ratio_credit"]:.1f}%')

    log('[7/7] 生成输出文件...')
    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    write_sheet1(wb_out, cross_matches)
    write_sheet2(wb_out, cross_matches)
    write_sheet3(wb_out, ledger_self)
    write_sheet4(wb_out, stmt_self)
    write_sheet5(wb_out, stmts)
    write_sheet6(wb_out, ledgers)
    write_sheet7(wb_out, stmt_stats, ledger_stats, bank_name, bank_account)
    wb_out.save(output_file)
    log(f'\n完成! 输出文件: {output_file}')

    return {
        'output_file': output_file,
        'bank_name': bank_name or '(未识别)',
        'bank_account': bank_account or '(未识别)',
        'stmt_count': len(stmts),
        'ledger_count': len(ledgers),
        'cross_groups': len(cross_matches) - len(daily_matches),
        'daily_groups': len(daily_matches),
        'stmt_self_groups': len(stmt_self),
        'ledger_self_groups': len(ledger_self),
    }


# ============================================================
# 命令行入口（保留原参数，方便批量/调试）
# ============================================================

def cli_main():
    if len(sys.argv) < 2:
        print('用法:')
        print('  python main.py <对账单文件> <日记账文件> [输出文件]     单对核对')
        print('  python main.py 对账单1 日记账1 对账单2 日记账2 ...      一次多对(两两一组,各生成一份结果)')
        print('  python main.py <合并文件> [输出文件]                    单文件合并模式')
        print('  python main.py                  (打开图形界面)')
        sys.exit(1)
    # 兼容 Windows 控制台中文编码，避免 UnicodeEncodeError
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    args = sys.argv[1:]
    # 批量模式: 4 个及以上且为偶数 → 每两个文件一组 (对账单, 日记账)
    if len(args) >= 4 and len(args) % 2 == 0:
        pairs = list(zip(args[::2], args[1::2]))
        print(f'批量模式: 共 {len(pairs)} 对，将逐对核对并各生成一份结果')
        for i, (s, l) in enumerate(pairs, 1):
            print(f'\n===== 第 {i}/{len(pairs)} 对 =====')
            run_reconcile(s, l, None)
        return
    if len(args) >= 2 and args[1].lower().endswith('.xlsx'):
        # 双文件模式: 对账单 + 日记账
        run_reconcile(args[0], args[1], args[2] if len(args) >= 3 else None)
    else:
        # 单文件合并模式: 文件内含 Sheet1 + 银行明细账
        run_reconcile(args[0], None, args[1] if len(args) >= 2 else None)


# ============================================================
# 图形界面 (tkinter)
# ============================================================

class BankReconApp:
    """简易 GUI：可一次添加多组「对账单+日记账」，逐对核对，每组各生成一份 7-Sheet 结果。"""

    def __init__(self, root):
        self.root = root
        self.running = False
        self.msg_q = queue.Queue()
        self.summaries = []   # 每对的汇总 dict

        root.title('银行对账核对工具')
        root.minsize(860, 660)
        root.geometry('980x720')

        # ---------- 顶部标题 ----------
        header = ttk.Frame(root)
        header.pack(fill='x', padx=14, pady=(12, 2))
        ttk.Label(header, text='银行对账核对工具',
                  font=('Microsoft YaHei UI', 15, 'bold')).pack(anchor='w')
        ttk.Label(header, foreground='#808080',
                  text='可一次添加多组「对账单 + 日记账」，自动逐对核对，每组生成一份 7-Sheet 对账结果。'
                       '（依据《银行对账任务规范》）'
                  ).pack(anchor='w', pady=(2, 0))

        # ---------- 文件对列表 ----------
        frm = ttk.LabelFrame(root, text=' 文件对（每对 = 一个对账单 + 它对应的日记账） ')
        frm.pack(fill='x', padx=14, pady=8)
        cols = ('idx', 'stmt', 'ledger')
        self.tree = ttk.Treeview(frm, columns=cols, show='headings', height=6)
        self.tree.heading('idx', text='#')
        self.tree.heading('stmt', text='对账单文件')
        self.tree.heading('ledger', text='日记账文件')
        self.tree.column('idx', width=36, anchor='center', stretch=False)
        self.tree.column('stmt', width=380, anchor='w')
        self.tree.column('ledger', width=380, anchor='w')
        vsb = ttk.Scrollbar(frm, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, columnspan=3, sticky='ew', padx=(10, 0), pady=(6, 2))
        vsb.grid(row=0, column=3, sticky='ns', padx=(0, 10), pady=(6, 2))
        frm.columnconfigure(1, weight=1)

        btns = ttk.Frame(frm)
        btns.grid(row=1, column=0, columnspan=3, sticky='w', padx=10, pady=(2, 8))
        ttk.Button(btns, text='＋ 添加文件对', command=self._add_pair).pack(side='left')
        ttk.Button(btns, text='－ 移除选中', command=self._remove_selected).pack(side='left', padx=6)
        ttk.Button(btns, text='清空列表', command=self._clear_pairs).pack(side='left')
        ttk.Label(btns, foreground='#a0a0a0',
                  text='提示：每组输出文件自动保存到该组「对账单」同目录，文件名=对账单名_对账结果.xlsx'
                  ).pack(side='left', padx=14)

        # ---------- 操作按钮 + 状态 ----------
        bar = ttk.Frame(root)
        bar.pack(fill='x', padx=14, pady=(4, 0))
        self.btn_start = ttk.Button(bar, text='开始核对', command=self._start)
        self.btn_start.pack(side='left')
        self.status_label = ttk.Label(bar, text='就绪', foreground='#1a7f37')
        self.status_label.pack(side='left', padx=16)
        self.btn_folder = ttk.Button(bar, text='打开结果所在文件夹',
                                     command=self._open_folder, state='disabled')
        self.btn_folder.pack(side='right')
        self.btn_open = ttk.Button(bar, text='打开最后一份结果',
                                   command=self._open_result, state='disabled')
        self.btn_open.pack(side='right', padx=(0, 8))

        # ---------- 日志区 ----------
        log_lf = ttk.LabelFrame(root, text=' 运行日志 ')
        log_lf.pack(fill='both', expand=True, padx=14, pady=(8, 6))
        self.log_text = ScrolledText(log_lf, height=14, state='disabled', wrap='word',
                                     font=('Consolas', 10))
        self.log_text.pack(fill='both', expand=True, padx=4, pady=4)

        self._append_log('使用说明:\n'
                         '  1. 点「添加文件对」，依次选择银行导出的 对账单 Excel 和 日记账 Excel (.xlsx)\n'
                         '  2. 可重复添加多组（例如 2022 年一组、2023 年一组），每组独立生成一份结果\n'
                         '  3. 点「开始核对」逐对自动处理，数万笔数据每组约需 1~3 分钟，请勿关闭窗口\n'
                         '  4. 完成后日志与弹窗会列出每一份结果的保存位置\n')
        self.root.after(150, self._poll)

    # ---------- 文件对增删 ----------
    def _add_pair(self):
        if self.running:
            return
        s = filedialog.askopenfilename(title='选择对账单文件',
                                       filetypes=[('Excel 文件', '*.xlsx'), ('所有文件', '*.*')])
        if not s:
            return
        l = filedialog.askopenfilename(title='选择与该对账单对应的日记账文件',
                                       filetypes=[('Excel 文件', '*.xlsx'), ('所有文件', '*.*')])
        if not l:
            messagebox.showinfo('未添加', '未选择日记账文件，本次未添加文件对。')
            return
        self.tree.insert('', 'end', values=(len(self.tree.get_children()) + 1, s, l))
        self._append_log(f'已添加文件对 #{len(self.tree.get_children())}: '
                         f'{os.path.basename(s)}  +  {os.path.basename(l)}')

    def _remove_selected(self):
        if self.running:
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo('提示', '请先在列表中选中要移除的文件对。')
            return
        for item in sel:
            self.tree.delete(item)
        self._renumber()

    def _clear_pairs(self):
        if self.running:
            return
        self.tree.delete(*self.tree.get_children())
        self._append_log('已清空文件对列表。')

    def _renumber(self):
        for i, item in enumerate(self.tree.get_children(), 1):
            vals = list(self.tree.item(item, 'values'))
            vals[0] = i
            self.tree.item(item, values=vals)

    # ---------- 开始核对 ----------
    def _start(self):
        if self.running:
            return
        pairs = []
        seen = set()
        for item in self.tree.get_children():
            _, s, l = self.tree.item(item, 'values')
            s, l = s.strip(), l.strip()
            for name, p in (('对账单', s), ('日记账', l)):
                if not p:
                    messagebox.showwarning('缺少文件', '存在未选全文件的文件对，请检查或移除。')
                    return
                if not os.path.isfile(p):
                    messagebox.showerror('文件不存在', f'{name}文件不存在:\n{p}')
                    return
                if not p.lower().endswith('.xlsx'):
                    messagebox.showerror('格式不支持', f'{name}文件不是 Excel (.xlsx):\n{p}')
                    return
            if os.path.normpath(s) == os.path.normpath(l):
                messagebox.showerror('路径冲突', '同一组内 对账单 与 日记账 不能是同一个文件。')
                return
            key = (os.path.normpath(s), os.path.normpath(l))
            if key in seen:
                messagebox.showerror('重复文件对', f'存在重复的文件对:\n{s}\n请移除后重试。')
                return
            seen.add(key)
            pairs.append((s, l))
        if not pairs:
            messagebox.showwarning('没有文件对', '请先点「添加文件对」添加至少一组文件。')
            return

        # 切到运行态
        self.running = True
        self.summaries = []
        self.btn_start.config(state='disabled')
        self.btn_open.config(state='disabled')
        self.btn_folder.config(state='disabled')
        self.status_label.config(text=f'核对中… 共 {len(pairs)} 对', foreground='#b35900')
        self.root.title('银行对账核对工具 — 核对中…')
        self.log_text.config(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.config(state='disabled')
        self._append_log(f'共 {len(pairs)} 对文件，开始逐对核对…')

        t = threading.Thread(target=self._worker, args=(pairs,), daemon=True)
        t.start()

    def _worker(self, pairs):
        try:
            n = len(pairs)
            results = []
            for i, (s, l) in enumerate(pairs, 1):
                def log_cb(msg, _i=i, _n=n):
                    self.msg_q.put(('log', f'[第{_i}/{_n}对] {msg}'))
                self.msg_q.put(('log', f'\n========== 开始核对第 {i}/{n} 对 =========='))
                summ = run_reconcile(s, l, None, log=log_cb)
                results.append(summ)
            self.msg_q.put(('done', results))
        except Exception as e:
            import traceback
            self.msg_q.put(('log', f'\n[出错] {e}\n'))
            self.msg_q.put(('log', traceback.format_exc()))
            self.msg_q.put(('error', str(e)))

    # ---------- 消息轮询 (主线程) ----------
    def _poll(self):
        try:
            while True:
                kind, payload = self.msg_q.get_nowait()
                if kind == 'log':
                    self._append_log(payload)
                elif kind == 'done':
                    self._on_done(payload)
                elif kind == 'error':
                    self._on_error(payload)
        except queue.Empty:
            pass
        # 持续轮询：任务结束后也要继续，方便下一次点击开始核对
        self.root.after(150, self._poll)

    def _append_log(self, text):
        self.log_text.config(state='normal')
        self.log_text.insert('end', text + '\n')
        self.log_text.see('end')
        self.log_text.config(state='disabled')
        self.root.update_idletasks()

    def _on_done(self, summaries):
        self.running = False
        self.summaries = summaries
        self.btn_start.config(state='normal')
        self.btn_open.config(state='normal')
        self.btn_folder.config(state='normal')
        self.status_label.config(text=f'核对完成（{len(summaries)} 份结果）', foreground='#1a7f37')
        self.root.title('银行对账核对工具 — 核对完成')
        lines = []
        for i, s in enumerate(summaries, 1):
            lines.append(f'{i}. {s["output_file"]}\n'
                         f'   对账单 {s["stmt_count"]} 条 / 日记账 {s["ledger_count"]} 条，'
                         f'交叉核对 {s["cross_groups"]} 组，日汇总 {s.get("daily_groups", 0)} 组')
        messagebox.showinfo('核对完成',
                            f'共生成 {len(summaries)} 份对账结果：\n\n' + '\n'.join(lines))

    def _on_error(self, msg):
        self.running = False
        self.btn_start.config(state='normal')
        self.status_label.config(text='核对失败', foreground='#d1242f')
        self.root.title('银行对账核对工具 — 核对失败')
        messagebox.showerror('核对失败', f'运行出错:\n{msg}')

    # ---------- 打开结果 ----------
    def _open_result(self):
        if not self.summaries:
            return
        self._open_path(self.summaries[-1]['output_file'])

    def _open_folder(self):
        if not self.summaries:
            return
        out = self.summaries[-1]['output_file']
        if sys.platform.startswith('win'):
            subprocess.Popen(f'explorer /select,"{os.path.normpath(out)}"')
        else:
            self._open_path(os.path.dirname(out))

    @staticmethod
    def _open_path(path):
        try:
            if sys.platform.startswith('win'):
                os.startfile(path)  # noqa
            elif sys.platform == 'darwin':
                subprocess.run(['open', path])
            else:
                subprocess.run(['xdg-open', path])
        except Exception:
            messagebox.showerror('无法打开', f'打开失败:\n{path}')


def launch_gui():
    """启动图形界面（无命令行参数时调用）。"""
    if not _TK_OK:
        print('当前环境缺少 tkinter，无法启动图形界面，请改用命令行模式。')
        sys.exit(1)
    if sys.platform.startswith('win'):
        try:  # 高 DPI 下界面更清晰
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = tk.Tk()
    BankReconApp(root)
    root.mainloop()


# ============================================================
# 入口：带参数 → 命令行；无参数 → 图形界面
# ============================================================

if __name__ == '__main__':
    if len(sys.argv) > 1:
        cli_main()
    else:
        launch_gui()
