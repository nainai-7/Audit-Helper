# -*- coding: utf-8 -*-
"""
人本股份 征信报告PDF批量核对（本地离线版 · GUI）
=================================================
功能：
1. GUI选择：PDF文件夹、Excel输出路径
2. 仅识别【未结清信贷及授信信息概要】→【相关还款责任信息概要】固定区间
3. 只取每行最右侧"合计余额"列（9列结构），不取账户数、不取合计行
   科目：短期借款/中长期借款/循环透支/银行承兑汇票(兼容折行)/信用证/贴现/银行保函
4. 全部结清（报告无未结清章节）的企业按0填列
5. 生成Excel：公司名称 + 7个金额列（8列），表头不换行、列宽按内容自适应

运行（全程离线，无任何联网调用）：
  python main.py
依赖：pdfplumber, openpyxl（tkinter为Python自带）
"""
import os
import re
import sys
import threading
import queue

import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ================= 常量 =================
# 行标签白名单（已去空白，兼容"银行承兑汇 票"折行）
SUBJECTS = {
    '短期借款':  '短期借款',
    '中长期借款': '中长期借款',
    '循环透支':  '循环透支',
    '银行承兑汇票': '银行承兑汇票',
    '信用证':    '信用证',
    '贴现':      '已贴现',
    '已贴现':    '已贴现',
    '银行保函':  '银行保函',
    '保函':      '银行保函',
}

HEADERS = ['公司名称',
           '短期借款金额（万元）', '长期借款金额（万元）', '循环透支（万元）',
           '应付票据金额（万元）', '信用证', '已贴现未到期金额（万元）', '保函（万元）']
KEYS = ['短期借款', '中长期借款', '循环透支', '银行承兑汇票', '信用证', '已贴现', '银行保函']


# ================= 工具函数 =================
def parse_num(s):
    """数值清洗：去千分位；--、/、空按0；非数字返回None"""
    if s is None:
        return None
    t = str(s).strip().replace(',', '').replace('，', '')
    if t in ('', '--', '-', '/'):
        return 0.0
    try:
        return float(t)
    except ValueError:
        return None


def clean(s):
    """去空格换行，用于标题/标签匹配"""
    return re.sub(r'\s+', '', s or '')


def visual_len(s):
    """按显示宽度估算：中文字符算2，其他算1"""
    n = 0
    for ch in str(s):
        n += 2 if '\u4e00' <= ch <= '\u9fff' or ch in '（）【】，。：' else 1
    return n


# ================= 核心处理 =================
def process_pdf(path):
    """处理单份PDF，返回 (公司名称, 科目余额dict)"""
    with pdfplumber.open(path) as pdf:
        # 公司名称：从首页"企业名称"行提取
        p0 = pdf.pages[0].extract_text() or ''
        m = re.search(r'企业名称[：:\s]*([^\n]+)', p0)
        name = m.group(1).strip() if m else os.path.basename(path)
        if name.endswith('分公'):
            name += '司'

        # 前4页搜标题，锁定区间
        head = ''.join((pdf.pages[i].extract_text() or '') for i in range(min(4, len(pdf.pages))))
        hc = clean(head)
        vals = {}

        if '未结清信贷及授信信息概要' in hc:
            # 起点：标题所在页
            sp = next(i for i, p in enumerate(pdf.pages)
                      if '未结清信贷及授信信息概要' in clean(p.extract_text()))
            # 终点：下一个标题（相关还款责任 / 已结清信贷，以后出现者为准）
            ep = len(pdf.pages)
            for i in range(sp, len(pdf.pages)):
                tc = clean(pdf.pages[i].extract_text())
                if i > sp and ('相关还款责任信息概要' in tc or '已结清信贷信息概要' in tc):
                    ep = i
                    break
            # 区间内提取表格
            for i in range(sp, ep):
                for tb in pdf.pages[i].extract_tables():
                    for row in tb:
                        if not row or not row[0]:
                            continue
                        lab = clean(str(row[0]))
                        if lab == '合计' or lab not in SUBJECTS:
                            continue           # 排除合计行与白名单外科目
                        if len(row) < 9:
                            continue           # 只认9列汇总表，天然排除已结清5列表
                        v = parse_num(row[8])  # 只取末列"合计余额"
                        if v is None:
                            continue
                        key = SUBJECTS[lab]
                        vals[key] = vals.get(key, 0.0) + v
        # else：全部结清企业（报告无未结清章节）→ vals保持空，按0填列
    return name, vals


def run_task(pdf_dir, out_xlsx, log, progress):
    """后台线程执行：提取 → 写Excel"""
    try:
        pdf_files = sorted(f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf'))
        if not pdf_files:
            log('错误：所选文件夹中没有PDF文件')
            return
        log('共发现 %d 份PDF，开始处理…' % len(pdf_files))

        companies = {}    # 公司名称 -> {'files': [...], 'vals': {...}}
        errors = []

        for idx, f in enumerate(pdf_files, 1):
            try:
                name, vals = process_pdf(os.path.join(pdf_dir, f))
            except Exception as e:
                errors.append(f)
                log('  [失败] %s：%s' % (f, str(e)[:80]))
                progress(idx, len(pdf_files))
                continue
            if name not in companies:
                companies[name] = {'files': [f], 'vals': vals}
            else:
                companies[name]['files'].append(f)
                for k, v in vals.items():
                    if v and not companies[name]['vals'].get(k):
                        companies[name]['vals'][k] = v   # 同名企业多份报告取非零值
            progress(idx, len(pdf_files))
            if idx % 10 == 0 or idx == len(pdf_files):
                log('  已处理 %d / %d' % (idx, len(pdf_files)))

        # ---- 写Excel（全新文件）----
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '征信报告核对'
        thin = Side(style='thin', color='999999')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        ws.append(HEADERS)
        for c in ws[1]:
            c.font = Font(bold=True, size=11)
            c.fill = PatternFill('solid', fgColor='DDEBF7')
            c.alignment = Alignment(horizontal='center', vertical='center')  # 不换行
            c.border = border

        for name, info in sorted(companies.items(), key=lambda kv: kv[1]['files'][0]):
            vals = info['vals']
            ws.append([name] + [vals.get(k, 0.0) for k in KEYS])

        for r in range(2, ws.max_row + 1):
            for c in range(1, len(HEADERS) + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = border
                if c >= 2:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='right')
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

        # ---- 列宽按内容自适应（表头+数据取最大显示宽度，加边距）----
        for c in range(1, len(HEADERS) + 1):
            col = get_column_letter(c)
            w = 0
            for r in range(1, ws.max_row + 1):
                v = ws.cell(row=r, column=c).value
                if v is None:
                    continue
                if c >= 2 and isinstance(v, (int, float)):
                    text = format(v, ',.2f')
                else:
                    text = str(v)
                w = max(w, visual_len(text))
            ws.column_dimensions[col].width = w + 3   # 边距

        ws.freeze_panes = 'A2'
        wb.save(out_xlsx)
        if errors:
            log('完成：Excel已保存 %s（有 %d 份解析失败，见上方日志）' % (out_xlsx, len(errors)))
        else:
            log('全部完成：公司 %d 家，失败 0 份。Excel已保存：%s' % (len(companies), out_xlsx))
    except Exception as e:
        log('发生错误：%s' % e)


# ================= GUI =================
class App:
    def __init__(self, root):
        self.root = root
        root.title('征信报告PDF批量核对（本地离线）')
        root.geometry('760x420')
        root.minsize(660, 380)

        main = ttk.Frame(root, padding=12)
        main.pack(fill='both', expand=True)

        # --- 路径选择区（初始为空，不预填路径）---
        frm = ttk.LabelFrame(main, text='配置', padding=10)
        frm.pack(fill='x')

        self.pdf_dir = tk.StringVar()
        self.out_xlsx = tk.StringVar()

        self._path_row(frm, 0, 'PDF文件夹：', self.pdf_dir, self.browse_dir)
        self._path_row(frm, 1, 'Excel输出：', self.out_xlsx, self.browse_save)

        # --- 运行按钮 ---
        self.btn = ttk.Button(main, text='开始处理', command=self.start)
        self.btn.pack(pady=(12, 4))

        # --- 进度条 ---
        self.pv = tk.DoubleVar()
        self.bar = ttk.Progressbar(main, variable=self.pv, maximum=100)
        self.bar.pack(fill='x', pady=(0, 8))

        # --- 日志 ---
        logf = ttk.LabelFrame(main, text='运行日志', padding=6)
        logf.pack(fill='both', expand=True)
        self.txt = tk.Text(logf, height=10, wrap='none', state='disabled',
                           font=('Microsoft YaHei', 9))
        self.txt.pack(fill='both', expand=True)

        # 后台线程 → 主线程的日志队列
        self.q = queue.Queue()
        root.after(100, self._poll)

    def _path_row(self, parent, r, label, var, browse_cmd):
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky='w', padx=(0, 6), pady=4)
        e = ttk.Entry(parent, textvariable=var)
        e.grid(row=r, column=1, sticky='ew', padx=(0, 6))
        ttk.Button(parent, text='浏览…', width=8, command=browse_cmd).grid(row=r, column=2)
        parent.columnconfigure(1, weight=1)

    # ---- 浏览按钮 ----
    def browse_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.pdf_dir.set(d)

    def browse_save(self):
        p = filedialog.asksaveasfilename(confirmoverwrite=False,
                                         initialfile='征信报告核对表.xlsx',
                                         defaultextension='.xlsx',
                                         filetypes=[('Excel文件', '*.xlsx')])
        if p:
            self.out_xlsx.set(p)

    # ---- 启动后台线程 ----
    def start(self):
        pdf_dir = self.pdf_dir.get().strip()
        out_xlsx = self.out_xlsx.get().strip()
        if not pdf_dir or not os.path.isdir(pdf_dir):
            messagebox.showerror('错误', '请选择有效的PDF文件夹')
            return
        if not out_xlsx:
            messagebox.showerror('错误', '请填写Excel输出路径')
            return
        if os.path.exists(out_xlsx):
            if not messagebox.askyesno('确认', 'Excel输出文件已存在，将覆盖保存。\n是否继续？'):
                return
        self.btn.config(state='disabled')
        self.pv.set(0)
        self.log('开始：PDF目录 %s' % pdf_dir)

        def log(msg):
            self.q.put(('log', msg))

        def progress(done, total):
            self.q.put(('prog', done * 100.0 / total))

        def worker():
            try:
                run_task(pdf_dir, out_xlsx, log, progress)
            finally:
                self.q.put(('done', None))

        threading.Thread(target=worker, daemon=True).start()

    # ---- 主线程轮询日志队列 ----
    def _poll(self):
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == 'log':
                    self.log(val)
                elif kind == 'prog':
                    self.pv.set(val)
                elif kind == 'done':
                    self.btn.config(state='normal')
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def log(self, msg):
        self.txt.config(state='normal')
        self.txt.insert('end', msg + '\n')
        self.txt.see('end')
        self.txt.config(state='disabled')


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
