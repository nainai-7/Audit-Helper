# -*- coding: utf-8 -*-
"""
调整分录汇总工具（支持多工作表 / 多文件分别合并）
=============================================

功能：
    - 选一个文件夹，勾选要汇总的工作表（可多选，如「调整分录」「凭证」）
    - 每个工作表【单独指定】表头所在行（1-based），不同 sheet 的表头行可以不同
    - 每个工作表【分别汇总】，各自输出一个 Excel 文件：
        <前缀><工作表名>_汇总表.xlsx
    - 同一工作表的所有 Excel 按【列名对齐】合并：
        同名列上下合并，不同列并排保留（补空）
    - 用「文件名」列标记每条数据来自哪个文件
    - 自动过滤空白行；遇到「合计」行时只保留合计行之前的明细，合计行及其下方内容全部舍去
    - 【可选】保留表头上方行：勾选后，表头所在行之前的各行（如第 1–7 行）一并保留在汇总表顶部，
      表头仍位于原行号，格式与源文件保持一致（多文件时仅以第一个文件的上方行为模板）

用法：
    GUI 模式（默认，推荐）：  python Combine.py
    命令行模式：              python Combine.py <folder> <output_dir> <spec> [prefix] [keep_leading]
                              spec 形如 调整分录:1,凭证:3 （冒号后是表头行，省略默认1）
                              keep_leading 可选：1/yes/keep 表示保留表头上方行，省略默认不保留
                              例如：python Combine.py "D:\\数据" "D:\\out" "调整分录:8" "" keep
"""

import os
import re
import sys
import threading
import queue

import pandas as pd

# 读写引擎统一使用 openpyxl：
# 1) 读时不做类型推断（dtype=object），原样保留单元格里的文本/数字
#    （如带前导零的编号 00123 不会被 calamine 推断成数字 123）；
# 2) 写时用 openpyxl 后端，字符串原样写出，不会丢失前导零等文本格式。
# 注意：pandas 写入不复制单元格样式（加粗/底色/数字格式等），仅保证「值」原样保留；
#       若需完整保留样式，需要改用直接复制单元格的方案（复杂度高、性能更慢）。
import openpyxl  # 确保可用；项目 venv 已安装

_READ_ENGINE = 'openpyxl'
_WRITE_ENGINE = 'openpyxl'


# ============================================================
# 公共工具
# ============================================================

_INVALID = re.compile(r'[\\/:*?"<>|]')


def _safe_name(s, limit=31):
    """把工作表名转成合法的工作表名（去掉文件名非法字符，超长截断）。"""
    s = _INVALID.sub('_', s).strip()
    return s[:limit] if s else 'sheet'


def _clean_and_filter(df):
    """过滤空白行；定位首个含「合计」的行，只保留该行之前的明细，
    合计行及其下方所有内容（合计、附注等）一律舍去。返回处理后的 df。"""
    df = df.reset_index(drop=True)  # 先重置索引，保证后续 iloc 按位置切片正确
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    # 找到第一个含「合计」字样的行，截断到它之前
    total_rows_mask = df.apply(
        lambda row: row.astype(str).str.contains('合计').any(), axis=1)
    total_rows_idx = df[total_rows_mask].index.tolist()
    if total_rows_idx:
        first = total_rows_idx[0]
        df = df.iloc[:first].reset_index(drop=True)
    # 去掉截断后残留的整行空白
    df = df.dropna(how='all').reset_index(drop=True)
    return df


def _dedup_columns(df, log=None):
    """确保列名唯一。遇到重复列名时追加 _2、_3 … 后缀。
    不同 DataFrame 列名不一致时 concat 会触发 pandas 的
    'Reindexing only valid with uniquely valued Index objects' 错误，
    所以每个文件解析后都强制去重一次。返回新的 DataFrame，不修改入参。"""
    cols = list(df.columns)
    seen = {}
    new_cols = []
    renames = []
    for c in cols:
        if c in seen:
            seen[c] += 1
            new_name = f'{c}_{seen[c]}'
            new_cols.append(new_name)
            renames.append((c, new_name))
        else:
            seen[c] = 0
            new_cols.append(c)
    if renames and log is not None:
        # 控制日志长度，最多列出前 8 个
        head = ', '.join(f'{a}->{b}' for a, b in renames[:8])
        more = f'（…另有 {len(renames) - 8} 个）' if len(renames) > 8 else ''
        log(f'    ⚠ 表头有 {len(renames)} 个重复列名，已自动重命名: {head}{more}')
    out = df.copy()
    out.columns = new_cols
    return out




# ============================================================
# 核心处理逻辑
# ============================================================

def run_task(folder_path, sheet_headers, output_dir, file_prefix='',
              keep_leading=False, log=None, progress=None, cancel=None):
    """执行多工作表【分别】合并汇总。

    Args:
        folder_path:   源文件夹路径（包含待汇总的 Excel 文件）
        sheet_headers: 字典 {工作表名: 表头行(1-based)}，每个 sheet 单独指定表头行
        output_dir:    汇总结果输出文件夹
        file_prefix:   输出文件名前缀（可选）
        keep_leading:  是否保留表头上方行（格式不变）；True 时把每个 sheet 表头之前的行
                       一并保留在汇总表顶部（多文件仅以第一个文件为模板）
        log:           日志回调 log(str)；None 时 print
        progress:      进度回调 progress(done, total)；None 时不调用

    Returns:
        list[dict]: 每个元素 {'sheet', 'output_file', 'rows'}；无有效数据返回 None
    """
    if log is None:
        log = print

    def _cancelled():
        return cancel is not None and cancel.is_set()

    if not folder_path or not os.path.isdir(folder_path):
        log(f'错误：源文件夹不存在：{folder_path}')
        return None
    if not sheet_headers:
        log('错误：未选择任何工作表')
        return None

    all_excels = [f for f in os.listdir(folder_path)
                  if f.lower().endswith(('.xlsx', '.xlsm', '.xls')) and not f.startswith('~$')]
    if not all_excels:
        log('错误：文件夹中没有 Excel 文件（仅支持 .xlsx / .xlsm）。')
        return None
    # 旧版 .xls（BIFF）openpyxl 读不了，单独跳过并提示
    xls_files = [f for f in all_excels if f.lower().endswith('.xls')]
    files = [f for f in all_excels if not f.lower().endswith('.xls')]
    if not files:
        log(f'错误：文件夹里只有 {len(xls_files)} 个旧版 .xls 文件，openpyxl 无法读取。'
            f'请先把它们另存为 .xlsx 或 .xlsm 再运行。')
        return None
    if xls_files:
        log(f'提示：已忽略 {len(xls_files)} 个旧版 .xls 文件（openpyxl 不支持），仅处理 .xlsx / .xlsm。')

    os.makedirs(output_dir, exist_ok=True)
    log(f'源文件夹「{folder_path}」共 {len(files)} 个 Excel 文件')
    log(f'将分别汇总 {len(sheet_headers)} 个工作表，输出到「{output_dir}」')

    results = []
    total_steps = len(files) * len(sheet_headers)
    done = 0

    # 以「工作表」为主循环：每个 sheet 单独汇总、单独输出一个文件
    for sheet, header_row in sheet_headers.items():
        if _cancelled():
            log('已取消：在处理下一工作表前退出。')
            break
        hr = max(0, int(header_row) - 1)  # pandas 的 header 是 0-based
        all_data = []
        first_leading = None   # 表头上方行（取第一个成功读取的文件）
        log(f'\n===== 工作表「{sheet}」（表头行：第 {header_row} 行）=====')

        for filename in sorted(files):
            if _cancelled():
                log('已取消：在处理下一文件前退出。')
                break
            file_name = os.path.splitext(filename)[0]
            file_path = os.path.join(folder_path, filename)
            done += 1
            if progress is not None:
                progress(done - 1, total_steps)

            try:
                xl = pd.ExcelFile(file_path)
            except Exception as e:
                log(f'  · [跳过] {filename} 无法打开：{e}')
                continue

            if sheet not in xl.sheet_names:
                log(f'  · {filename} 没有工作表「{sheet}」，跳过')
                continue
            try:
                # 一次性读全表（header=None），再从表头行切分，避免二次读取
                # 统一用 openpyxl 引擎，dtype=object 保证值原样保留（含前导零文本）
                raw = pd.read_excel(
                    file_path, sheet_name=sheet, header=None,
                    dtype=object, keep_default_na=False, engine=_READ_ENGINE,
                )
            except Exception as e:
                log(f'  · 读取 {filename}「{sheet}」失败：{e}')
                continue

            nrows = raw.shape[0]
            if nrows <= hr:
                log(f'  · {filename}「{sheet}」行数不足，跳过')
                continue
            header_vals = [str(c) for c in raw.iloc[hr].tolist()]

            # 表头上方行（第 1…hr 行）+ 表头行本身（第 hr+1 行）一并保留为模板，
            # 使汇总表格式与源文件一致：表头仍位于原行号
            if keep_leading and first_leading is None:
                leading = raw.iloc[:hr + 1].copy()
                leading.columns = header_vals
                leading = _dedup_columns(leading)  # 模板列名也去重
                first_leading = leading

            # 数据：表头行之后
            df = raw.iloc[hr + 1:].copy()
            df.columns = header_vals
            df = _clean_and_filter(df)
            if df.empty:
                log(f'  · {filename}「{sheet}」无有效数据，跳过')
                continue

            # 标记出处：文件名（来源工作表对本文件恒定，省略）
            df = _dedup_columns(df, log=log)  # 防重复列名触发 concat InvalidIndexError
            df.insert(0, '文件名', file_name)
            all_data.append(df)
            log(f'  · {filename}「{sheet}」已加入（{len(df)} 行）')

        if progress is not None:
            progress(done, total_steps)

        if not all_data:
            log(f'  → 「{sheet}」无有效数据，未生成文件')
            continue

        # 合并：pandas 自动按列名对齐（同名列上下合并，不同列并排）
        merged_df = pd.concat(all_data, ignore_index=True)

        # 重新生成连续「序号」（仅当存在该列；多文件各自从 1 开始时统一重排）
        if '序号' in merged_df.columns:
            merged_df['序号'] = range(1, len(merged_df) + 1)

        # 保留表头上方行（格式不变）：把模板行拼到数据最前面
        if keep_leading and first_leading is not None:
            first_leading = first_leading.copy()
            first_leading.insert(0, '文件名', '')
            # 让表头行（第 hr 行，0-based）的「文件名」列也显示「文件名」，否则该列无表头
            first_leading.iloc[hr, 0] = '文件名'
            merged_df = pd.concat([first_leading, merged_df], ignore_index=True)

        # 不同文件列宽微差时补空，避免写出 NaN
        merged_df = merged_df.fillna('')

        # 保留表头上方行时，源文件表头区（标题/副标题/表头）应坐在汇总表顶部，
        # 不能让 pandas 再把列名写成第 1 行（否则会整体下挤一行，表头错位）。
        # 故 keep_leading 时 header=False；非 keep_leading 时保留 pandas 列名作为表头。
        write_header = not keep_leading

        # 输出：每个 sheet 一个独立文件。
        # 若目标被占用（如 Excel 正打开同名文件），自动尝试 _2/_3… 序号后缀，
        # 确保能写出而不中断整个任务。
        base_name = f'{file_prefix}{sheet}_汇总表.xlsx'
        out_path = os.path.join(output_dir, base_name)
        stem, ext = os.path.splitext(out_path)
        written = False
        # 写引擎统一用 openpyxl（原样保留值；不复制样式）
        write_engine = _WRITE_ENGINE
        for seq in range(1, 100):
            cand = out_path if seq == 1 else f'{stem}_{seq}{ext}'
            try:
                with pd.ExcelWriter(cand, engine=write_engine, mode='w') as writer:
                    merged_df.to_excel(
                        writer, sheet_name=_safe_name(sheet), index=False, header=write_header,
                    )
                out_path = cand
                written = True
                if seq > 1:
                    log(f'  · 原文件名被占用，已写出到：{os.path.basename(cand)}')
                break
            except PermissionError:
                log(f'  · 「{os.path.basename(cand)}」被占用，换一个名称重试…')
                continue
        if not written:
            log(f'  · [权限错误] 无法写出汇总表（磁盘满 / 目录只读 / 所有候选名均被占用）。')
            log(f'     请关闭相关 Excel 或换个输出目录后重试。')
            return None

        log(f'  → 已保存：{out_path}（{len(merged_df)} 行）')
        results.append({'sheet': sheet, 'output_file': out_path, 'rows': len(merged_df)})

    if progress is not None:
        progress(total_steps, total_steps)

    if not results:
        log('\n未生成任何汇总文件，请检查输入。')
        return None

    log('\n全部完成！共生成 %d 个汇总文件：' % len(results))
    for r in results:
        log(f"  · 「{r['sheet']}」-> {r['output_file']}（{r['rows']} 行）")
    return results


# ============================================================
# 图形界面（GUI）
# ============================================================

def launch_gui():
    """启动调整分录汇总工具的图形界面。"""
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from tkinter.scrolledtext import ScrolledText

    try:  # 高 DPI 清晰化
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    root.title('调整分录汇总工具')
    root.minsize(560, 480)
    root.resizable(True, True)   # 窗口可拉伸；宽度/高度按内容自适应

    outer = ttk.Frame(root, padding=12)
    outer.pack(fill='both', expand=True)

    # ---------- 源文件夹 ----------
    ttk.Label(outer, text='源文件夹：').grid(row=0, column=0, sticky='e', padx=(0, 6), pady=4)
    var_folder = tk.StringVar()
    ttk.Entry(outer, textvariable=var_folder).grid(row=0, column=1, sticky='ew', padx=(0, 6))

    def pick_folder():
        d = filedialog.askdirectory(title='选择源文件夹（包含要汇总的 Excel）')
        if d:
            var_folder.set(d)
            load_sheets()

    ttk.Button(outer, text='浏览…', width=8, command=pick_folder).grid(row=0, column=2)

    # ---------- 各子函数定义（必须先于引用它们的控件，否则触发 UnboundLocalError） ----------
    sheet_header_values = {}   # sheet -> 上一次提交的表头行
    sheet_header_vars = {}     # sheet -> IntVar（当前控件）

    def rebuild_header_cfg():
        # 先把当前控件里的值提交下来，避免切换选择时丢失
        for s, var in sheet_header_vars.items():
            try:
                sheet_header_values[s] = int(var.get())
            except Exception:
                pass
        for child in hdr_cfg_frame.winfo_children():
            child.destroy()
        sheet_header_vars.clear()
        selected = [listbox.get(i) for i in listbox.curselection()]
        if not selected:
            ttk.Label(hdr_cfg_frame, text='（请先在上方选择工作表）',
                      foreground='#808080').pack(anchor='w')
            return
        for s in selected:
            row = ttk.Frame(hdr_cfg_frame)
            row.pack(fill='x', pady=2)
            ttk.Label(row, text=s, width=26, anchor='w').pack(side='left')
            ttk.Label(row, text='表头行：').pack(side='left')
            v = tk.IntVar(value=sheet_header_values.get(s, 1))
            sheet_header_vars[s] = v
            ttk.Spinbox(row, from_=1, to=50, textvariable=v, width=6).pack(side='left')
            ttk.Label(row, text='（第 N 行为表头）', foreground='#808080').pack(side='left')

    def load_sheets():
        """扫描文件夹里的第一份 Excel，把所有 Sheet 名填入列表框。"""
        folder = var_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            return
        try:
            first_xlsx = None
            for f in sorted(os.listdir(folder)):
                if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$'):
                    first_xlsx = os.path.join(folder, f)
                    break
            if not first_xlsx:
                messagebox.showinfo('提示', '所选文件夹中没有 Excel 文件。')
                return
            xl = pd.ExcelFile(first_xlsx)
            names = list(xl.sheet_names)
            listbox.delete(0, tk.END)
            for n in names:
                listbox.insert(tk.END, n)
            if '调整分录' in names:
                listbox.selection_set(names.index('调整分录'))
            elif names:
                listbox.selection_set(0)
            rebuild_header_cfg()
        except Exception as exc:
            messagebox.showerror('读取失败', f'无法读取该 Excel：\n{exc}')

    # ---------- 工作表（多选 Listbox） ----------
    ttk.Label(outer, text='工作表（可多选）：').grid(
        row=1, column=0, sticky='ne', padx=(0, 6), pady=4)
    sheet_frame = ttk.Frame(outer)
    sheet_frame.grid(row=1, column=1, columnspan=2, sticky='ew', padx=(0, 6), pady=4)
    listbox = tk.Listbox(sheet_frame, selectmode='multiple', height=5,
                         exportselection=False, font=('Microsoft YaHei UI', 10))
    sb_sheet = ttk.Scrollbar(sheet_frame, orient='vertical', command=listbox.yview)
    listbox.configure(yscrollcommand=sb_sheet.set)
    listbox.pack(side='left', fill='both', expand=True)
    sb_sheet.pack(side='right', fill='y')
    ttk.Button(outer, text='刷新', width=8, command=load_sheets).grid(
        row=1, column=3, padx=(6, 0), sticky='n')
    ttk.Label(outer,
              text='选中文件夹后自动列出；按住 Ctrl 可多选，Shift 连选；默认勾选「调整分录」',
              foreground='#808080').grid(row=2, column=1, columnspan=3, sticky='w', pady=(0, 4))

    # ---------- 各工作表表头行（动态生成） ----------
    ttk.Label(outer, text='各表表头行：').grid(row=3, column=0, sticky='ne', padx=(0, 6), pady=4)
    hdr_outer = ttk.LabelFrame(outer, text='每个工作表的表头所在行（可不同）', padding=6)
    hdr_outer.grid(row=3, column=1, columnspan=3, sticky='ew', padx=(0, 6), pady=4)
    hdr_cfg_frame = ttk.Frame(hdr_outer)
    hdr_cfg_frame.pack(fill='x')
    ttk.Label(hdr_cfg_frame, text='（请在上方选择工作表后，在此设置各自的表头行）',
              foreground='#808080').pack(anchor='w')

    # 是否保留表头上方行（格式不变）
    var_keep_leading = tk.BooleanVar(value=False)
    ttk.Checkbutton(
        hdr_outer, text='保留表头上方行（格式不变）', variable=var_keep_leading
    ).pack(anchor='w', pady=(6, 0))
    ttk.Label(
        hdr_outer,
        text='勾选后，表头之前的各行（如第 1–7 行）会一并保留在汇总表顶部，表头仍位于原行号；'
             '多文件时以第一个文件的上方行为模板。',
        foreground='#808080').pack(anchor='w', pady=(0, 2))

    listbox.bind('<<ListboxSelect>>', lambda e: rebuild_header_cfg())

    # ---------- 输出文件夹 + 文件名前缀 ----------
    ttk.Label(outer, text='输出文件夹：').grid(row=4, column=0, sticky='e', padx=(0, 6), pady=4)
    var_outdir = tk.StringVar()
    ttk.Entry(outer, textvariable=var_outdir).grid(row=4, column=1, sticky='ew', padx=(0, 6))

    def pick_outdir():
        d = filedialog.askdirectory(title='选择输出文件夹')
        if d:
            var_outdir.set(d)

    ttk.Button(outer, text='浏览…', width=8, command=pick_outdir).grid(row=4, column=2)

    ttk.Label(outer, text='文件名前缀：').grid(row=5, column=0, sticky='e', padx=(0, 6), pady=4)
    var_prefix = tk.StringVar()
    ttk.Entry(outer, textvariable=var_prefix).grid(row=5, column=1, sticky='ew', padx=(0, 6))
    ttk.Label(outer, text='可留空；输出命名为「前缀 + 工作表名 + _汇总表.xlsx」',
              foreground='#808080').grid(row=5, column=2, columnspan=2, sticky='w', padx=(6, 0))

    outer.columnconfigure(1, weight=1)

    # ---------- 进度条 ----------
    pv = tk.DoubleVar()
    bar = ttk.Progressbar(outer, variable=pv, maximum=100)
    bar.grid(row=6, column=0, columnspan=4, sticky='ew', pady=(10, 6))

    # ---------- 日志区 ----------
    ttk.Label(outer, text='运行日志').grid(row=7, column=0, columnspan=4, sticky='w')
    log_text = ScrolledText(outer, height=12, state='disabled', wrap='word',
                            font=('Consolas', 10))
    log_text.grid(row=8, column=0, columnspan=4, sticky='nsew', pady=(4, 8))
    outer.rowconfigure(8, weight=1)

    q = queue.Queue()
    log_lines = []  # 收集日志，用于失败时显示真实原因
    cancel_flag = threading.Event()  # 用户可随时中止正在运行的任务

    def log(msg):
        msg = str(msg)
        log_lines.append(msg)
        q.put(('log', msg))
        try:  # 同时输出到控制台（PyCharm 运行配置下可见），方便排查
            print(msg)
        except Exception:
            pass

    def progress(done, total):
        if total > 0:
            q.put(('prog', done * 100.0 / total))

    def poll():
        try:
            while True:
                kind, payload = q.get_nowait()
                if kind == 'log':
                    log_text.config(state='normal')
                    log_text.insert('end', payload + '\n')
                    # 行数封顶：避免日志无限增长拖慢界面（保留最近 2000 行）
                    if float(log_text.index('end-1c').split('.')[0]) > 2000:
                        log_text.delete('1.0', '500.0')
                    log_text.see('end')
                    log_text.config(state='disabled')
                elif kind == 'prog':
                    pv.set(payload)
                elif kind == 'done':
                    btn_run.config(state='normal')
                    btn_cancel.config(state='disabled')
                    cancel_flag.clear()
                    if payload is None:
                        detail = '\n'.join(log_lines[-12:]) if log_lines else '（无日志）'
                        messagebox.showwarning(
                            '未完成',
                            '未生成任何汇总文件。常见原因：\n'
                            '· 勾选的工作表名与文件里的不一致（含空格/全半角）\n'
                            '· 表头行填错（导致把标题当成数据）\n'
                            '· 源文件是旧版 .xls（openpyxl 不支持）或 .xlsm 未识别\n\n'
                            '最近日志（或见 PyCharm 控制台）：\n' + detail)
                    else:
                        lines = '\n'.join(
                            f"· 「{r['sheet']}」 {r['output_file']}（{r['rows']} 行）"
                            for r in payload)
                        if messagebox.askyesno(
                            '完成',
                            f'汇总完成！共生成 {len(payload)} 个文件：\n\n{lines}\n\n'
                            f'是否打开生成的汇总文件？'
                        ):
                            try:
                                if sys.platform.startswith('win'):
                                    for r in payload:
                                        os.startfile(r['output_file'])
                            except Exception:
                                pass
        except queue.Empty:
            pass
        root.after(120, poll)

    def collect_cfg():
        folder = var_folder.get().strip()
        outdir = var_outdir.get().strip()
        prefix = var_prefix.get().strip()
        # 提交当前表头控件的值
        for s, var in sheet_header_vars.items():
            try:
                sheet_header_values[s] = int(var.get())
            except Exception:
                pass
        sheets = [listbox.get(i) for i in listbox.curselection()]
        if not folder or not os.path.isdir(folder):
            raise ValueError('请选择有效的源文件夹')
        if not sheets:
            raise ValueError('请至少选择一个工作表')
        if not outdir:
            raise ValueError('请选择输出文件夹')
        sheet_headers = {s: sheet_header_values.get(s, 1) for s in sheets}
        return folder, sheet_headers, outdir, prefix, var_keep_leading.get()

    def worker(cfg):
        try:
            result = run_task(cfg[0], cfg[1], cfg[2], cfg[3],
                              keep_leading=cfg[4], log=log, progress=progress,
                              cancel=cancel_flag)
            q.put(('done', result))
        except Exception as e:
            log(f'[错误] {e}')
            q.put(('done', None))

    def on_run():
        try:
            cfg = collect_cfg()
        except Exception as e:
            messagebox.showerror('参数有误', str(e))
            return
        log_text.config(state='normal')
        log_text.delete('1.0', 'end')
        log_text.config(state='disabled')
        pv.set(0)
        btn_run.config(state='disabled')
        btn_cancel.config(state='normal')
        cancel_flag.clear()
        threading.Thread(target=worker, args=(cfg,), daemon=True).start()

    def on_cancel():
        cancel_flag.set()
        btn_cancel.config(state='disabled')
        log('用户请求取消，将在处理完当前文件/工作表后退出…')

    btn_run = ttk.Button(outer, text='开始汇总', command=on_run)
    btn_run.grid(row=9, column=0, columnspan=3, pady=(0, 4), ipadx=20, ipady=4)
    btn_cancel = ttk.Button(outer, text='取消', command=on_cancel, state='disabled')
    btn_cancel.grid(row=9, column=3, pady=(0, 4), ipadx=12, ipady=4, sticky='e')

    # 窗口尺寸按内容自适应（宽度尤其随提示/勾选项变化），并封顶屏幕高度 90%
    root.update_idletasks()
    req_w = outer.winfo_reqwidth() + 28
    req_h = min(outer.winfo_reqheight() + 28, int(root.winfo_screenheight() * 0.9))
    root.geometry(f'{req_w}x{req_h}')

    root.after(120, poll)
    root.mainloop()


# ============================================================
# 入口
# ============================================================

if __name__ == '__main__':
    # 命令行模式：python Combine.py <folder> <output_dir> <spec> [prefix]
    #   spec 形如  调整分录:1,凭证:3   （冒号后是表头行，省略默认 1）
    if len(sys.argv) >= 4:
        folder = sys.argv[1]
        output_dir = sys.argv[2]
        spec = sys.argv[3]
        prefix = sys.argv[4] if len(sys.argv) > 4 else ''
        keep_arg = sys.argv[5] if len(sys.argv) > 5 else ''
        keep_leading = str(keep_arg).strip().lower() in ('1', 'yes', 'y', 'true', 'keep', '保留')
        sheet_headers = {}
        for part in spec.split(','):
            part = part.strip()
            if not part:
                continue
            if ':' in part:
                name, _, hr = part.rpartition(':')
                name = name.strip()
                try:
                    row = int(hr.strip())
                except Exception:
                    row = 1
            else:
                name = part
                row = 1
            if name:
                sheet_headers[name] = row
        if not sheet_headers:
            print('错误：未解析到任何工作表', file=sys.stderr)
            sys.exit(1)
        try:
            result = run_task(folder, sheet_headers, output_dir, prefix,
                              keep_leading=keep_leading)
            if result is None:
                sys.exit(1)
        except Exception as e:
            print(f'错误: {e}', file=sys.stderr)
            sys.exit(1)
    else:
        launch_gui()
