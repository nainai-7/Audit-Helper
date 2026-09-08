# -*- coding: utf-8 -*-
"""AuditHelper - 审计小工具集合（统一启动器）

点击下方按钮即可启动对应子工具；每个子工具在自己的独立窗口中运行。
本启动器不修改任何子模块代码，仅作为入口聚合。

打包为单个 exe 后，子模块会被作为独立进程启动，互不影响。
"""

import os
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext


SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))


# ============================================================
# 功能清单
# ============================================================
FEATURES = [
    {
        'name': '银行对账核对',
        'desc': '对账单 ↔ 日记账交叉核对，自动生成 7-Sheet 对账结果（含日汇总兜底）。',
        'script': 'bankrecon.py',
    },
    {
        'name': '生成函证附件PDF',
        'desc': '按指定列分组，把每个分组导出为一份 PDF 表格（支持横/纵向、合计列、自适应列宽）。',
        'script': 'to_pdf.py',
    },
    {
        'name': '征信报告 PDF 批量核对',
        'desc': '从企业征信 PDF 中提取未结清信贷信息，输出 8 列 Excel 对照表。',
        'script': 'zhengxin.py',
    },
    {
        'name': '调整分录汇总',
        'desc': '把指定文件夹下所有 Excel 的「调整分录」工作表合并到一份总表（支持自由选文件夹/工作表/输出位置）。',
        'script': 'Combine.py',
    },
]


# ============================================================
# 子进程启动
# ============================================================

def launch_module(script_name, log):
    """在子进程中启动指定脚本（GUI 由各模块自带）。"""
    script_path = os.path.join(SCRIPTS_DIR, script_name)
    if not os.path.isfile(script_path):
        messagebox.showerror('启动失败', f'找不到模块文件：\n{script_path}')
        log(f'[错误] 找不到 {script_path}\n')
        return
    try:
        # 当前 Python 解释器 + 子模块脚本
        # 在打包场景：sys.executable 即为 exe 路径
        subprocess.Popen([sys.executable, script_path], cwd=SCRIPTS_DIR)
        log(f'已启动：{script_name}\n')
    except Exception as e:
        messagebox.showerror('启动失败', str(e))
        log(f'[错误] 启动 {script_name} 失败：{e}\n')


# ============================================================
# 主界面
# ============================================================

def main():
    try:  # 高 DPI 清晰化
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    root.title('AuditHelper - 审计小工具集合')
    root.geometry('760x600')          # 初值，稍后自适应内容高度
    root.minsize(540, 360)

    outer = ttk.Frame(root, padding=14)
    outer.pack(fill='both')

    # ---------- 标题 ----------
    ttk.Label(outer, text='AuditHelper',
              font=('Microsoft YaHei UI', 15, 'bold')).pack(anchor='w')
    ttk.Label(outer,
              text='点击下方按钮启动对应子工具；每个工具弹出独立窗口，互不影响。',
              foreground='#666').pack(anchor='w', pady=(1, 8))

    # ---------- 功能列表（紧凑横向卡片：左=名称+描述，右=按钮） ----------
    features_box = ttk.LabelFrame(outer, text='功能列表', padding=8)
    features_box.pack(fill='x', pady=(0, 6))

    for feat in FEATURES:
        row = ttk.Frame(features_box, padding=(4, 3))
        row.pack(fill='x')

        left = ttk.Frame(row)
        left.pack(side='left', fill='both', expand=True)
        ttk.Label(left, text=feat['name'],
                  font=('Microsoft YaHei UI', 11, 'bold')).pack(anchor='w')
        ttk.Label(left, text=feat['desc'],
                  foreground='#666', wraplength=560,
                  justify='left').pack(anchor='w', pady=(1, 0))

        ttk.Button(row, text='启动 →', width=10,
                   command=lambda s=feat['script']: launch_module(s, log)
                   ).pack(side='right', padx=(10, 0))

    # ---------- 启动日志（固定 5 行，自身可局部滚动，不影响主窗口） ----------
    ttk.Label(outer, text='启动日志').pack(anchor='w', pady=(6, 0))
    log_text = scrolledtext.ScrolledText(outer, height=5, state='disabled',
                                          wrap='word', font=('Consolas', 10))
    log_text.pack(fill='x', pady=(3, 0))

    def log(msg):
        log_text.config(state='normal')
        log_text.insert('end', msg)
        log_text.see('end')
        log_text.config(state='disabled')

    log('提示：每个功能都会在独立窗口中运行；关闭子窗口不影响本程序。\n')

    # ---------- 自适应窗口高度：刚好包住全部内容，无需滚动条 ----------
    root.update_idletasks()
    screen_h = root.winfo_screenheight()
    # 内容需要的总高度 + 窗口边框余量；上限取屏幕高度的 90%，避免超出屏幕
    target_h = min(outer.winfo_reqheight() + 30, int(screen_h * 0.9))
    root.geometry(f'{max(outer.winfo_reqwidth() + 30, 560)}x{target_h}')

    root.mainloop()


if __name__ == '__main__':
    main()