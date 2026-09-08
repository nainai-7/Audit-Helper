# -*- coding: utf-8 -*-
"""
账龄首次划分（FIFO 瀑布）GUI 工具
================================
- 一键获取空白模板：生成「首次划分账龄-资产类」标准 A~X 列 Excel
- 读取已填写的模板：用清晰逻辑复算账龄桶 R~W + 勾稽校验 X
- 导出结果：回写计算值并附异常备注

依赖：tkinter（Python 内置）、openpyxl（需 pip install openpyxl）
运行：python aging_gui.py
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# tkinter 延迟/守护导入：无 GUI 环境下（如仅作逻辑库导入自测）也不报错
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:  # pragma: no cover
    tk = ttk = filedialog = messagebox = None


# ==========================================================================
# 1) 复算核心逻辑（与「首次划分账龄-资产类.xlsx」巨公式等价，拆分为可读函数）
# ==========================================================================
def net_debit(dr, cr):
    """基期余额 >=0 时，该期「借/贷」对更老桶的净冲减额。"""
    if dr < 0 and cr < 0:
        return -dr
    if dr < 0:
        return cr - dr
    return cr


def net_credit(dr, cr):
    """基期余额 <0（红字）时，该期「借/贷」对更老桶的净冲减额（负数归零）。"""
    if dr < 0 and cr < 0:
        inner = -cr
    elif cr < 0:
        inner = dr - cr
    else:
        inner = dr
    return inner if inner >= 0 else 0


def _bucket(base, subs, older, red):
    """计算单个账龄桶（FIFO 瀑布，含红字分支与地板截断）。"""
    if red:
        val = base + sum(net_credit(d, c) for (d, c) in subs) - sum(older)
        return val if val < 0 else 0.0
    reduction = sum(max(net_debit(d, c), 0.0) for (d, c) in subs) + sum(older)
    val = base - reduction
    return val if val > 0 else 0.0


@dataclass
class AgingResult:
    name: str
    B: float = 0.0
    C: float = 0.0
    D: float = 0.0
    E: float = 0.0
    F: float = 0.0
    G: float = 0.0
    H: float = 0.0
    I: float = 0.0
    J: float = 0.0
    K: float = 0.0
    L: float = 0.0
    M: float = 0.0
    N: float = 0.0
    O: float = 0.0
    P: float = 0.0
    Q: float = 0.0
    R: float = 0.0
    S: float = 0.0
    T: float = 0.0
    U: float = 0.0
    V: float = 0.0
    W: float = 0.0
    warnings: list = field(default_factory=list)

    @property
    def bucket_sum(self):
        return self.R + self.S + self.T + self.U + self.V + self.W

    @property
    def X(self):
        """勾稽校验列 = R+S+T+U+V+W − Q（应≈0）。"""
        return self.bucket_sum - self.Q


def roll_forward(B, C, D, F, G, I, J, K, L, M, O, P, name="", strict=True):
    """单客户账龄瀑布计算，返回 AgingResult。"""
    B = B or 0.0
    C = C or 0.0
    D = D or 0.0
    F = F or 0.0
    G = G or 0.0
    I = I or 0.0
    J = J or 0.0
    K = K or 0.0
    L = L or 0.0
    M = M or 0.0
    O = O or 0.0
    P = P or 0.0

    E = B + C - D
    H = E + F - G
    N = K + L - M
    Q = N + O - P

    W = _bucket(B, [(C, D), (F, G), (I, J), (L, M), (O, P)], [], red=(B < 0))
    V = _bucket(E, [(F, G), (I, J), (L, M), (O, P)], [W], red=(E < 0))
    U = _bucket(H, [(I, J), (L, M), (O, P)], [W, V], red=(H < 0))
    T = _bucket(K, [(L, M), (O, P)], [W, V, U], red=(K < 0))
    S = _bucket(N, [(O, P)], [W, V, U, T], red=(N < 0))
    R = Q - W - V - U - T - S

    res = AgingResult(
        name=name, B=B, C=C, D=D, E=E, F=F, G=G, H=H, I=I, J=J,
        K=K, L=L, M=M, N=N, O=O, P=P, Q=Q,
        R=R, S=S, T=T, U=U, V=V, W=W,
    )
    if strict:
        _check(res)
    return res


def _check(res: AgingResult, tol=1e-6):
    """勾稽校验 + 红字/负数异常提示。"""
    diff = res.Q - res.bucket_sum
    if abs(diff) > tol:
        res.warnings.append(
            f"勾稽不符: R+S+T+U+V+W={res.bucket_sum:.2f} 与 Q={res.Q:.2f} 差 {diff:.2f}"
        )
    for col, val in (("B", res.B), ("E", res.E), ("H", res.H),
                     ("K", res.K), ("N", res.N), ("Q", res.Q)):
        if val < -tol:
            res.warnings.append(f"{col} 为期初/期末红字(贷方余额) {val:.2f}，账龄方向需复核")
    if res.R < -tol:
        res.warnings.append(f"R(一年以内)为负 {res.R:.2f}：可能上期账龄漏填或期初为红字")
    for col in ("R", "S", "T", "U", "V", "W"):
        v = getattr(res, col)
        if v < -tol:
            res.warnings.append(f"{col} 桶为负 {v:.2f}")


def _cell_float(ws, row, col):
    v = ws[f"{col}{row}"].value
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# ==========================================================================
# 2) 模板定义与生成
# ==========================================================================
TEMPLATE_HEADERS = [
    ("A", "客户/科目"), ("B", "五年前期末"), ("C", "借"), ("D", "贷"),
    ("E", "四年前期末"), ("F", "借"), ("G", "贷"), ("H", "三年前期末"),
    ("I", "借"), ("J", "贷"), ("K", "二年前期末"), ("L", "借"), ("M", "贷"),
    ("N", "一年前期末"), ("O", "借"), ("P", "贷"), ("Q", "期末余额"),
    ("R", "一年以内"), ("S", "1-2年"), ("T", "2-3年"), ("U", "3-4年"),
    ("V", "4-5年"), ("W", "五年以上"), ("X", "勾稽校验"),
]
# 模板中自动计算的滚调链（用户只需填 A + B,C,D,F,G,I,J,K,L,M,O,P）
# 注：R~W 账龄桶与 X 勾稽由本工具复算后写入，模板内不预置公式以免显示误导值
TEMPLATE_FORMULAS = {
    "E": "=B{row}+C{row}-D{row}",
    "H": "=E{row}+F{row}-G{row}",
    "N": "=K{row}+L{row}-M{row}",
    "Q": "=N{row}+O{row}-P{row}",
}


def make_template(path, nrows=500):
    """生成空白模板：标准 A~X 列头 + 滚调链公式（E/H/N/Q/X）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "账龄首次划分"
    header_font = Font(bold=True)
    for col, label in TEMPLATE_HEADERS:
        c = ws[f"{col}1"]
        c.value = label
        c.font = header_font
        c.alignment = Alignment(horizontal="center")
    for r in range(2, nrows + 2):
        for col, formula in TEMPLATE_FORMULAS.items():
            ws[f"{col}{r}"] = formula.format(row=r)
    widths = {"A": 22}
    for col, _ in TEMPLATE_HEADERS:
        if col != "A":
            widths[col] = 11
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    wb.save(path)
    return path


# ==========================================================================
# 3) 读取已填写的模板并复算
# ==========================================================================
RAW_INPUT_COLS = ["B", "C", "D", "F", "G", "I", "J", "K", "L", "M", "O", "P"]


def load_and_compute(path, strict=True):
    """读取模板（data_only 取缓存值），逐行复算，返回 (results, mismatches)。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    results = []
    mismatches = []
    for r in range(2, ws.max_row + 1):
        name = ws[f"A{r}"].value
        raw = {c: _cell_float(ws, r, c) for c in RAW_INPUT_COLS}
        if name is None and all(v == 0.0 for v in raw.values()):
            continue  # 空行
        res = roll_forward(
            B=raw["B"], C=raw["C"], D=raw["D"],
            F=raw["F"], G=raw["G"], I=raw["I"], J=raw["J"],
            K=raw["K"], L=raw["L"], M=raw["M"],
            O=raw["O"], P=raw["P"],
            name=str(name) if name is not None else f"第{r}行",
            strict=strict,
        )
        # 与原表缓存的 R~W 比对（仅当原表确有值）
        tol = 1e-4
        for c in "RSTUVW":
            orig = _cell_float(ws, r, c)
            if abs(orig) > tol and abs(getattr(res, c) - orig) > tol:
                msg = f"与原表{c}列差异: 计算={getattr(res, c):.2f} 原表={orig:.2f}"
                res.warnings.append(msg)
                mismatches.append((r, c, getattr(res, c), orig))
        results.append(res)
    return results, mismatches


def export_results(results, path):
    """将复算结果回写为标准 A~X 列 + Y 列异常备注。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "账龄首次划分-结果"
    headers = [("A", "客户/科目")] + TEMPLATE_HEADERS[1:] + [("Y", "异常备注")]
    for i, (col, label) in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=label)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
    order = list("BCDEFGHIJKLMNOPQRSTUVWX")  # 不含 A（客户名单列）
    for ri, res in enumerate(results, start=2):
        ws.cell(row=ri, column=1, value=res.name)
        for ci, col in enumerate(order, start=2):
            cell = ws.cell(row=ri, column=ci, value=round(getattr(res, col), 2))
            cell.alignment = Alignment(horizontal="right")
        # 红字/勾稽异常标红
        if res.warnings:
            warn_cell = ws.cell(row=ri, column=len(order) + 2, value="；".join(res.warnings))
            warn_cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            warn_cell.font = Font(color="9C0006")
    for i in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 22 if i in (1, len(headers)) else 11
    wb.save(path)
    return path


# ==========================================================================
# 4) GUI
# ==========================================================================
DISPLAY_COLS = (
    [("客户", "name", 120)]
    + [(h, k, 72) for k, h in TEMPLATE_HEADERS[1:]]  # B..X
    + [("异常", "_warn", 60)]
)


class AgingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("账龄首次划分（FIFO 瀑布）GUI")
        self.root.geometry("1280x720")
        self.results = []

        # ---- 工具栏 ----
        frm = ttk.Frame(root)
        frm.pack(fill=tk.X, padx=6, pady=4)
        ttk.Button(frm, text="① 获取模板", command=self.on_get_template).pack(side=tk.LEFT, padx=3)
        ttk.Button(frm, text="② 选择文件并复算", command=self.on_load).pack(side=tk.LEFT, padx=3)
        ttk.Button(frm, text="③ 导出结果", command=self.on_export).pack(side=tk.LEFT, padx=3)
        ttk.Button(frm, text="清空", command=self.on_clear).pack(side=tk.LEFT, padx=3)
        self.status = ttk.Label(frm, text="未加载数据", foreground="#555")
        self.status.pack(side=tk.RIGHT, padx=6)

        # ---- 表格 ----
        table_frame = ttk.Frame(root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)
        self.tree = ttk.Treeview(table_frame, show="headings", height=20)
        self.tree["columns"] = [k for _, k, _ in DISPLAY_COLS]
        for header, key, w in DISPLAY_COLS:
            anchor = tk.W if key in ("name", "_warn") else tk.E
            self.tree.heading(key, text=header)
            self.tree.column(key, width=w, anchor=anchor)
        self.tree.tag_configure("warn", background="#FFE3E3")
        ys = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        xs = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ys.pack(side=tk.RIGHT, fill=tk.Y)
        xs.pack(side=tk.BOTTOM, fill=tk.X)

        # ---- 异常明细 ----
        warn_frame = ttk.LabelFrame(root, text="异常明细")
        warn_frame.pack(fill=tk.BOTH, padx=6, pady=4, expand=False)
        self.warn_text = tk.Text(warn_frame, height=6, wrap=tk.WORD)
        self.warn_text.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

    # ---- 事件 ----
    def on_get_template(self):
        default = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "首次划分账龄-资产类-模板.xlsx")
        path = filedialog.asksaveasfilename(
            title="保存模板", defaultextension=".xlsx",
            initialfile=os.path.basename(default),
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            make_template(path)
            messagebox.showinfo("完成", f"模板已生成：\n{path}\n\n在 Excel 中填写 A 列客户名与 "
                                       f"B,C,D,F,G,I,J,K,L,M,O,P 各期借贷/期末即可。")
        except Exception as e:  # pragma: no cover
            messagebox.showerror("生成失败", str(e))

    def on_load(self):
        path = filedialog.askopenfilename(
            title="选择已填写的账龄模板",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            results, mismatches = load_and_compute(path, strict=True)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        self.results = results
        self._render()
        n_warn = sum(1 for r in results if r.warnings)
        self.status.config(
            text=f"共 {len(results)} 行 | 勾稽/红字异常 {n_warn} 行 | 与原表差异 {len(mismatches)} 处"
        )
        if not results:
            messagebox.showinfo("提示", "未读取到有效数据行（请确认模板格式与列头一致）。")

    def on_export(self):
        if not self.results:
            messagebox.showwarning("提示", "请先「选择文件并复算」。")
            return
        default = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "首次划分账龄-资产类-结果.xlsx")
        path = filedialog.asksaveasfilename(
            title="导出结果", defaultextension=".xlsx",
            initialfile=os.path.basename(default),
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            export_results(self.results, path)
            messagebox.showinfo("完成", f"结果已导出：\n{path}")
        except Exception as e:  # pragma: no cover
            messagebox.showerror("导出失败", str(e))

    def on_clear(self):
        self.results = []
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.warn_text.delete("1.0", tk.END)
        self.status.config(text="未加载数据")

    def _render(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.warn_text.delete("1.0", tk.END)
        for res in self.results:
            vals = []
            for _, key, _ in DISPLAY_COLS:
                if key == "name":
                    vals.append(res.name)
                elif key == "_warn":
                    vals.append(f"⚠ {len(res.warnings)}" if res.warnings else "")
                else:
                    vals.append(f"{getattr(res, key):.2f}")
            tag = ("warn",) if res.warnings else ()
            self.tree.insert("", tk.END, values=vals, tags=tag)
            for w in res.warnings:
                self.warn_text.insert(tk.END, f"[{res.name}] {w}\n")


def main():
    if tk is None:
        raise SystemExit("当前环境缺少 tkinter，无法启动 GUI（请使用带 Tk 的 Python 运行）。")
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.0)
    except Exception:
        pass
    AgingApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
