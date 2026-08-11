#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PPT流程图工具 — Word文档拆分 + PPT流程图 一键生成（图形界面）

本工具将「文档拆分」与「流程图生成」两大能力封装为本地可视化工具，
无需命令行，无需联网，双击运行即可。

功能：
  标签页① Word文档拆分：按大纲级别拆分 docx，保留图片，输出到指定目录。
  标签页② PPT流程图生成：从 CSV 节点表 + 配色预设生成流程图 PPT。

依赖：Python 3.x（标准库，无第三方依赖）

用法：
  py -3 小工具/PPT流程图工具.py
  （或双击 小工具/启动_PPT流程图工具.bat）
"""
import argparse
import csv
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ---------- 路径定位 ----------
TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(TOOL_DIR)
SPLIT_SCRIPT = os.path.join(ROOT_DIR, "tools", "split_docx_by_level.py")
DOCX2CSV_SCRIPT = os.path.join(ROOT_DIR, "tools", "docx_to_flow_csv.py")
CSV_SCRIPT = os.path.join(ROOT_DIR, "生成脚本", "csv_to_flowchart.py")
PRESETS_DIR = os.path.join(ROOT_DIR, "skills", "flowchart-skill", "presets")
BLUE = "#1F3864"
LIGHT_BG = "#F4F6FB"
ROW_PAD = 8


# ---------- 工具函数 ----------
def run_cmd(cmd, desc):
    """后台执行命令并回显输出，返回 (returncode, 输出文本)。"""
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    try:
        result = subprocess.run(cmd, capture_output=True, text=True,
                                encoding="utf-8", errors="replace",
                                env=env, timeout=600)
        out = (result.stdout or "") + ("\n" + result.stderr if result.stderr else "")
        return result.returncode, out.strip()
    except Exception as e:
        return 1, f"执行异常: {e}"


def list_presets():
    """列出可用配色预设。"""
    if not os.path.isdir(PRESETS_DIR):
        return []
    return sorted(f[:-5] for f in os.listdir(PRESETS_DIR) if f.endswith(".json"))


# ---------- 主窗口 ----------
class FlowchartTool(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PPT流程图工具")
        self.geometry("860x620")
        self.minsize(760, 540)
        self.configure(bg=LIGHT_BG)

        self._build_header()
        self._build_tabs()
        self._build_footer()

    def _build_header(self):
        bar = tk.Frame(self, bg=BLUE, height=64)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="PPT流程图工具 · Word拆分 / PPT流程图 离线工具",
                 bg=BLUE, fg="white", font=("Microsoft YaHei", 15, "bold")).pack(side="left", padx=18, pady=12)
        tk.Label(bar, text="离线可用 · 无需联网 · 无需命令行",
                 bg=BLUE, fg="#C9DAF8", font=("Microsoft YaHei", 9)).pack(side="right", padx=18)

    def _build_tabs(self):
        self.tab = ttk.Notebook(self)
        self.tab.pack(fill="both", expand=True, padx=14, pady=10)
        self._tab_split = ttk.Frame(self.tab)
        self._tab_flow = ttk.Frame(self.tab)
        self.tab.add(self._tab_split, text="  ① Word文档拆分  ")
        self.tab.add(self._tab_flow, text="  ② PPT流程图生成  ")
        self._build_split_tab()
        self._build_flow_tab()

    def _build_footer(self):
        foot = tk.Frame(self, bg=LIGHT_BG)
        foot.pack(fill="x", side="bottom")
        tk.Label(foot, text="© 2026 PPT流程图工具 · 双击启动_PPT流程图工具.bat 运行",
                 bg=LIGHT_BG, fg="#888", font=("Microsoft YaHei", 9)).pack(padx=18, pady=6, anchor="w")

    # ---------- 标签①：文档拆分 ----------
    def _build_split_tab(self):
        f = self._tab_split
        self._add_section_title(f, "一、选择源文档")

        row = tk.Frame(f); row.pack(fill="x", padx=16, pady=4)
        tk.Label(row, text="Word 文档:", font=("Microsoft YaHei", 10)).pack(side="left")
        self.split_docx_var = tk.StringVar()
        tk.Entry(row, textvariable=self.split_docx_var, width=62).pack(side="left", padx=8)
        tk.Button(row, text="浏览…", command=self._pick_split_docx,
                  bg="#E8EDF7", font=("Microsoft YaHei", 10)).pack(side="left")

        self._add_section_title(f, "二、拆分设置")
        row = tk.Frame(f); row.pack(fill="x", padx=16, pady=4)
        tk.Label(row, text="拆分级别:", font=("Microsoft YaHei", 10)).pack(side="left")
        self.split_level_var = tk.StringVar(value="1")
        cmb = ttk.Combobox(row, textvariable=self.split_level_var, values=["1", "2", "3"],
                           width=4, state="readonly")
        cmb.pack(side="left", padx=8)
        tk.Label(row, text="(1=一级标题章节, 2=二级标题章节)", fg="#666",
                 font=("Microsoft YaHei", 9)).pack(side="left")

        row = tk.Frame(f); row.pack(fill="x", padx=16, pady=4)
        tk.Label(row, text="输出目录:", font=("Microsoft YaHei", 10)).pack(side="left")
        self.split_out_var = tk.StringVar()
        tk.Entry(row, textvariable=self.split_out_var, width=62).pack(side="left", padx=8)
        tk.Button(row, text="浏览…", command=self._pick_split_out,
                  bg="#E8EDF7", font=("Microsoft YaHei", 10)).pack(side="left")

        self._add_section_title(f, "三、执行")
        row = tk.Frame(f); row.pack(fill="x", padx=16, pady=6)
        self.split_btn = tk.Button(row, text="开始拆分", command=self._run_split,
                                   bg=BLUE, fg="white", font=("Microsoft YaHei", 11, "bold"),
                                   padx=20, pady=6)
        self.split_btn.pack(side="left")
        tk.Button(row, text="打开输出目录", command=self._open_split_dir,
                  bg="#E8EDF7", font=("Microsoft YaHei", 10)).pack(side="left", padx=10)

        self._add_section_title(f, "执行日志")
        self.split_log = tk.Text(f, height=8, bg="#1E1E1E", fg="#D4D4D4",
                                 font=("Consolas", 9), state="disabled")
        self.split_log.pack(fill="both", expand=True, padx=16, pady=4)

    # ---------- 标签②：流程图生成 ----------
    def _build_flow_tab(self):
        f = self._tab_flow
        self._add_section_title(f, "〇、CSV 来源（推荐：AI 生成后填入下方节点表）")
        row = tk.Frame(f); row.pack(fill="x", padx=16, pady=4)
        tk.Label(row, text="Word 文档:", font=("Microsoft YaHei", 10)).pack(side="left")
        self.flow_docx_var = tk.StringVar()
        tk.Entry(row, textvariable=self.flow_docx_var, width=56).pack(side="left", padx=8)
        tk.Button(row, text="浏览…", command=self._pick_flow_docx,
                  bg="#E8EDF7", font=("Microsoft YaHei", 10)).pack(side="left")
        self.docx2csv_btn = tk.Button(row, text="规则预览", command=self._run_docx_to_csv,
                                      bg="#FFF2CC", font=("Microsoft YaHei", 10))
        self.docx2csv_btn.pack(side="left", padx=8)
        tk.Label(row, text="(离线规则预览仅作参考，正式CSV请由AI生成后填入)",
                 fg="#666", font=("Microsoft YaHei", 9)).pack(side="left")

        self._add_section_title(f, "一、选择数据文件")

        row = tk.Frame(f); row.pack(fill="x", padx=16, pady=4)
        tk.Label(row, text="CSV 节点表:", font=("Microsoft YaHei", 10)).pack(side="left")
        self.flow_csv_var = tk.StringVar()
        tk.Entry(row, textvariable=self.flow_csv_var, width=58).pack(side="left", padx=8)
        tk.Button(row, text="浏览…", command=self._pick_flow_csv,
                  bg="#E8EDF7", font=("Microsoft YaHei", 10)).pack(side="left")

        row = tk.Frame(f); row.pack(fill="x", padx=16, pady=4)
        tk.Label(row, text="配色方案:", font=("Microsoft YaHei", 10)).pack(side="left")
        self.flow_preset_var = tk.StringVar(value="green")
        presets = list_presets() or ["green", "blue", "red", "yellow"]
        cmb = ttk.Combobox(row, textvariable=self.flow_preset_var, values=presets,
                           width=10, state="readonly")
        cmb.pack(side="left", padx=8)
        tk.Label(row, text=" 标题:", font=("Microsoft YaHei", 10)).pack(side="left", padx=(16, 0))
        self.flow_title_var = tk.StringVar()
        tk.Entry(row, textvariable=self.flow_title_var, width=30).pack(side="left", padx=8)

        self._add_section_title(f, "二、连接线设置（默认不连线）")
        row = tk.Frame(f); row.pack(fill="x", padx=16, pady=4)
        self.flow_conn_var = tk.BooleanVar(value=False)
        tk.Checkbutton(row, text="生成连接线", variable=self.flow_conn_var,
                       font=("Microsoft YaHei", 10)).pack(side="left")

        self._add_section_title(f, "三、输出")
        row = tk.Frame(f); row.pack(fill="x", padx=16, pady=4)
        tk.Label(row, text="输出文件:", font=("Microsoft YaHei", 10)).pack(side="left")
        self.flow_out_var = tk.StringVar()
        tk.Entry(row, textvariable=self.flow_out_var, width=58).pack(side="left", padx=8)
        tk.Button(row, text="浏览…", command=self._pick_flow_out,
                  bg="#E8EDF7", font=("Microsoft YaHei", 10)).pack(side="left")

        self._add_section_title(f, "四、执行")
        row = tk.Frame(f); row.pack(fill="x", padx=16, pady=6)
        self.flow_btn = tk.Button(row, text="生成流程图", command=self._run_flow,
                                  bg=BLUE, fg="white", font=("Microsoft YaHei", 11, "bold"),
                                  padx=20, pady=6)
        self.flow_btn.pack(side="left")
        self.flow_template_btn = tk.Button(row, text="查看CSV模板", command=self._open_csv_template_dir,
                                           bg="#E8EDF7", font=("Microsoft YaHei", 10))
        self.flow_template_btn.pack(side="left", padx=10)
        tk.Button(row, text="打开输出目录", command=self._open_flow_dir,
                  bg="#E8EDF7", font=("Microsoft YaHei", 10)).pack(side="left", padx=10)

        self._add_section_title(f, "执行日志")
        self.flow_log = tk.Text(f, height=8, bg="#1E1E1E", fg="#D4D4D4",
                                font=("Consolas", 9), state="disabled")
        self.flow_log.pack(fill="both", expand=True, padx=16, pady=4)

    # ---------- 通用 UI 辅助 ----------
    def _add_section_title(self, parent, text):
        tk.Label(parent, text=text, bg="#E2E8F5", fg=BLUE,
                 font=("Microsoft YaHei", 10, "bold"),
                 anchor="w").pack(fill="x", padx=0, pady=(8, 2))

    def _log(self, widget, msg):
        widget.config(state="normal")
        widget.insert("end", msg + "\n")
        widget.see("end")
        widget.config(state="disabled")

    # ---------- 文件选择 ----------
    def _pick_split_docx(self):
        p = filedialog.askopenfilename(title="选择Word文档",
                                       filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if p:
            self.split_docx_var.set(p)

    def _pick_split_out(self):
        p = filedialog.askdirectory(title="选择输出目录")
        if p:
            self.split_out_var.set(p)

    def _pick_flow_docx(self):
        p = filedialog.askopenfilename(title="选择Word文档生成CSV",
                                       filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if p:
            self.flow_docx_var.set(p)
            if not self.flow_title_var.get():
                self.flow_title_var.set(os.path.splitext(os.path.basename(p))[0])

    def _run_docx_to_csv(self):
        docx = self.flow_docx_var.get().strip()
        if not docx or not os.path.isfile(docx):
            messagebox.showwarning("提示", "请先选择有效的 Word 文档")
            return
        base = os.path.splitext(os.path.basename(docx))[0]
        out = os.path.join(os.path.dirname(docx), base + "_规则预览.csv")
        self.docx2csv_btn.config(state="disabled")
        self._log(self.flow_log, "▶ 开始规则预览（离线参考，正式CSV建议由AI生成）…")
        self._log(self.flow_log, f"  Word: {docx}")
        self._log(self.flow_log, f"  输出: {out}")
        threading.Thread(target=self._do_docx_to_csv, args=(docx, out), daemon=True).start()

    def _do_docx_to_csv(self, docx, out):
        code, text = run_cmd([sys.executable, DOCX2CSV_SCRIPT, docx, "--out", out],
                             "规则预览")
        self.after(0, self._docx_to_csv_done, code, text, out)

    def _docx_to_csv_done(self, code, text, out):
        self.docx2csv_btn.config(state="normal")
        self._log(self.flow_log, text)
        if code == 0 and os.path.isfile(out):
            # 自动加载到 CSV 节点表输入框
            self.flow_csv_var.set(out)
            # 自动设置标题
            base = os.path.splitext(os.path.basename(out))[0].replace("_规则预览", "")
            if not self.flow_title_var.get():
                self.flow_title_var.set(base)
            self._log(self.flow_log, "✓ 规则预览已输出（仅供参考）。已自动加载到 CSV 节点表。")
            self._log(self.flow_log, "  提示：正式使用建议由 AI 生成 CSV（参考 flowchart-skill Step 3）")
            messagebox.showinfo("提示", "规则预览已输出\n已自动加载到 CSV 节点表\n\n⚠️ 仅供参考，正式CSV建议由AI生成")
        else:
            self._log(self.flow_log, "✗ 预览生成失败，请查看日志")
            messagebox.showerror("失败", "规则预览生成失败，请查看日志")

    def _pick_flow_csv(self):
        p = filedialog.askopenfilename(title="选择CSV节点表",
                                       filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")])
        if p:
            self.flow_csv_var.set(p)
            if not self.flow_title_var.get():
                base = os.path.splitext(os.path.basename(p))[0]
                self.flow_title_var.set(base)

    def _pick_flow_out(self):
        p = filedialog.asksaveasfilename(title="保存流程图",
                                         defaultextension=".pptx",
                                         filetypes=[("PowerPoint", "*.pptx")])
        if p:
            self.flow_out_var.set(p)

    # ---------- 打开目录 ----------
    def _open_split_dir(self):
        d = self.split_out_var.get().strip()
        if d and os.path.isdir(d):
            os.startfile(d)
        else:
            self._run_open(ROOT_DIR if not d else d)

    def _open_flow_dir(self):
        out = self.flow_out_var.get().strip()
        if out:
            self._run_open(os.path.dirname(out) or ROOT_DIR)
        else:
            self._run_open(ROOT_DIR)

    def _run_open(self, path):
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showwarning("提示", f"无法打开目录: {e}")

    def _open_csv_template_dir(self):
        self._run_open(os.path.join(ROOT_DIR, "skills", "flowchart-skill", "templates"))

    # ---------- 执行：文档拆分 ----------
    def _run_split(self):
        docx = self.split_docx_var.get().strip()
        if not docx or not os.path.isfile(docx):
            messagebox.showwarning("提示", "请先选择有效的 Word 文档")
            return
        out = self.split_out_var.get().strip()
        if not out:
            out = os.path.join(os.path.dirname(docx), os.path.splitext(os.path.basename(docx))[0] + "_拆分")
            self.split_out_var.set(out)
        try:
            level = int(self.split_level_var.get())
        except ValueError:
            level = 1

        self.split_btn.config(state="disabled")
        threading.Thread(target=self._do_split, args=(docx, out, level), daemon=True).start()

    def _do_split(self, docx, out, level):
        self._log(self.split_log, "▶ 开始拆分文档…")
        self._log(self.split_log, f"  源文档: {docx}")
        self._log(self.split_log, f"  级别: {level}  →  输出: {out}")
        code, text = run_cmd([sys.executable, SPLIT_SCRIPT, docx, out, str(level)],
                             "拆分")
        self._log(self.split_log, text)
        self.after(0, self._split_done, code)

    def _split_done(self, code):
        self.split_btn.config(state="normal")
        if code == 0:
            self._log(self.split_log, "✓ 拆分完成！")
            messagebox.showinfo("成功", "Word 文档拆分完成")
        else:
            self._log(self.split_log, "✗ 拆分失败，请查看日志")
            messagebox.showerror("失败", "拆分失败，请查看日志")

    # ---------- 执行：流程图生成 ----------
    def _run_flow(self):
        csv_path = self.flow_csv_var.get().strip()
        if not csv_path or not os.path.isfile(csv_path):
            messagebox.showwarning("提示", "请先选择有效的 CSV 节点表")
            return
        out = self.flow_out_var.get().strip()
        if not out:
            base = os.path.splitext(os.path.basename(csv_path))[0]
            out = os.path.join(os.path.dirname(csv_path), base + "_流程图.pptx")
            self.flow_out_var.set(out)

        cmd = [sys.executable, CSV_SCRIPT, csv_path, "--out", out]
        preset = self.flow_preset_var.get().strip()
        if preset:
            cmd += ["--preset", preset]
        title = self.flow_title_var.get().strip()
        if title:
            cmd += ["--title", title]
        if self.flow_conn_var.get():
            cmd.append("--connectors")

        self.flow_btn.config(state="disabled")
        self._log(self.flow_log, "▶ 开始生成流程图…")
        self._log(self.flow_log, f"  CSV: {csv_path}")
        self._log(self.flow_log, f"  配色: {preset}  标题: {title}" +
                  ("   [连线]" if self.flow_conn_var.get() else "   [默认无连线]"))
        threading.Thread(target=self._do_flow, args=(cmd, csv_path), daemon=True).start()

    def _do_flow(self, cmd, csv_path):
        # 先打印 CSV 预览（节点数量）
        try:
            with open(csv_path, encoding="utf-8-sig") as f:
                n = sum(1 for r in csv.DictReader(f) if r.get("seq") and r.get("content"))
            self._log(self.flow_log, f"  CSV 节点数: {n}")
        except Exception:
            pass
        code, text = run_cmd(cmd, "生成流程图")
        self._log(self.flow_log, text)
        self.after(0, self._flow_done, code)

    def _flow_done(self, code):
        self.flow_btn.config(state="normal")
        if code == 0:
            self._log(self.flow_log, "✓ 流程图生成完成！")
            messagebox.showinfo("成功", "PPT 流程图生成完成")
        else:
            self._log(self.flow_log, "✗ 生成失败，请查看日志")
            messagebox.showerror("失败", "流程图生成失败，请查看日志")


def main():
    app = FlowchartTool()
    app.mainloop()


if __name__ == "__main__":
    main()
