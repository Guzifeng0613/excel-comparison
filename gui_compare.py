#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel比较工具 - 拖拽界面版 v2.1
"""

import os
import sys
import subprocess
import tempfile
from typing import Optional
import tkinter as tk
from tkinter import ttk, messagebox

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False
    from tkinter import filedialog


class ExcelCompareGUI:
    def __init__(self):
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        self.root.title("Excel 比较工具 v2.1")
        self.root.geometry("800x650")
        self.root.minsize(700, 550)

        # 居中窗口
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        win_w, win_h = 800, 650
        x = (screen_w - win_w) // 2
        y = (screen_h - win_h) // 2
        self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")

        self.file1_path: Optional[str] = None
        self.file2_path: Optional[str] = None

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)

    def _build_ui(self):
        """构建界面"""
        # 标题栏
        header = tk.Frame(self.root, bg="#2C5AA0", height=60)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)

        tk.Label(
            header, text="Excel 文件比较工具",
            font=("Microsoft YaHei", 20, "bold"),
            fg="white", bg="#2C5AA0"
        ).pack(pady=15)

        # 主容器
        main = tk.Frame(self.root, bg="#F5F5F5")
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # ===== 文件拖放区 =====
        drop_frame = tk.LabelFrame(
            main, text="拖放 Excel 文件",
            font=("Microsoft YaHei", 11, "bold"),
            bg="white", padx=10, pady=10
        )
        drop_frame.pack(fill=tk.BOTH, expand=True)

        file_row = tk.Frame(drop_frame, bg="white")
        file_row.pack(fill=tk.BOTH, expand=True)

        # 文件1
        f1 = tk.Frame(file_row, bg="#FAFAFA", relief=tk.RAISED, bd=2)
        f1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))

        tk.Label(f1, text="📁 文件 1", font=("Microsoft YaHei", 11, "bold"),
                 bg="#FAFAFA").pack(pady=(10, 5))

        self.file1_label = tk.Label(
            f1, text="拖放文件到此处\n或 双击选择",
            font=("Microsoft YaHei", 10), width=25, height=8,
            bg="white", relief=tk.SUNKEN, bd=1, cursor="hand2"
        )
        self.file1_label.pack(pady=5, padx=10)

        self.file1_path_lbl = tk.Label(f1, text="未选择", font=("Microsoft YaHei", 8),
                                        fg="#999", bg="#FAFAFA")
        self.file1_path_lbl.pack(pady=(0, 8))

        # 文件2
        f2 = tk.Frame(file_row, bg="#FAFAFA", relief=tk.RAISED, bd=2)
        f2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))

        tk.Label(f2, text="📁 文件 2", font=("Microsoft YaHei", 11, "bold"),
                 bg="#FAFAFA").pack(pady=(10, 5))

        self.file2_label = tk.Label(
            f2, text="拖放文件到此处\n或 双击选择",
            font=("Microsoft YaHei", 10), width=25, height=8,
            bg="white", relief=tk.SUNKEN, bd=1, cursor="hand2"
        )
        self.file2_label.pack(pady=5, padx=10)

        self.file2_path_lbl = tk.Label(f2, text="未选择", font=("Microsoft YaHei", 8),
                                        fg="#999", bg="#FAFAFA")
        self.file2_path_lbl.pack(pady=(0, 8))

        # 绑定拖放事件
        if HAS_DND:
            self.file1_label.drop_target_register(DND_FILES)
            self.file2_label.drop_target_register(DND_FILES)
            self.file1_label.dnd_bind('<<Drop>>', lambda e: self._drop(e, 1))
            self.file2_label.dnd_bind('<<Drop>>', lambda e: self._drop(e, 2))
        else:
            self.file1_label.bind('<Double-Button-1>', lambda e: self._pick(1))
            self.file2_label.bind('<Double-Button-1>', lambda e: self._pick(2))

        # ===== 选项区 =====
        opt = tk.LabelFrame(main, text="选项",
                            font=("Microsoft YaHei", 10, "bold"),
                            bg="white", padx=15, pady=10)
        opt.pack(fill=tk.X, pady=(15, 0))

        # 第一行
        r1 = tk.Frame(opt, bg="white")
        r1.pack(fill=tk.X)

        tk.Label(r1, text="主键列:", font=("Microsoft YaHei", 10), bg="white").pack(side=tk.LEFT)

        self.key_var = tk.StringVar()
        self.key_entry = tk.Entry(r1, textvariable=self.key_var, width=20, font=("Microsoft YaHei", 10))
        self.key_entry.pack(side=tk.LEFT, padx=(5, 25))
        self.key_entry.insert(0, "（留空用第一列）")
        self.key_entry.config(fg="#999")
        self.key_entry.bind('<FocusIn>', lambda e: self._placeholder(e, True))
        self.key_entry.bind('<FocusOut>', lambda e: self._placeholder(e, False))

        tk.Label(r1, text="比较类型:", font=("Microsoft YaHei", 10), bg="white").pack(side=tk.LEFT, padx=(10, 0))

        self.type_var = tk.StringVar(value="精确匹配")
        types = ["精确匹配", "忽略空白", "包含匹配", "正则匹配"]
        self.type_combo = ttk.Combobox(r1, textvariable=self.type_var,
                                        values=types, width=12, state="readonly")
        self.type_combo.pack(side=tk.LEFT, padx=5)

        # 第二行
        r2 = tk.Frame(opt, bg="white")
        r2.pack(fill=tk.X, pady=(8, 0))

        self.case_var = tk.BooleanVar()
        tk.Checkbutton(r2, text="忽略大小写", variable=self.case_var,
                       font=("Microsoft YaHei", 9), bg="white").pack(side=tk.LEFT, padx=(0, 15))

        self.trim_var = tk.BooleanVar(value=True)
        tk.Checkbutton(r2, text="去首尾空格", variable=self.trim_var,
                       font=("Microsoft YaHei", 9), bg="white").pack(side=tk.LEFT, padx=15)

        self.merge_var = tk.BooleanVar(value=True)
        tk.Checkbutton(r2, text="合并空格", variable=self.merge_var,
                       font=("Microsoft YaHei", 9), bg="white").pack(side=tk.LEFT, padx=15)

        self.nl_var = tk.BooleanVar(value=True)
        tk.Checkbutton(r2, text="移除换行", variable=self.nl_var,
                       font=("Microsoft YaHei", 9), bg="white").pack(side=tk.LEFT, padx=15)

        # 第三行：标准文件选项
        r3 = tk.Frame(opt, bg="white")
        r3.pack(fill=tk.X, pady=(8, 0))

        self.file1_standard_var = tk.BooleanVar(value=True)
        tk.Checkbutton(r3, text="第一张表为标准（只高亮第二张的差异）",
                       variable=self.file1_standard_var,
                       font=("Microsoft YaHei", 9, "bold"), bg="white",
                       activebackground="white").pack(side=tk.LEFT)

        # ===== 按钮区 =====
        btn_row = tk.Frame(main, bg="#F5F5F5")
        btn_row.pack(fill=tk.X, pady=(20, 0))

        self.compare_btn = tk.Button(
            btn_row, text="▶ 开始比较",
            font=("Microsoft YaHei", 13, "bold"),
            bg="#2C5AA0", fg="white",
            padx=40, pady=10,
            state=tk.DISABLED,
            command=self._do_compare
        )
        self.compare_btn.pack(side=tk.LEFT)

        tk.Button(btn_row, text="清空",
                  font=("Microsoft YaHei", 10),
                  padx=20, pady=10,
                  command=self._clear
                  ).pack(side=tk.LEFT, padx=10)

        tk.Button(btn_row, text="帮助",
                  font=("Microsoft YaHei", 10),
                  padx=20, pady=10,
                  command=self._help
                  ).pack(side=tk.LEFT)

        # 状态栏
        self.status = tk.Label(self.root, text="就绪",
                               font=("Microsoft YaHei", 9),
                               fg="#666", anchor=tk.W,
                               bg="#EEE", pady=6)
        self.status.pack(fill=tk.X, side=tk.BOTTOM)

    def _placeholder(self, e, entering):
        if entering and self.key_entry.get() == "（留空用第一列）":
            self.key_entry.delete(0, tk.END)
            self.key_entry.config(fg="black")
        elif not entering and self.key_entry.get() == "":
            self.key_entry.insert(0, "（留空用第一列）")
            self.key_entry.config(fg="#999")

    def _drop(self, event, num):
        files = self.root.tk.splitlist(event.data)
        if files:
            self._set_file(num, files[0].strip('{}'))

    def _pick(self, num):
        path = filedialog.askopenfilename(
            title=f"选择文件 {num}",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("所有", "*.*")]
        )
        if path:
            self._set_file(num, path)

    def _set_file(self, num, path):
        if not os.path.exists(path):
            messagebox.showerror("错误", f"文件不存在:\n{path}")
            return

        ext = os.path.splitext(path)[1].lower()
        if ext not in ['.xlsx', '.xls', '.xlsm']:
            messagebox.showerror("错误", "请选择 Excel 文件")
            return

        name = os.path.basename(path)
        if num == 1:
            self.file1_path = path
            self.file1_label.config(text=f"✓ {name}", bg="#E8F5E9", fg="#2E7D32")
            self.file1_path_lbl.config(text=path, fg="#666")
        else:
            self.file2_path = path
            self.file2_label.config(text=f"✓ {name}", bg="#E8F5E9", fg="#2E7D32")
            self.file2_path_lbl.config(text=path, fg="#666")

        self._update_btn()
        self.status.config(text=f"已加载: {name}")

    def _update_btn(self):
        state = tk.NORMAL if (self.file1_path and self.file2_path) else tk.DISABLED
        self.compare_btn.config(state=state, bg="#2C5AA0" if state == tk.NORMAL else "#AAA")

    def _clear(self):
        self.file1_path = self.file2_path = None
        for lbl, path_lbl in [(self.file1_label, self.file1_path_lbl),
                               (self.file2_label, self.file2_path_lbl)]:
            lbl.config(text="拖放文件到此处\n或 双击选择", bg="white", fg="black")
            path_lbl.config(text="未选择", fg="#999")
        self._update_btn()
        self.status.config(text="就绪")

    def _help(self):
        messagebox.showinfo("帮助", """Excel 比较工具 v2.1

【比较类型】
• 精确匹配：值相等（123 = 123.0）
• 忽略空白：忽略首尾空格
• 包含匹配：检查是否包含
• 正则匹配：使用正则表达式

【主键列】
唯一标识行的列，逗号分隔多列
留空则使用第一列

【输出】
自动生成Excel报告，包含：
汇总统计 / 差异明细 / 列统计
仅存在文件1 / 仅存在文件2
""")

    def _do_compare(self):
        if not self.file1_path or not self.file2_path:
            messagebox.showwarning("提示", "请先选择两个文件")
            return

        self.status.config(text="比较中...")
        self.compare_btn.config(text="比较中...", state=tk.DISABLED)

        key_text = self.key_var.get().strip()
        key_cols = None
        if key_text and key_text != "（留空用第一列）":
            key_cols = [k.strip() for k in key_text.split(',')]

        try:
            from excel_compare import compare_excel_files
        except ImportError:
            messagebox.showerror("错误", "找不到比较模块")
            self._update_btn()
            return

        output = os.path.join(
            os.path.dirname(self.file1_path) or tempfile.gettempdir(),
            f"比较报告_{os.path.basename(self.file1_path)}.xlsx"
        ).replace(' ', '_')

        try:
            r = compare_excel_files(
                self.file1_path, self.file2_path,
                key_columns=key_cols,
                output_path=output,
                compare_type=self.type_var.get(),
                normalize_case=self.case_var.get(),
                trim_whitespace=self.trim_var.get(),
                merge_spaces=self.merge_var.get(),
                remove_newlines=self.nl_var.get(),
                file1_standard=self.file1_standard_var.get()
            )

            if r["success"]:
                s = r["statistics"]
                diffs = {k: v for k, v in s.get("diff_by_column", {}).items() if v > 0}

                # 获取高亮备份文件路径
                backup_files = r.get("backup_files", {})
                backup1 = backup_files.get("file1")
                backup2 = backup_files.get("file2")

                msg = f"""比较完成！

📁 文件1: {os.path.basename(self.file1_path)}
📁 文件2: {os.path.basename(self.file2_path)}

━━━━━━━━━━━━━━━
📊 统计结果
━━━━━━━━━━━━━━━
• 总主键数: {s['total_unique_keys']}
• 匹配行数: {s['matched_rows']}
• ✅ 匹配率: {s['match_percentage']}%
• ❌ 差异行: {s['different_rows']}
• ➖ 仅文件1: {s['file1_only_rows']}
• ➕ 仅文件2: {s['file2_only_rows']}"""

                if diffs:
                    msg += "\n\n📋 差异列分布:"
                    for col, cnt in sorted(diffs.items(), key=lambda x: -x[1]):
                        pct = cnt / s['different_rows'] * 100
                        msg += f"\n   • {col}: {cnt}处 ({pct:.0f}%)"

                msg += f"\n\n📄 报告:\n{output}"

                if backup1 or backup2:
                    msg += "\n\n🖍️ 高亮备份已生成（差异处已标红）"
                    if backup1:
                        msg += f"\n   • 文件1: {os.path.basename(backup1)}"
                    if backup2:
                        msg += f"\n   • 文件2: {os.path.basename(backup2)}"

                self.status.config(text="比较完成")
                messagebox.showinfo("完成", msg)

                # 打开高亮备份文件
                if backup1:
                    self._open(backup1)
                if backup2:
                    self._open(backup2)
            else:
                messagebox.showerror("失败", r.get("error", "未知错误"))
                self.status.config(text="比较失败")

        except Exception as e:
            messagebox.showerror("错误", str(e))
            import traceback
            traceback.print_exc()
        finally:
            self.compare_btn.config(text="▶ 开始比较")
            self._update_btn()

    def _open(self, path):
        try:
            if sys.platform == 'win32':
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.run(['open', path])
            else:
                subprocess.run(['xdg-open', path])
        except:
            pass

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    ExcelCompareGUI().run()
