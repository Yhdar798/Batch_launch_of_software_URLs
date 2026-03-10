# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import json
import os
import webbrowser

CONFIG_FILE = "launcher_config.json"


class LauncherApp:

    def __init__(self, root):

        self.root = root
        self.root.title("软件启动器  作者：tomato BUG反馈：hydra_7633@163.com")

        self.center_window(800, 480)

        self.data = {"apps": [], "urls": []}

        self.tooltip = None
        self.edit_entry = None

        self.load_config()
        self.create_ui()

    # ==========================
    # 窗口居中
    # ==========================

    def center_window(self, width, height):

        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()

        x = int((sw - width) / 2)
        y = int((sh - height) / 2)

        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def center_child(self, win, width=350, height=200):

        win.update_idletasks()

        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()

        x = int((sw - width) / 2)
        y = int((sh - height) / 2)

        win.geometry(f"{width}x{height}+{x}+{y}")

    # ==========================
    # UI
    # ==========================

    def create_ui(self):

        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = ("选择", "类型", "名称", "路径/URL", "浏览器")

        self.tree = ttk.Treeview(frame, columns=columns, show="headings")

        self.tree.heading("选择", text="☐", command=self.toggle_all)
        self.tree.heading("类型", text="类型")
        self.tree.heading("名称", text="名称")
        self.tree.heading("路径/URL", text="路径/URL")
        self.tree.heading("浏览器", text="浏览器")

        self.tree.column("选择", width=30, anchor="center")
        self.tree.column("类型", width=60, anchor="center")
        self.tree.column("名称", width=120, anchor="center")
        self.tree.column("路径/URL", width=420, anchor="w")
        self.tree.column("浏览器", width=90, anchor="center")

        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Button-1>", self.toggle_checkbox)
        self.tree.bind("<Motion>", self.show_tooltip)
        self.tree.bind("<Leave>", self.hide_tooltip)
        self.tree.bind("<Double-1>", self.inline_edit)

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="添加软件", command=self.add_app).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="添加URL", command=self.add_url).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="删除", command=self.delete_item).grid(row=0, column=2, padx=5)
        ttk.Button(btn_frame, text="启动", command=self.launch_selected).grid(row=0, column=3, padx=5)
        ttk.Button(btn_frame, text="导出BAT", command=self.export_bat).grid(row=0, column=4, padx=5)
        ttk.Button(btn_frame, text="导出VBS", command=self.export_vbs).grid(row=0, column=5, padx=5)

        self.refresh_tree()

    # ==========================
    # Tooltip
    # ==========================

    def show_tooltip(self, event):

        row = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)

        if col != "#4":
            self.hide_tooltip(event)
            return

        if not row:
            return

        text = self.tree.item(row)["values"][3]

        if self.tooltip:
            self.tooltip.destroy()

        self.tooltip = tk.Toplevel(self.root)
        self.tooltip.wm_overrideredirect(True)

        label = tk.Label(
            self.tooltip,
            text=text,
            background="yellow",
            relief="solid",
            borderwidth=1
        )

        label.pack()

        x = event.x_root + 15
        y = event.y_root + 10

        self.tooltip.wm_geometry(f"+{x}+{y}")

    def hide_tooltip(self, event):

        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

    # ==========================
    # 表格内编辑
    # ==========================

    def inline_edit(self, event):

        row = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)

        if col not in ("#3", "#4"):
            return

        bbox = self.tree.bbox(row, col)

        if not bbox:
            return

        x, y, w, h = bbox

        column_index = int(col.replace("#", "")) - 1
        value = self.tree.item(row)["values"][column_index]

        entry = tk.Entry(self.tree)
        entry.insert(0, value)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus()

        def save(event=None):

            new_value = entry.get()

            values = list(self.tree.item(row)["values"])
            values[column_index] = new_value

            self.tree.item(row, values=values)

            entry.destroy()

        entry.bind("<Return>", save)
        entry.bind("<FocusOut>", save)

    # ==========================
    # 删除
    # ==========================

    def delete_item(self):

        item = self.tree.selection()

        if not item:
            return

        values = self.tree.item(item)["values"]

        if values[1] == "软件":

            for a in self.data["apps"]:
                if a["path"] == values[3]:
                    self.data["apps"].remove(a)
                    break

        elif values[1] == "URL":

            for u in self.data["urls"]:
                if u["url"] == values[3]:
                    self.data["urls"].remove(u)
                    break

        self.save_config()
        self.refresh_tree()


    # ==========================
    # 勾选
    # ==========================

    def toggle_checkbox(self, event):

        col = self.tree.identify_column(event.x)

        if col != "#1":
            return

        row = self.tree.identify_row(event.y)

        if not row:
            return

        values = list(self.tree.item(row, "values"))
        values[0] = "☑" if values[0] == "☐" else "☐"

        self.tree.item(row, values=values)

    # ==========================
    # 全选
    # ==========================

    def toggle_all(self):

        items = self.tree.get_children()

        if not items:
            return

        all_checked = True

        for item in items:
            if self.tree.item(item)["values"][0] != "☑":
                all_checked = False
                break

        new_state = "☐" if all_checked else "☑"

        for item in items:
            values = list(self.tree.item(item, "values"))
            values[0] = new_state
            self.tree.item(item, values=values)

        self.tree.heading("选择", text=new_state, command=self.toggle_all)

    # ==========================
    # 配置
    # ==========================

    def load_config(self):

        if os.path.exists(CONFIG_FILE):

            with open(CONFIG_FILE, "r", encoding="utf8") as f:
                self.data = json.load(f)

    def save_config(self):

        with open(CONFIG_FILE, "w", encoding="utf8") as f:
            json.dump(self.data, f, indent=4, ensure_ascii=False)

    # ==========================
    # 刷新
    # ==========================

    def refresh_tree(self):

        for i in self.tree.get_children():
            self.tree.delete(i)

        for app in self.data["apps"]:

            self.tree.insert("", "end",
                             values=("☐", "软件", app["name"], app["path"], ""))

        for url in self.data["urls"]:

            self.tree.insert(
                "",
                "end",
                values=("☐", "URL",
                        url.get("name", ""),
                        url["url"],
                        url["browser"])
            )

    # ==========================
    # 添加软件
    # ==========================

    def add_app(self):

        path = filedialog.askopenfilename()

        if not path:
            return

        win = tk.Toplevel(self.root)
        win.title("软件名称")

        self.center_child(win, 350, 150)
        win.grab_set()

        ttk.Label(win, text="名称").pack(pady=5)

        name_entry = ttk.Entry(win, width=40)
        name_entry.pack()

        self.root.after(100, name_entry.focus_set)

        def save():

            name = name_entry.get().strip()

            if not name:
                messagebox.showerror("错误", "名称不能为空")
                return

            self.data["apps"].append({
                "name": name,
                "path": path
            })

            self.save_config()
            self.refresh_tree()

            win.destroy()

        name_entry.bind("<Return>", lambda e: save())

        ttk.Button(win, text="保存", command=save).pack(pady=10)

    # ==========================
    # 添加URL
    # ==========================

    def add_url(self):

        win = tk.Toplevel(self.root)
        win.title("添加URL")

        self.center_child(win, 380, 260)
        win.grab_set()

        ttk.Label(win, text="名称").pack(pady=5)
        name_entry = ttk.Entry(win, width=40)
        name_entry.pack()

        self.root.after(100, name_entry.focus_set)

        ttk.Label(win, text="URL").pack(pady=5)
        url_entry = ttk.Entry(win, width=40)
        url_entry.pack()

        ttk.Label(win, text="浏览器").pack(pady=5)

        browser = ttk.Combobox(win, state="readonly")
        browser["values"] = ("default", "chrome", "edge", "firefox")
        browser.current(0)
        browser.pack()
        browser.bind("<Return>", lambda e: save())

        def save():

            url = url_entry.get().strip()

            if not url:
                messagebox.showerror("错误", "URL不能为空")
                return

            self.data["urls"].append({
                "name": name_entry.get(),
                "url": url,
                "browser": browser.get()
            })

            self.save_config()
            self.refresh_tree()

            win.destroy()

        ttk.Button(win, text="保存", command=save).pack(pady=10)

    # ==========================
    # 启动
    # ==========================

    def launch_selected(self):

        for item in self.tree.get_children():

            values = self.tree.item(item)["values"]

            if values[0] != "☑":
                continue

            type_ = values[1]
            target = values[3]
            browser = values[4]

            if type_ == "软件":

                subprocess.Popen(target)

            elif type_ == "URL":

                if browser == "default":

                    webbrowser.open(target)

                elif browser == "chrome":

                    subprocess.Popen(
                        [r"C:\Program Files\Google\Chrome\Application\chrome.exe", target]
                    )

                elif browser == "edge":

                    subprocess.Popen(
                        [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", target]
                    )

                elif browser == "firefox":

                    subprocess.Popen(
                        [r"C:\Program Files\Mozilla Firefox\firefox.exe", target]
                    )

    # ==========================
    # 导出BAT
    # ==========================

    def export_bat(self):

        items = []

        for item in self.tree.get_children():

            values = self.tree.item(item)["values"]

            if values[0] != "☑":
                continue

            type_ = values[1]
            target = values[3]
            browser = values[4]

            if type_ == "软件":

                items.append(f'start "" "{target}"')

            elif type_ == "URL":

                if browser == "default":
                    items.append(f'start "" "{target}"')

                elif browser == "chrome":
                    items.append(
                        f'start "" "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" "{target}"'
                    )

                elif browser == "edge":
                    items.append(
                        f'start "" "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe" "{target}"'
                    )

                elif browser == "firefox":
                    items.append(
                        f'start "" "C:\\Program Files\\Mozilla Firefox\\firefox.exe" "{target}"'
                    )

        if not items:
            messagebox.showwarning("提示", "请先勾选项目")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".bat",
            filetypes=[("BAT文件", "*.bat")]
        )

        if not path:
            return

        with open(path, "w", encoding="gbk") as f:

            f.write("@echo off\n")
            f.write("\n".join(items))

        messagebox.showinfo("完成", "BAT脚本已导出")

    # ==========================
    # 导出VBS
    # ==========================

    def export_vbs(self):

        items = []

        for item in self.tree.get_children():

            values = self.tree.item(item)["values"]

            if values[0] != "☑":
                continue

            type_ = values[1]
            target = values[3]
            browser = values[4]

            if type_ == "软件":

                items.append(f'WshShell.Run """{target}"""')

            elif type_ == "URL":

                if browser == "default":
                    items.append(f'WshShell.Run "{target}"')

                elif browser == "chrome":
                    items.append(
                        f'WshShell.Run """C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"" ""{target}"""'
                    )

                elif browser == "edge":
                    items.append(
                        f'WshShell.Run """C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"" ""{target}"""'
                    )

                elif browser == "firefox":
                    items.append(
                        f'WshShell.Run """C:\\Program Files\\Mozilla Firefox\\firefox.exe"" ""{target}"""'
                    )

        if not items:
            messagebox.showwarning("提示", "请先勾选项目")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".vbs",
            filetypes=[("VBS文件", "*.vbs")]
        )

        if not path:
            return

        with open(path, "w", encoding="utf-8") as f:

            f.write('Set WshShell = CreateObject("WScript.Shell")\n')
            f.write("\n".join(items))

        messagebox.showinfo("完成", "VBS脚本已导出")


if __name__ == "__main__":

    root = tk.Tk()
    app = LauncherApp(root)
    root.mainloop()
