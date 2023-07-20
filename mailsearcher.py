import base64
import email
import email.policy
import glob
import os
import tkinter as tk
from tkinter import ttk

import dateutil.parser
import pytz
from bs4 import BeautifulSoup as bs


class MailSearcher():
    def __init__(self):
        self.result = []
        self.files = []
        self.config = self.read_config()
        self.path = os.path.join(self.config["path"], "**\*.eml")
        self.files = glob.glob(self.path, recursive=True)
        self.threading = self.config["threading"]
        self.timezone = self.config["timezone"]
        self.years = self.config["year"]
        self.years.insert(0, "All")

    # Config
    def check_config(self, config):
        # check path
        path_exist = os.path.exists(config["path"])
        if not path_exist:
            from tkinter import messagebox
            messagebox.showerror("Error", "Email directory not exist")
            from sys import exit
            exit(1)

        # check threading
        threading = config["threading"].upper()
        if threading != "TRUE" and threading != "FALSE":
            config["threading"] = False
            from tkinter import messagebox
            messagebox.showwarning("Warning", "Unknown threading config, set to False")
        else:
            config["threading"] = True if threading == "TRUE" else False

        # check timezone
        if config["timezone"] not in pytz.all_timezones:
            from tkinter import messagebox
            messagebox.showerror("Error", "Unknown timezone config.")
            from sys import exit
            exit(1)

        # check year
        yr = config["year"].upper()
        if "NOW" in yr:
            import datetime
            yr = yr.replace("NOW", str(datetime.datetime.now().year))
        yr = yr.split("-")
        config["year"] = list(range(int(yr[0]), int(yr[1]) + 1))

        return config

    def read_config(self):
        if not os.path.exists("./config.txt"):
            from tkinter import messagebox
            messagebox.showerror("Error", "Config file not exist.")
            from sys import exit
            exit(1)

        config = {}
        try:
            with open("./config.txt", 'r', encoding="utf-8") as file:
                config_from_file = file.read().replace("\n", "").split(";")
                for item in config_from_file:
                    kv = item.split("=")
                    config[kv[0]] = kv[1].strip()
        except Exception:
            from tkinter import messagebox
            messagebox.showerror("Error", "Config file format error.")
            from sys import exit
            exit(1)

        config = self.check_config(config)
        return config

    # Utilities
    def decode_str_brute_force(self, s: str, encodings=('utf-8', 'gbk', 'gb2312', 'big5', 'gb18030')) -> str:
        for encoding in encodings:
            try:
                return s.decode(encoding)
            except Exception:
                pass

    def decode_str(self, header) -> str:
        from email.header import decode_header
        text_and_charset = decode_header(header)
        text = text_and_charset[0][0]
        charset = text_and_charset[0][1]
        if charset is not None:
            try:
                decoded_text = text.decode(charset)
            except Exception:
                decoded_text = self.decode_str_brute_force(text)
        else:
            decoded_text = text
        return decoded_text

    def get_content_and_decode(self, msg: email.message.Message) -> str:
        content = msg.get_payload()
        encoding = msg.get('Content-Transfer-Encoding')
        charset = msg.get('Content-Type').split(";")[1].split("=")[1]
        if encoding == "base64":
            content = base64.b64decode(content)
        if encoding == "quoted-printable":
            import quopri
            content = quopri.decodestring(content)
        if isinstance(content, bytes):
            try:
                content = content.decode(charset)
            except Exception:
                content = self.decode_str_brute_force(content)
        return content

    def get_subject(self, msg: str) -> str:
        from email.parser import HeaderParser
        parser = HeaderParser()
        headers = parser.parsestr(msg)
        for k, v in headers.items():
            if k == "Subject":
                subject = self.decode_str(v)
                return subject

    def get_file_path(self, file):
        for item in self.result:
            if item['address'] == file[1] and item['date'][1] == file[3]:
                return item['path']

    def search(self, file: str, info: list):
        keyword, sender, year = info[0], info[1], info[2]
        keyword = keyword.upper()
        with open(file, 'r', encoding="utf-8") as f:
            msg = email.message_from_file(f, policy=email.policy.SMTP)
            # date
            dt = dateutil.parser.parse(msg['date'])
            if year != "All":
                if dt.year != int(year):
                    return
            ts = int(dt.timestamp())
            hkt = dt.astimezone(pytz.timezone(self.timezone))
            fdt = hkt.strftime("%Y.%m.%d %a %H.%M.%S")
            # sender
            if "<" not in msg['from'] and ">" not in msg['from']:
                sender_name = ""
                sender_address = msg['from'].strip()
            else:
                sender_name = msg['from'].split('<')[0].strip().replace('"', '')
                sender_address = msg['from'].split('<')[1].split('>')[0].strip()
            if sender not in sender_name and sender not in sender_address:
                return
            flag = False
            # subject
            subject = self.get_subject(str(msg))
            if isinstance(subject, bytes):
                subject = self.decode_str_brute_force(subject)
            # content
            html = ""
            plain = ""
            attachment = "No"
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    html = self.get_content_and_decode(part)
                    html = bs(html, "html.parser")
                    for data in html(['style', 'script']):
                        data.decompose()
                    html = " ".join(html.stripped_strings)
                if part.get_content_type() == "text/plain":
                    plain = self.get_content_and_decode(part)
                if part.get_content_disposition() == "attachment":
                    attachment = "Yes"
                    filename = part.get_filename()
                    if keyword in filename.upper():
                        flag = True
            html = html.upper()
            plain = plain.upper()
            if html and keyword in html:
                flag = True
            if plain and keyword in plain:
                flag = True
            if keyword in subject.upper() or keyword in sender_name.upper() or keyword in sender_address.upper():
                flag = True
            if flag:
                self.result.append({"name": sender_name, "address": sender_address, "subject": subject, "date": (ts, fdt), "attachment": attachment, "path": file})

    def linear_search(self, keyword: str, sender: str, year: str):
        for file in self.files:
            self.search(file, [keyword, sender, year])

    def multi_thread_search(self, keyword: str, sender: str, year: str):
        import threading
        threads = []
        for path in self.files:
            thread = threading.Thread(target=self.search, args=[path, [keyword, sender, year]])
            thread.start()
            threads.append(thread)
        for thread in threads:
            thread.join()


class MailSearcherGUI(MailSearcher):
    def __init__(self):
        super().__init__()
        self.last_sort = None
        self.create_gui()

    def open_file(self, event):
        selection = self.treeview.selection()
        if selection:
            item = self.treeview.item(selection[0])
            file_path = self.get_file_path(item["values"])
            os.startfile(file_path)

    def clear_result(self):
        for row in self.treeview.get_children():
            self.treeview.delete(row)

    def display_result(self):
        for item in self.result:
            row = (item["name"], item["address"], item["subject"], item["date"][1], item["attachment"])
            self.treeview.insert("", "end", values=row)

    def reset_table(self):
        # clear existing rows in table and result
        self.clear_result()
        for col in self.treeview["columns"]:
            self.treeview.heading(col, text=col)
        self.result = []
        # bring the scrollbar to top
        self.treeview.yview_moveto(0)

    def search_email(self):
        self.reset_table()

        # get input values from user
        keyword = self.keyword_entry.get().lstrip().rstrip()
        sender = self.sender_entry.get().lstrip().rstrip()
        year = self.year_var.get()

        if self.threading:
            self.multi_thread_search(keyword, sender, year)
        else:
            self.linear_search(keyword, sender, year)

        self.display_result()

    def sort_column(self):
        # get the clicked column header
        column = self.treeview.identify_column(self.treeview.winfo_pointerx() - self.treeview.winfo_rootx())
        if column == "#1":
            header = "Sender"
            key = "name"
        elif column == "#2":
            header = "Email Address"
            key = "address"
        elif column == "#3":
            header = "Subject"
            key = "subject"
        elif column == "#4":
            header = "Date"
            key = "date"
        elif column == "#5":
            header = "Attachment"
            key = "attachment"
        else:
            return
        # reset colunmn header
        for col in self.treeview["columns"]:
            if col == header:
                # skip update the clicked column header
                continue
            else:
                self.treeview.heading(col, text=col)
        # if last sort is the same as this sort, reverse the order
        if self.last_sort == header:
            self.result.reverse()
            if self.treeview.heading(column)["text"][-1] == "↓":
                self.treeview.heading(column, text=header + "↑")
            else:
                self.treeview.heading(column, text=header + "↓")
        else:
            self.result.sort(key=lambda x: x[key])
            self.treeview.heading(column, text=header + "↑")
        self.clear_result()
        self.display_result()
        # update last sort
        self.last_sort = header

    def create_gui(self):
        # create the main window
        self.root = tk.Tk()
        if os.path.exists("icon.ico"):
            self.root.iconbitmap("icon.ico")
        # self.root.minsize(800, 500)
        self.root.geometry("900x500")
        self.root.resizable(False, False)
        self.root.title("Email Searcher")

        # create the search input frame
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=10)

        # create the keyword label and entry
        keyword_label = tk.Label(input_frame, text="Keyword:")
        keyword_label.pack(side=tk.LEFT, padx=10)
        self.keyword_entry = tk.Entry(input_frame, width=30)
        self.keyword_entry.pack(side=tk.LEFT)

        # create the sender label and entry
        sender_label = tk.Label(input_frame, text="Sender:")
        sender_label.pack(side=tk.LEFT, padx=10)
        self.sender_entry = tk.Entry(input_frame, width=30)
        self.sender_entry.pack(side=tk.LEFT)

        # create the year label and dropdown
        year_label = tk.Label(input_frame, text="Year:")
        year_label.pack(side=tk.LEFT, padx=10)
        self.year_var = tk.StringVar(value="All")
        year_dropdown = ttk.Combobox(input_frame, width=10, textvariable=self.year_var, state="readonly")
        year_dropdown["values"] = self.years
        year_dropdown.pack(side=tk.LEFT)

        # create the search button
        search_button = tk.Button(self.root, text="Search", command=self.search_email)
        search_button.pack(pady=10)

        # create the table frame
        table_frame = tk.Frame(self.root)
        table_frame.pack()

        # define the table columns
        columns = ("Sender", "Email Address", "Subject", "Date", "Attachment")

        # create the table
        self.treeview = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        self.treeview.pack(side=tk.LEFT)

        self.treeview.column("Sender", anchor=tk.W, width=int(1000*0.15))
        self.treeview.column("Email Address", anchor=tk.W, width=int(1000*0.15))
        self.treeview.column("Subject", anchor=tk.W, width=int(1000*0.3))
        self.treeview.column("Date", anchor=tk.W, width=int(1000*0.15))
        self.treeview.column("Attachment", anchor=tk.W, width=int(1000*0.09))
        # set the column headings and run sort_colunmn when any colunmn header is clicked
        for col in columns:
            self.treeview.heading(col, text=col.title(), command=self.sort_column)

        self.treeview.bind("<Double-Button>", self.open_file)

        # create a vertical scrollbar
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.treeview.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # configure the table to use the vertical scrollbar
        self.treeview.configure(yscrollcommand=vsb.set)

        # start the GUI event loop
        self.root.mainloop()


mail_searcher = MailSearcherGUI()
