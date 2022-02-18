#########################
##### Version 0.1.1 #####
#########################

import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import datetime
import jpholiday

dir_name = os.path.abspath(os.path.dirname(__file__))
week_list = ["月", "火", "水", "木", "金", "土", "日"]


def button1_clicked():
    file_path = filedialog.askopenfilename(
        initialdir=dir_name, filetypes=[("Excel Files", "*xlsx")]
    )
    if file_path:
        name1.set(file_path)


def button2_clicked():
    file_path: str = filedialog.askopenfilename(
        initialdir=dir_name, filetypes=[("Excel Files", "*xlsx")]
    )
    if file_path:
        name2.set(file_path)


def button3_clicked():
    schedule_book = openpyxl.load_workbook(name1.get())
    template_book = openpyxl.load_workbook(name2.get())
    template_sheet_names = template_book.sheetnames
    if len(schedule_book.sheetnames) == 1 and len(template_sheet_names) == 1:
        df = pd.read_excel(name1.get(), sheet_name=0)
        doctors_in_schedule = df["担当"].unique().tolist()

        for i in range(df.shape[0]):
            row_list = df.iloc[i].to_numpy()
            new_work_title = ""
            if row_list[1] == "当日直":
                df.iloc[i, 1] = "当直"
                new_work_title = "日直"
            elif row_list[1] == "日当日直":
                df.iloc[i, 1] = "日当直"
                new_work_title = "日直"
            elif row_list[1] == "当直＋午前":
                df.iloc[i, 1] = "当直"
                new_work_title = "午前"
            if new_work_title != "":
                new_row = pd.Series(
                    [
                        row_list[0],
                        new_work_title,
                        row_list[2],
                        row_list[6],
                        row_list[4],
                        row_list[5],
                        row_list[6],
                        row_list[7],
                        row_list[8],
                    ],
                    index=df.columns,
                )
                df = df.append(new_row, ignore_index=True)

        df["year"] = df["開始日"].dt.strftime("%Y").map(lambda str: int(str))
        df["month"] = df["開始日"].dt.strftime("%m").map(lambda str: int(str))
        df["date"] = df["開始日"].dt.day
        df["week_index"] = df["開始日"].dt.weekday
        df["weekday"] = df["week_index"].map(lambda index: week_list[index])
        df["doctor"] = df["担当"]
        df["value"] = df[["施設名", "仕事内容"]].apply(" ".join, axis=1)
        df = df.pivot(
            index=["year", "month", "date", "weekday"], columns="doctor", values="value"
        )
        df = df.rename_axis().reset_index()
        date_count = df.shape[0]

        year = df["year"].mode().tolist()
        month = df["month"].mode().tolist()

        ws = template_book[template_sheet_names[0]]
        ws["B2"].value = "ネーベン表 " + str(year[0]) + " 年 " + str(month[0]) + " 月"
        num_del_rows = 34 - date_count
        if num_del_rows > 0:
            ws.delete_rows(idx=6, amount=num_del_rows)

        num_col = ws.max_column
        for i in range(date_count):
            if df["month"][i] != month or df["date"][i] == 1:
                ws.cell(5 + i, 2).value = str(df["month"][i]) + "/" + str(df["date"][i])
            else:
                ws.cell(5 + i, 2).value = str(df["date"][i])
            ws.cell(5 + i, 3).value = df["weekday"][i]

        for i in range(4, num_col - 2):
            doctor = ws.cell(4, i).value
            if doctor in doctors_in_schedule:
                for j in range(date_count):
                    value = df[doctor][j]
                    ws.cell(5 + j, i).value = value
                    if pd.isna(value):
                        ws.cell(5 + j, i).fill = PatternFill(fill_type=None)

        for i in range(date_count):
            weekday = df["weekday"][i]
            date = datetime.datetime(df["year"][i], df["month"][i], df["date"][i])
            if weekday == "日" or jpholiday.is_holiday(date):
                for j in range(2, num_col):
                    ws.cell(5 + i, j).fill = PatternFill(
                        patternType="solid", fgColor="ff3333"
                    )
            if weekday == "土":
                for j in range(2, num_col):
                    ws.cell(5 + i, j).fill = PatternFill(
                        patternType="solid", fgColor="ffff00"
                    )

        dest_path = os.path.join(dir_name, "output.xlsx")
        message.set(dest_path)
        template_book.save(dest_path)

    else:
        message.set("入力されたファイルの形式が仕様に合いません")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("当直表作りお助けプログラム")
    root.geometry("600x300")

    frame1 = ttk.Frame(root, padding=10)
    frame1.grid()

    set1 = tk.StringVar()
    set1.set("スケジュールファイル: ")
    label1 = ttk.Label(frame1, textvariable=set1)
    label1.grid(row=0, column=0)

    name1 = tk.StringVar()
    entry1 = ttk.Entry(frame1, textvariable=name1, width=50)
    entry1.grid(row=0, column=1)

    button1 = ttk.Button(frame1, text="参照", command=button1_clicked)
    button1.grid(row=0, column=2)

    frame2 = ttk.Frame(root, padding=10)
    frame2.grid()

    set2 = tk.StringVar()
    set2.set("テンプレートファイル: ")
    label2 = ttk.Label(frame2, textvariable=set2)
    label2.grid(row=0, column=0)

    name2 = tk.StringVar()
    entry2 = ttk.Entry(frame2, textvariable=name2, width=50)
    entry2.grid(row=0, column=1)

    button2 = ttk.Button(frame2, text="参照", command=button2_clicked)
    button2.grid(row=0, column=2)

    frame3 = ttk.Frame(root, padding=10)
    frame3.grid()

    button3 = ttk.Button(frame3, text="実行", command=button3_clicked)
    button3.grid(row=0, column=0)

    message = tk.StringVar()
    entry3 = ttk.Entry(frame3, textvariable=message, width=50)
    entry3.grid(row=0, column=1)

    root.mainloop()
