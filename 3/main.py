from tkinter import Tk, BOTH
from tkinter.ttk import Frame, Button
from tkinter import messagebox as mbox
from pyzbar import pyzbar
import cv2
import sqlite3
from sqlite3 import Error
from datetime import datetime


def draw_barcode(decoded, image):
    image = cv2.rectangle(image, (decoded.rect.left, decoded.rect.top), (decoded.rect.left + decoded.rect.width, decoded.rect.top + decoded.rect.height), color=(0, 255, 0), thickness=5)
    return image


def decode(image):
    code = False
    decoded_objects = pyzbar.decode(image)
    for obj in decoded_objects:
        image = draw_barcode(obj, image)
        code = (obj.data).decode("utf-8")
    return [image, code]


def sql_connection1():
    try:
        con = sqlite3.connect('mydatabase1.db')
        return con
    except Error:
        print(Error)


def sql_connection():
    try:
        con = sqlite3.connect('mydatabase.db')
        return con
    except Error:
        print(Error)


def get_status(con, code):
    cursorObj = con.cursor()
    ex = f"SELECT is_done FROM docs WHERE id='{code}'"
    cursorObj.execute(ex)
    res = cursorObj.fetchall()[0][0]
    con.commit()
    return res


def get_if_done(con, code):
    cursorObj = con.cursor()
    ex = f"SELECT date_done FROM docs WHERE id='{code}'"
    cursorObj.execute(ex)
    res = cursorObj.fetchall()[0][0]
    con.commit()
    return res


def get_positions(code):
    con = sql_connection()
    cursorObj = con.cursor()
    ex = f"SELECT position FROM docs WHERE document='{code}'"
    cursorObj.execute(ex)
    res = cursorObj.fetchall()[0][0]
    con.commit()
    return res


def update(code):
    con = sql_connection1()
    cursorObj = con.cursor()
    ex = f'Update docs set is_done = 1 where id = "{code}"'
    cursorObj.execute(ex)
    now = datetime.now().strftime('%d.%m.%Y, %H:%M')
    ex = f'Update docs set date_done = "{now}" where id = "{code}"'
    cursorObj.execute(ex)
    con.commit()


class Example(Frame):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.master.title("Окно")
        self.pack()
        cap = cv2.VideoCapture(0)
        while True:
            _, frame = cap.read()
            decoded_objects, code = decode(frame)
            if type(code) == str:
                break
            cv2.imshow("frame", frame)
            if cv2.waitKey(1) == ord("q"):
                break
        con = sql_connection1()
        if get_status(con, code):
            text = f"Накладная {code}\nТовары были выданы {get_if_done(con, code)}"
            mbox.showinfo("Информация", text)
        else:
            pos = get_positions(code)
            text = f"Накладная {code}\n Выдать следующие товары со склада: {pos}"
            answer = mbox.askquestion("Вопрос", text)
            print(answer)
            if answer == "yes":
                update(code)
            else:
                pass
        root.destroy()


global root
root = Tk()
ex = Example()
root.geometry("300x150+300+300")
root.mainloop()
