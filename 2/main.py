import pandas as pd
from datetime import datetime
import sqlite3
from sqlite3 import Error


def read_data():
    data = {}
    df = pd.read_excel("data.xlsx")
    df = df.sort_values("Грузополучатель")
    for i, row in df.iterrows():
        data[row["Грузополучатель"]] = {}
        data[row["Грузополучатель"]]["Позиции"] = []
        data[row["Грузополучатель"]]["{{NUMBER}}"] = len(list(data.keys()))
    for i, row in df.iterrows():
        data[row["Грузополучатель"]]["{{TO}}"] = row["Грузополучатель"]
        data[row["Грузополучатель"]]["{{FROM}}"] = row["Грузоотправитель"]
        data[row["Грузополучатель"]]["{{DATE}}"] = datetime.today().date().strftime("%d.%m.%y")
        data[row["Грузополучатель"]]["Позиции"].append([len(data[row["Грузополучатель"]]["Позиции"]) + 1, row["Товар"], row["Ед."],  row["Кол-во"]])
        data[row["Грузополучатель"]]["CODE"] = f"AB{data[row['Грузополучатель']]['{{NUMBER}}']}_{datetime.today().date().strftime('%d.%m.%y')}"
    for k, v in data.items():
        res = []
        for item in v["Позиции"]:
            res.append(f"{item[0]}) {item[1]} {item[2]} {item[3]}")
        res = ", ".join(res)
        v["Позиции"] = res
        del v["{{DATE}}"]
    return data


def sql_connection():
    try:
        con = sqlite3.connect('mydatabase1.db')
        return con
    except Error:
        print(Error)


def sql_table(con):
    cursorObj = con.cursor()
    ex = 'CREATE TABLE docs(id text PRIMARY KEY, is_done boolean, date_done datetime)'
    cursorObj.execute(ex)
    con.commit()


def sql_insert(con, entities):
    cursorObj = con.cursor()
    ex = 'INSERT INTO docs(id, is_done, date_done) VALUES(?, ?, ?)'
    cursorObj.execute(ex, entities)
    con.commit()


def main():
    con = sql_connection()
    sql_table(con)
    for k, v in read_data().items():
        entities = (v["CODE"], False, None)
        sql_insert(con, entities)


if __name__ == "__main__":
    main()
