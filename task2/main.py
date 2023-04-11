import pandas as pd
from datetime import datetime
from docx import Document
import re


def replace_text(paragraph, key, value):
    "Переставляет текстовые переменные в шаблоне"
    if key in paragraph.text:
        paragraph.text = paragraph.text.replace(key, value)


def read_data():
    "Читает данные из файла excel в dict"
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
        data[row["Грузополучатель"]]["{{DOCUMENT}}"] = row["Основание"]
        data[row["Грузополучатель"]]["{{DATE}}"] = datetime.today().date().strftime("%d.%m.%y")
        data[row["Грузополучатель"]]["Позиции"].append([len(data[row["Грузополучатель"]]["Позиции"]) + 1, row["Товар"], row["Ед."], row["Цена"], row["Кол-во"], row["Сумма"]])
        data[row["Грузополучатель"]]["{{GLOBAL_PRICE}}"] = sum(el[5] for el in data[row["Грузополучатель"]]["Позиции"])
        data[row["Грузополучатель"]]["{{GLOBAL_COUNT}}"] = len(data[row["Грузополучатель"]]["Позиции"])
        data[row["Грузополучатель"]]["{{DOCUMENT}}"] = row["Основание"]
    return data


def main():
    "Основная функция"
    data = read_data()
    pattern = r"{{\w{1,}}}"
    template = 'template.docx'
    n = 1
    for key, value in data.items():
        result = f'накладная_{n:02}_{key}.docx'
        template_doc = Document(template)
        for k, v in value.items():
            if re.fullmatch(pattern, k.strip()):
                for paragraph in template_doc.paragraphs:
                    replace_text(paragraph, k, str(v))
        my_table = template_doc.tables[0]
        for line in value["Позиции"]:
            row = my_table.add_row()
            i = 0
            for cell in row.cells:
                cell.text = str(line[i])
                i += 1
        template_doc.save(f"result/{result}")
        n += 1

if __name__ == "__main__":
    main()
