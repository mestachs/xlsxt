from xlst import ExcelTemplateProcessor
import json

with open("demo.json", "r", encoding="utf-8") as file:
    context = json.load(file)

xlst = ExcelTemplateProcessor("demo.xlsx")
xlst.process_template(context)
xlst.save("output.xlsx")
