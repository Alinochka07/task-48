from flask import Flask, render_template, redirect, request
from openpyxl import Workbook, load_workbook
import xlsxwriter



app = Flask(__name__)

@app.route('/')
def homepage():
    my_excel = Workbook()
    my_sheet = my_excel.active
    my_excel.save('goods.xlsx')
    page = my_excel.active
    page['A1'] = 'Goods'
    my_excel.save('goods.xlsx')
    return render_template('index.html')


@app.route('/add/', methods=['POST'])
def add_goods():
    good = request.form['good']
    # excel = load_workbook('goods.xlsx')
    excel = xlsxwriter.Workbook('goods.xlsx')
    sheet = excel.add_worksheet()
    row = 0
    column = 0
    # page = excel.active
    content = [good]

    for item in content:
        sheet.write(row+1, column, good)
        excel.close()
        row += 1
        excel.close()
    
 
    return """
        <h1>Товары добавлены</h1>
        <a href='/'>Вернуться на главную страницу</a>
    """    
