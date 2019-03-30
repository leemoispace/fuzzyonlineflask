from flask import Flask,render_template,request,redirect
from flask import send_file, send_from_directory,app
from .fuzzycompare import compare2list
import openpyxl
from openpyxl import Workbook

import sys, os
sys.path.append(os.getcwd())



app = Flask(__name__)
# Flask 用这个参数确定应用的位置，进而找到应
# 用中其他文件的位置，例如图像和模板。

#@app.route('/')
# 定义路由的最简便方式，是使用应用实例提供的app.route 装饰器
#def index():
#    return '<h1>Hello World!</h1>'

@app.route('/user/<name>')#动态路由
def user(name):
    return '<h1>Hello, {}!</h1>'.format(name)
    

@app.route('/browser')
def browser():
    user_agent = request.headers.get('User-Agent')
    return '<p>Your browser is {}</p>'.format(user_agent)

@app.route('/')
def index():
    #return redirect('http://youtube.com')
    return render_template('index.html')


@app.route('/submit',methods=['GET','POST'] )
def processing():
    #string to list flask的方法不一样
    leftl=request.form.get('from').split("\r\n")
    rightl=request.form.get('standard').split("\r\n")
    #这里可以处理数据了
    resultdic={}
    compare2list(leftl,rightl,resultdic)

    print(resultdic)
    #resultdic怎么呈现到模板渲染，从服务器到客户端

    #生成excel
    wb = Workbook()
    ws = wb.active
    next_row=1
    for (key,value) in resultdic.items():
        ws.cell(column=1 , row=next_row, value=key)
        ws.cell(column=2 , row=next_row, value=value[0])
        ws.cell(column=3 , row=next_row, value=value[1][0])
        ws.cell(column=4 , row=next_row, value=value[1][1])
        next_row += 1
    wb.save("donefiles/test.xlsx") #按客户命名
    #下载后删除 hook?

    #返回渲染处理后的结果
    return render_template('results.html')
    #todo 多个extract




@app.route("/download/<filename>", methods=['GET'])
def download_file(filename):
    # 需要知道2个参数, 第1个参数是本地目录的path, 第2个参数是文件名(带扩展名)
    directory = os.getcwd()  # 假设在当前目录
    return send_from_directory(directory, filename, as_attachment=True)

@app.route("/<filepath>", methods=['GET'])
#/donefiles/test.xlsx
def download_file_static(filepath):
    # 此处的filepath是文件的路径，但是文件必须存储在static文件夹下， 比如images\test.jpg
    return app.send_static_file(filepath) 