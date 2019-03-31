from flask import Flask,render_template,request,redirect
from flask import send_file, send_from_directory,app
from .fuzzycompare import compare2list
import openpyxl
from openpyxl import Workbook
import sys, os
from flask_bootstrap import Bootstrap
sys.path.append(os.getcwd())

app = Flask(__name__)
bootstrap = Bootstrap(app)

# Flask 用这个参数确定应用的位置，进而找到应
# 用中其他文件的位置，例如图像和模板。

#@app.route('/')
# 定义路由的最简便方式，是使用应用实例提供的app.route 装饰器
@app.route('/')
def index():
    #return redirect('http://youtube.com')

    return render_template('index.html')

@app.route('/user/<name>')#动态路由获得参数例子
def user(name):
    return render_template('user.html', name=name)

@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'),404

@app.errorhandler(500)
def internal_server_error(e):
    return render_template('500.html'), 500

@app.route('/process',methods=['GET','POST'] )
def processing():
    #resp.set_cookie('passwd', '123456') make response

    #string to list flask处理request的方法不一样
    leftl=request.form.get('from').split("\r\n")
    rightl=request.form.get('standard').split("\r\n")
    resultdic={}
    compare2list(leftl,rightl,resultdic)
 
    print(request.cookies)
    print(request.headers)

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
    #按客户命名
    wb.save("donefiles/test.xlsx")

    #下载后删除 hook

    #返回处理后的结果 ，更新form显示
    return render_template('results.html',leftl=leftl,rightl=rightl)

    #todo 多个extract


@app.route("/download", methods=['GET'])
def download():
    filename="test.xlsx"
    if os.path.isfile(os.path.join('donefiles/', filename)):
        # # 需要知道2个参数, 第1个参数是本地目录的path, 第2个参数是文件名(带扩展名)
        # directory = os.path.join(os.getcwd(),'donefiles/')
        # print(directory,filename)
        # return send_from_directory(directory,filename,as_attachment=True)
        
        #return send_file("donefiles/test.xlsx",cache_timeout=-1)
        return app.send_static_file("donefiles/test.xlsx")