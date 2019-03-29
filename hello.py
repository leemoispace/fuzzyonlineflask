from flask import Flask
from flask import request
from flask import redirect

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
    return redirect('http://youtube.com')

