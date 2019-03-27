from flask import Flask
app = Flask(__name__)
# Flask 用这个参数确定应用的位置，进而找到应
# 用中其他文件的位置，例如图像和模板。

@app.route('/')
# 定义路由的最简便方式，是使用应用实例提供的app.route 装饰器
def index():
    return '<h1>Hello World!</h1>'

@app.route('/user/<name>')
def user(name):
    return '<h1>Hello, {}!</h1>'.format(name)