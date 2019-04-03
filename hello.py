from flask import Flask,render_template,request,redirect,session,url_for
from flask import send_file, send_from_directory,app
from .fuzzycompare import compare2list
import openpyxl
from openpyxl import Workbook
import sys, os
from flask_bootstrap import Bootstrap
from flask_wtf import FlaskForm
from wtforms import TextAreaField, SubmitField
from wtforms.fields.html5 import EmailField
from threading import Thread
from wtforms.validators import DataRequired,InputRequired,Email
# Flask 用这个参数确定应用的位置，进而找到应用中其他文件的位置，例如图像和模板。
app = Flask(__name__)



#数据库
from flask_sqlalchemy import SQLAlchemy
basedir = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] ='sqlite:///' + os.path.join(basedir, 'data.sqlite')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)


class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(64), unique=True)
    cnt= db.Column(db.Integer)
    def __repr__(self):
        return '<Role %r>' % self.name



sys.path.append(os.getcwd())
#表单key
bootstrap = Bootstrap(app)
app.config['SECRET_KEY'] = 'fuzzyflask' #表单类使用为了增强安全性，密钥不应该直接写入源码，而要保存在环境变量中。这一技术在第 7 章介绍。

#邮件
from flask_mail import Mail

app.config.update(dict(
    DEBUG = True,
    MAIL_SERVER = 'smtp.qq.com',#'smtp.gmail.com',
    MAIL_PORT = '465',#
    MAIL_USE_TLS = False,#587
    MAIL_USE_SSL = True,#465
    MAIL_USERNAME = '1274543351@qq.com',#os.environ.get('MAIL_USERNAME'),
    MAIL_PASSWORD = 'lvjgwyfpbzatgibc',# os.environ.get('MAIL_PASSWORD') export MAIL_PASSWORD=
    MAIL_DEFAULT_SENDER	= '1274543351@qq.com',
))

mail = Mail(app)#connection refused

from flask_mail import Message
app.config['FLASKY_MAIL_SUBJECT_PREFIX'] = '[matchfuzzydata]'
app.config['FLASKY_MAIL_SENDER'] = 'matchfuzzydata <1274543351@qq.com>'

#send_email() 函数的参数分别为收件人地址、主题、渲染邮件正文的模板和关键字参数列表。
# 调用者传入的关键字参数将传给 render_template() 函数，作为模板变量提供给
# 模板使用，用于生成电子邮件正文。 todo我们可以把执行 send_async_email() 函数的操作发给 Celery 任务队列。
def send_email(to, subject, template, **kwargs):
    msg = Message(app.config['FLASKY_MAIL_SUBJECT_PREFIX'] + subject,sender=app.config['FLASKY_MAIL_SENDER'], recipients=[to])
    msg.body = render_template(template + '.txt', **kwargs)
    msg.html = render_template(template + '.html', **kwargs)
    with app.open_resource("donefiles/"+to+".xlsx") as fp:
        msg.attach("donefiles/"+to+".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fp.read())
    thr = Thread(target=send_async_email, args=[app, msg])
    thr.start()
    return thr

def send_async_email(app, msg):
    with app.app_context():
        mail.send(msg)


#@app.route('/')
# 定义路由的最简便方式，是使用应用实例提供的app.route 装饰器
@app.route('/',methods=['GET','POST'] )
def index():
    #return redirect('http://youtube.com')
    form = NameForm()
    if form.validate_on_submit(): #刷新不重复提交post,这里没啥用……s
        session['leftl'] = form.leftl.data
        session['rightl'] = form.leftl.data
        return redirect(url_for('index'))
    return render_template('index.html',form=form,leftl=session.get('leftl'),rightl=session.get('rightl'))

@app.route('/about',methods=['GET','POST'] )
def about():
    return render_template('about.html')



#主程序
@app.route('/process',methods=['GET','POST'] )
def process():
    form = NameForm()
    #resp.set_cookie('passwd', '123456') make response
    #string to list flask处理request的方法不一样
    leftl=request.form.get('leftl').split("\r\n")
    rightl=request.form.get('rightl').split("\r\n")
    resultdic={}
    compare2list(leftl,rightl,resultdic)
    #print(request.cookies)
    #print(request.headers)
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
    #按客户邮箱命名文件
    wb.save("donefiles/"+form.email.data+".xlsx")
    #数据库处理邮件地址，判断使用次数3次
    address=form.email.data
    user=User.query.filter_by(email=form.email.data).first()
    if user is None:#新客户
        user=User(email=form.email.data)
        db.session.add(user)
        user.cnt=1
        db.session.add(user)
        db.session.commit()
        sendfile(address)
    elif user.cnt<3:
        user.cnt+=1
        db.session.add(user)
        db.session.commit()
        sendfile(address)
    else:#超过3次,付费后send file
        user.cnt+=1
        db.session.add(user)
        db.session.commit()
        print("first wechat payment part")
        sendfile(address)
        #todo 付款环节

    #返回处理后的结果 ，更新form显示
    return render_template('results.html',form=form,leftl=leftl,rightl=rightl)

def sendfile(address):
    print("file sending ")
    #发送邮件 todo send_email() 函数的参数分别为收件人地址、主题、渲染邮件正文的模板和关键字参数列表。
    #这里发给自己测试
    send_email('leemoispace@gmail.com', 'fuzzy match done!','mail/filedone', user=user)
    #多线程，高并发


    print("file sent to "+address)
    os.remove("donefiles/"+address+".xlsx")



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


class NameForm(FlaskForm):
    leftl = TextAreaField('paste the nonstandard data you want to match', validators=[DataRequired()])
    rightl = TextAreaField('paste the standard data ', validators=[DataRequired()])
    email = EmailField("Email address where match result file will be sent to.",  validators=[InputRequired("Please enter your email address"), Email("Please enter your email address.")])
    submit = SubmitField('Start match and send result to me!')


@app.route('/user/<name>')#动态路由获得参数例子
def user(name):
    return render_template('user.html', name=name)

@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'),404

@app.errorhandler(500)
def internal_server_error(e):
    return render_template('500.html'), 500
