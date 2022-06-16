from pickle import FALSE, TRUE
from altitude import altitude
from flask import Flask, render_template,request,redirect,url_for;
from matplotlib.pyplot import text
# from pymysql import NULL
from datetime import date,datetime;
from colorama import Fore, init, Back, Style
import pandas as pd
# import mysql.connector
from flask import * 
from model import *
from air_density import *
from make import *
from grid_frequency import *
from hub_height import *
from short_name import *
from cod import *
from MA import *
from ID_comparison import *
from altitude import *
from SystemNumber import *
from APM_Wind_Turbine_ID import *
from Rotor_Diammeter import *
from rating import *
from SourceKeyGenerator import *
from SystemNumber import *

# mydb = mysql.connector.connect(
#     host = "localhost",
#     user = "root",
#     passwd = "cse73626N",
#     database ="amdt"
# )
# mycursor = mydb.cursor()
# mycursor.execute("use amdt")
app = Flask(__name__)

# @app.route('/home_l',methods=['GET','POST'])
# def home_l():
#     pwd = request.form['pwd'] 
#     sso_id = request.form['sso_id']
#     query = "select pwd from users where sso_id = " + sso_id
    
#     mycursor.execute(query)
#     for r in mycursor:
#         if r[0]==pwd:
#             return render_template('index.html')
#         else:
#             return render_template('login.html',text="*Please verify your username or password")
#     return render_template('register.html',text="*You are not a registered user.")          
# @app.route('/home_r',methods=['GET','POST'])
# def home_r():
#     email = request.form['email']
#     pwd = request.form['pwd'] 
#     sso_id = request.form['sso_id']
#     cpwd = request.form['cpwd']
#     today = date.today()
    
#     if(pwd!=cpwd):
#         return render_template('register.html',text="*The passwords don't match..!")
#     query = "select pwd from users where sso_id = " + sso_id
#     if(len(pwd)<8):
#         return render_template('register.html',text="*Set minimum length of password to 8")
#     mycursor.execute(query)
#     for r in mycursor:
#         return render_template('login.html',text="*You are alredy a registered user, pls Login.!")
#     query = "Insert into users values('"+ sso_id +"','"+ email +"','"+ pwd +"','"+str(today)+"');"
#     print(query)
#     mycursor.execute(query) 
#     mydb.commit()
#     mydb.close()
#     return render_template('index.html')
@app.route('/',methods=['GET','POST'])
def index():
    return render_template('LR.html') 
@app.route('/login',methods=['GET','POST'])
def login(): 
    return render_template('login.html')  
@app.route('/register',methods=['GET','POST'])
def register():
    return render_template('register.html')
@app.route('/index',methods=['GET','POST'])
def index_a():
    return render_template('index.html')
@app.route('/data',methods=['GET','POST'])
def data():
    if request.method == 'POST':
        
        file1=request.form['upload-file-1'] 
        attribute=request.form.get('attribute')        

        if(attribute=='air_density'):
            result = air_density(file1,attribute)
        elif (attribute=='model'):
            result = model(file1,attribute)
        elif (attribute=='make'):
            result = make(file1,attribute)
        elif (attribute=='grid_frequency'):
            result = grid_frequency(file1,attribute) 
        elif (attribute=='hub_height'):
            result = hub_height(file1,attribute)
        elif (attribute=='short_name'):
            result = short_name(file1,attribute)
        elif (attribute=='COD'):
            result = cod(file1,attribute) 
        elif (attribute=='MA'):
            result = MA(file1)
        elif(attribute=='ID'):
            result = ID_comp(file1)
        elif(attribute=='altitude'):
            result = altitude(file1)
        elif(attribute=='APM'):
            result = APM(file1)
        elif(attribute=='RD'):
            result = RotorDiamter(file1)
        elif(attribute=='RD'):
            result = RotorDiamter(file1)
        elif(attribute=='rating'):
            result = rating(file1)
        elif(attribute=='SKG'):
            result = SKG(file1)
        elif(attribute=='SN'):
            result = SystemNumber(file1, attribute)
    return render_template('data.html',text=result)
@app.route('/todo')
def todo():
    query = "select * from todo"
    res = mycursor.execute(query)
    items=[]
    print("&&_______",type(mycursor))
    for r in mycursor:
        new_item = {}
        new_item = dict({'id':str(r[0]),'title':str(r[1]),'complete':str(r[2])})
        items.append(new_item)
        print(r)
        print("***********************************") 
    return render_template('base.html', todo_list=items)
@app.route('/add',methods=["POST"])
def add():
    title = request.form.get("title")
    query = "Insert into todo(title,complete) values('"+ title +"','False');"
    res = mycursor.execute(query)
    return redirect(url_for("todo"))

@app.route('/update/<int:id>')
def update(id):
    query = "select complete from todo where id ="+str(id)
    res = mycursor.execute(query)
    for r in mycursor:
        comp = r[0]
    if comp == 'True':
        query = "update todo set complete = 'False' where id="+str(id)
    else:
        query = "update todo set complete = 'True' where id="+str(id)
    res = mycursor.execute(query)
    return redirect(url_for("todo"))
@app.route('/delete/<int:id>')
def delete(id):
    query = "DELETE FROM todo where id="+str(id)
    res = mycursor.execute(query)
    return redirect(url_for("todo"))
 
if __name__=='__main__':
    app.run(debug=False)

        
    