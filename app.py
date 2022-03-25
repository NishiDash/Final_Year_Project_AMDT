from pickle import FALSE, TRUE
from flask import Flask, render_template,request
from gevent.pywsgi import WSGIServer
from matplotlib.pyplot import text
from pymysql import NULL
from datetime import date,datetime;
from colorama import Fore, init, Back, Style
import pandas as pd
import mysql.connector
from flask import * 
from model import *
from air_density import *
from make import *
from grid_frequency import *
from hub_height import *
from short_name import *
from cod import *


mydb = mysql.connector.connect(
    host = "localhost",
    user = "root",
    passwd = "cse73626N",
    database ="amdt"
)
mycursor = mydb.cursor()
mycursor.execute("use amdt")
app = Flask(__name__)

@app.route('/home_l',methods=['GET','POST'])
def home_l():
    pwd = request.form['pwd'] 
    sso_id = request.form['sso_id']
    query = "select password from users1 where sso_id = " + sso_id
    
    mycursor.execute(query)
    for r in mycursor:
        if r[0]==pwd:
            return render_template('index.html')
        else:
            return render_template('login.html',text="*Please verify your username or password")
    return render_template('register.html',text="*You are not a registered user.")          
@app.route('/home_r',methods=['GET','POST'])
def home_r():
    email = request.form['email']
    pwd = request.form['pwd'] 
    sso_id = request.form['sso_id']
    cpwd = request.form['cpwd']
    today = date.today()
    
    if(pwd!=cpwd):
        return render_template('register.html',text="*The passwords don't match..!")
    query = "select password from users1 where sso_id = " + sso_id
    if(len(pwd)<8):
        return render_template('register.html',text="*Set minimum length of password to 8")
    mycursor.execute(query)
    for r in mycursor:
        return render_template('login.html',text="*You are alredy a registered user, pls Login.!")
    query = "Insert into users1 values('"+ sso_id +"','"+ email +"','"+ pwd +"','"+str(today)+"');"
    print(query)
    mycursor.execute(query)
    mydb.commit()
    mydb.close()
    return render_template('index.html')
@app.route('/',methods=['GET','POST'])
def index():
    return render_template('LR.html') 
@app.route('/login',methods=['GET','POST'])
def login(): 
    return render_template('login.html')
@app.route('/temp',methods=['GET','POST'])
def temp():
    return render_template('temp.html')   
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
            result =model(file1,attribute)
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
        
          
                  
    return render_template('data.html',text=result)

if __name__=='__main__':
    app.run(debug=True)

        
    