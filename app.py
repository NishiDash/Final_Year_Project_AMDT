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

app = Flask(__name__)

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

if __name__=='__main__':
    app.run(debug=False)

        
    