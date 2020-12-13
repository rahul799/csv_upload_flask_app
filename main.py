from flask import Flask, render_template, request, redirect, url_for
import os
from os.path import join, dirname, realpath
import pandas as pd
from time import sleep
import string
import xlsxwriter
from xlsxwriter import Workbook
from time import sleep
import xlrd
import requests
from xlrd import open_workbook
import xlutils
from xlutils.copy import copy

class Run:
    def __init__(self):
        self.item_id = ''
        self.customer_id=[]
        self.status=[]
        self.name='asd'
        self.contact='9716567856'
        self.email='you@gmail.com'
        self.invoice_id=[]
        self.col_name=[]
        self.col_mob=[]
        self.col_email=[]
        self.col_cost=[]
        self.col_amount=[]
        self.amount_paid=[]
        self.index=0
        self.workbook=Workbook("/home/ananya/Downloads/demo.xlsx")
        self.worksheet=self.workbook.add_worksheet("School Payment")
        try:
            self.read_sheet()
            #print("yes")
            #self.get_item()
            self.generate_invoice()
            self.check_payment()
            self.update_sheet()
            self.workbook.close()
        except Exception as e:
            print(e)
                    
            
    def read_sheet(self):
        wb = xlrd.open_workbook("/home/ananya/Downloads/data.xlsx")
        sheet = wb.sheet_by_index(0)
        for i in range(1,sheet.nrows):
            self.col_name.append(sheet.cell_value(i,0))
            self.col_mob.append(sheet.cell_value(i,1))
            self.col_email.append(sheet.cell_value(i,2))
            self.col_cost.append(sheet.cell_value(i,3))
            self.col_amount.append(sheet.cell_value(i,4))
			
            self.name=sheet.cell_value(i,0)
            self.contact=sheet.cell_value(i,1)
            self.email=sheet.cell_value(i,2)
            self.add_customer()
        
    def add_customer(self):
        url="https://api.razorpay.com/v1/customers"
        PARAMS = {
        'name':self.name,'contact':self.contact,'email':self.email}
        r = requests.post(url = url,auth=('rzp_test_nszN0TzwchFKDB','tDDE5guqiyPCqNJlrN0BGvWJ'), params = PARAMS)
        data = r.json()
        print(data)
        self.customer_id.append(data['id'])

    def get_item(self):
        url="https://api.razorpay.com/v1/items"
        print(self.col_name[self.index])
        print(self.col_amount[self.index])
        PARAMS ={'name':self.col_name[self.index],'amount':int(self.col_amount[self.index]),'currency':'INR'}
        r = requests.post(url = url,auth=('rzp_test_nszN0TzwchFKDB','tDDE5guqiyPCqNJlrN0BGvWJ'),params=PARAMS)
        data = r.json()
        print(data)
        item=(data['id'])
        self.item_id=item
        print(self.item_id)

        
    def generate_invoice(self):
        print(self.customer_id)
        for i in self.customer_id:
            self.get_item()
            print("chutmarike")
            self.index=self.index+1
            print(i)
            url="https://api.razorpay.com/v1/invoices"
            PARAMS = {
            'customer_id':i,
            'line_items': [
        {
          "item_id": self.item_id
        }
      ],   
            "sms_notify": 1,
            "email_notify": 1
            }
            r = requests.post(url = url,auth=('rzp_test_nszN0TzwchFKDB','tDDE5guqiyPCqNJlrN0BGvWJ'), json = PARAMS)
            data = r.json()
            self.invoice_id.append(data['id'])
            print(data)
            
    def check_payment(self):
        url="https://api.razorpay.com/v1/invoices/"
        for flag in self.invoice_id:
            r = requests.get(url = url+flag,auth=('rzp_test_nszN0TzwchFKDB','tDDE5guqiyPCqNJlrN0BGvWJ'))
            data = r.json()
            print(data['status'])
            self.status.append(data['status'])
            self.amount_paid.append(data['amount_paid'])
            
        

    def update_sheet(self):
        try:
            self.worksheet.write(0,0,'student_name')
            self.worksheet.write(0,1,'contact_no')
            self.worksheet.write(0,2,'email')
            self.worksheet.write(0,3,'cost')
            self.worksheet.write(0,4,'due')
            ind=0
            row=1
            print(self.status)
            for flag in self.status:
                print(ind)
                self.worksheet.write(row,0,self.col_name[ind])
                self.worksheet.write(row,1,self.col_mob[ind])
                self.worksheet.write(row,2,self.col_email[ind])
                cost = int(self.col_cost[ind])
                due = int(self.col_amount[ind])
                
                if flag=='paid':
                    print("paid")
                    final=cost+due-(self.amount_paid[ind])
                else:
                    print("failed")
                    final=cost+due
                self.worksheet.write(row,3,cost)
                self.worksheet.write(row,4,final)
                ind=ind+1
                row=row+1
        except Exception as e:
            print (e)

app = Flask(__name__)

# enable debugging mode
app.config["DEBUG"] = True

# Upload folder
UPLOAD_FOLDER = 'static/files'
app.config['UPLOAD_FOLDER'] =  UPLOAD_FOLDER


# Root URL
@app.route('/')
def index():
     # Set The upload HTML template '\templates\index.html'
    return render_template('index.html')


# Get the uploaded files
@app.route("/", methods=['POST'])
def uploadFiles():
      # get the uploaded file
      uploaded_file = request.files['file']
      if uploaded_file.filename != '':
           file_path = os.path.join(app.config['UPLOAD_FOLDER'], uploaded_file.filename)
          # set the file path
           uploaded_file.save(file_path)
          # save the file
      return redirect(url_for('index'))

def parseCSV(file_path):
    csvData = pd.read_csv(file_path,names = col_names, header = None)
    for i in csvData.iterrows():
        print(i)


if (__name__ == "__main__"):
     app.run(port = 5000)