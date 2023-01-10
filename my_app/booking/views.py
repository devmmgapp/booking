# -*- coding: utf-8 -*-
import uuid

from datetime import date, datetime, timedelta 
from flask import jsonify, request, current_app, send_file, Blueprint
from flask import session, request, jsonify
from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from io import BytesIO 
from bson.objectid import ObjectId
from datetime import datetime
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.styles import borders
from openpyxl.styles.borders import Border
from openpyxl.styles.alignment import Alignment
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from copy import copy
# added for local time 
from datetime import datetime, timezone
# parse date 
from dateutil import parser


import pandas as pd
import gridfs
import calendar
import smtplib
import socket
import json
import os
import re
from dotenv import load_dotenv
load_dotenv()

import checkLogged
from my_app import db,  client



booking = Blueprint('booking', __name__)

#########################################################################################################
## booking
#########################################################################################################

#db
##eleaveDtl = db["eleave_dtl"]
reportMap = db["fileDirectory"]


@booking.route('/api/checkDuplicateID', methods=['POST'])
@checkLogged.check_logged
def check_duplicate_inpsection_id():
    content = request.get_json() #python data 
    _id = content['_id']
    col = db["inspectionBooking"]
    query =  { "_id": _id}
    exists = col.find_one(query)
    if (exists):
        return "t",200
    else:        
        return "f",201

@booking.route('/api/save', methods=['POST'])
@checkLogged.check_logged
def save_inspection():
    content = request.get_json() #python data 

    #content = request.data # json data 
    _id = content['_id']
    
    items = content['items']
    itemsTotal = content['itemsTotal'] 
    poList = content['poList'] 
    main = content['main']
    misc = content['misc']
    update_history = content['updateHistory']    
                   

    # get the current time and convert to string
    updated_time = datetime.now(timezone.utc).ctime()

    if (session.get("email")):
         updated_by = session.get("email")
    else:
         updated_by = "vincent.cheng@macys.com"   
    
    # Convert dictionary object into string, this is for change tracking only
    misc_str = json.dumps(misc)   
    
    #main['inspection_date'] = parser.parse(main['inspection_date'])
    main['inspection_date'] = main['inspection_date']
        
    if (update_history ==[]) :        
        update_current = { "id" :  str(uuid.uuid4()), "misc" :misc_str,  "updated_by" : updated_by, "updated_time" : updated_time, "updated_mode" : "create"}
    else:
        update_current = { "id" :  str(uuid.uuid4()), "misc" :misc_str,  "updated_by" : updated_by, "updated_time" : updated_time, "updated_mode" : "update"}

    update_history.append(update_current)    
       
    new_content = { "_id" : _id, "main" : main,  "misc" : misc,   "items" : items, "itemsTotal" : itemsTotal, 
    "poList" : poList,   "update_history" : update_history }
  
    col = db["inspectionBooking"]

    query =  { "_id": _id}
    exists = col.find_one(query)

    ## convert datetime string of change tracking into datetime object in Mongodb 
    for hist in update_history:
        hist['updated_time'] = parser.parse( hist['updated_time'])

    
    if (exists):
        change =  { "$set":  {  "main" : main, "misc": misc, "items" : items, "itemsTotal" : itemsTotal, "poList" : poList,  "update_history": update_history} }    # change     
        col.update_one(query, change)
        return "ok",200
    else:        
        x = col.insert_one(new_content)
        #print(x.inserted_id)
        return "ok",200
        

def isLegitAPI(mc_no): 
    isFound = False        
    mcTable = session.get("mcTable")
    for rec in mcTable:
        if rec['mc_no'] == mc_no:
            #print ("su, mf = {0}, {1}".format(rec["su_no"], rec["mf_no"]))
            isFound = True     
            break 
            
    return (isFound)

@booking.route('/api/search',methods=['POST'])
@checkLogged.check_logged
def search_inspection():
        
    content = request.get_json() #python data     
    ##print("Search content",content)
    _id = content['_id']
    
    col = db["inspectionBooking"]

    query =  { "_id": _id}
    results = col.find_one(query)        

    if (results):
       return  jsonify(results), 200 
    else:
       return "Error", 404 

@booking.route('/api/delete',methods=['POST'])
@checkLogged.check_logged
def delete_inspection():
        
    content = request.get_json() #python data     
    ##print("Search content",content)
     
    try:  
        
        col = db["inspectionBooking"]
        delete_log_col = db["Delete_Log"]
        _id = content['_id']

        query =  { "_id": _id}

        ## add to delete_log in Mongo
        updated_by = "development@heroku" if session.get("email") == None else session.get("email")               
        existing_inspectionBooking = col.find_one(query)
        delete_log_col.insert_one( { "_id": str(uuid.uuid4()), 
        "updated_by" : updated_by, "time":datetime.now(timezone.utc),"doc_type": "inspectionBooking", "rec" : existing_inspectionBooking })

        result = col.delete_one(query)
        if result.deleted_count > 0:
            return  "OK", 200 
        else:
            return  "Not OK", 400 
    
    #except mongoengine.errors.OperationError:           
    except Exception as e: 
        print("error", e)
        return e, 400 

@booking.route('/api/searchInspByMC',methods=['POST'])
@checkLogged.check_logged
def searchInspMC():
        
    content = request.get_json() #python data     
    mc = content['mc'] 
 
    col = db["inspectionBooking"]
    
    ## Reserverd for later use
    # search = []
    # for _filter in session['mfList']:
    #     search.append(   {  '$and': [ { 'main.su_no': { '$eq': _filter['SU'] } }, { 'main.mf_no' : { '$eq': _filter['MF'] } } ]  } )        

    query =   {
       "_id.mc" : {
       "$regex": mc,
       "$options" :'i' # case-insensitive
       } }
     
    #print(search, '$$search')     

    ##results = col.find(query).limit(5)    
    ## returns 10 at a time
    ##  to access print(rec["_id"]["mc"])    
    results = col.find(query).limit(60)     


    id_array = []
    for result in results:
        ##'id' : {uuid.uuid4()
        rec = { '_id' : result['_id'], 'main' : result['main'] }
        id_array.append(rec)        
        
    if (results):
       return  jsonify(id_array), 200 
    else:
       return "Error", 404 


 

@booking.route('/api/getMCtable',methods=['POST'])
@checkLogged.check_logged
def getMCtable():            
  
    content = request.get_json() #python data     
    su = content['su']
    mf  = content['mf']        
     
    #get MCs for the SU and MF only
    query =  {  '$and': [ { 'su_no': { '$eq': su } }, { 'mf_no' : { '$eq': mf } } ]  }          
    col = db["mcTable"]     
                    
    mc_array = []     
    results = col.find(query)
    for result in results:
        result["_id"] = str(result["_id"])
        mc_array.append(result) 
    
    session["mcTable"] = mc_array    

    if len(mc_array) > 0: 
        return  jsonify(mc_array), 200    


####################################################################################
#  Genearte Excel Report - Start 
####################################################################################


#Version 01  1/9/23
@booking.route('/printreport',methods=['POST'])
@checkLogged.check_logged
def printreport():            
    #filename when using in Heroku:
    fs = gridfs.GridFS(db)
    wb = load_workbook(filename=BytesIO(fs.get(ObjectId(rpt["file"]["fileObj"])).read()))
    rpt = reportMap.find_one ( { "report": "Leave Summary"} )

    # filename in development:
    #wb = load_workbook(filename=rpt["file"]["fileName"])
    ws = wb[rpt["file"]["wsName"]]
   
    try:
        para = json.loads(request.headers['parameters'])       
    except Exception as e:
        print (e)
        return jsonify({"error_message" : "Sorry, we failed to generate Application form"}), 501    

    # Output 
    out = BytesIO()
    wb.save(out)
    out.seek(0)

    wb.close()            
    print('sending file...')

    return send_file(out,  attachment_filename='a_file.xls', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


####################################################################################
#  Genearte Excel Report - End
####################################################################################
 



