
from datetime import date
from flask import jsonify, request, current_app, Blueprint
from flask import render_template, session, request, redirect, url_for, jsonify, send_from_directory 

import msal
import pandas as pd
import json
import os
from dotenv import load_dotenv
load_dotenv()

import checkLogged
import requests

from my_app import database, db 

#########################################################################################################
## Gloval variables  
#########################################################################################################

eleaveDtl = db["eleave_dtl"]

#########################################################################################################
## BluePrint Declaration  
#########################################################################################################

entry = Blueprint('entry', __name__)

#########################################################################################################
## login
#########################################################################################################

def _load_cache():
    cache = msal.SerializableTokenCache()
    if session.get("token_cache"):
        cache.deserialize(session["token_cache"])
    return cache

def _save_cache(cache):
    if cache.has_state_changed:
        session["token_cache"] = cache.serialize()

def _build_msal_app(cache=None, authority=None):
    return msal.ConfidentialClientApplication(
        os.environ['CLIENT_ID'], authority=authority or os.environ['AUTHORITY'],
        client_credential=os.environ['CLIENT_SECRET'], token_cache=cache)

def _build_auth_code_flow(authority=None, scopes=None, redirect_uri=None):
    return _build_msal_app(authority=authority).initiate_auth_code_flow(
        scopes or [],
        redirect_uri or [])
        #redirect_uri=url_for("authorized2", _external=True))

def _get_token_from_cache(scope=None):
    cache = _load_cache()  # This web app maintains one cache per session
    cca = _build_msal_app(cache=cache)
    accounts = cca.get_accounts()
    if accounts:  # So all account(s) belong to the current signed-in user
        result = cca.acquire_token_silent(scope, account=accounts[0])
        _save_cache(cache)
        return result


@entry.route("/api/getPhoto/<email>")
@checkLogged.check_logged
def getPhoto(email=None):

    try:
    
        token = _get_token_from_cache(json.loads(os.environ['SCOPE']))
        if not token and not os.environ:
            return redirect(url_for("entry.login"))
    
        ## Getting photo          
        
        ##endpoint = "https://graph.microsoft.com/v1.0/me/photos/120x120/$value"

        endpoint = f"https://graph.microsoft.com/v1.0/users/{email}/photos/120x120/$value"
        ##endpoint = "https://graph.microsoft.com/v1.0/users/ken.yip@macys.com/photos/120x120/$value"
                
        photo_response = requests.get(  # Use token to call downstream service
            endpoint,
            headers={'Authorization': 'Bearer ' + token['access_token']},
            stream=True) 
        photo_status_code = photo_response.status_code
        if photo_status_code == 200:
            photo = photo_response.raw.read()
            return photo 
        else:        
            return  send_from_directory("frontend/build/static/img", "anonymous.jpg")
    except:
            return  send_from_directory("frontend/build/static/img", "anonymous.jpg")

      

@entry.route("/")
@checkLogged.check_logged
def index():    
    if not session.get("user"):
        return redirect(url_for("entry.login"))
    return render_template('index.html', user=session["user"], version=msal.__version__)

@entry.route("/login", defaults={'timeout':None}) 
@entry.route("/login/<timeout>") 
def login(timeout):
    if (timeout):
        print ("Entering login process with "  + timeout)
    # Technically we could use empty list [] as scopes to do just sign in,
    # here we choose to also collect end user consent upfront
 
    session["flow"] = _build_auth_code_flow(scopes=json.loads(os.environ['SCOPE']), redirect_uri=url_for("entry.authorized", _external=True))    
    #  auth_uri an be added with prompt=login to force sign in     

    return render_template("login.html", auth_url=session["flow"]["auth_uri"],  version=msal.__version__, timeout_message=timeout)

@entry.route(os.environ['REDIRECT_PATH'])  # Its absolute URL must match your app's redirect_uri set in AAD
def authorized():
    try:
        print("Entering " + os.environ['REDIRECT_PATH'])
        cache = _load_cache()
        result = _build_msal_app(cache=cache).acquire_token_by_auth_code_flow(
            session.get("flow", {}), request.args)
        if "error" in result:
            return render_template("auth_error.html", result=result)
        session["user"] = result.get("id_token_claims")        
        session["email"] = (result.get("id_token_claims").get('email')).lower()      
        _save_cache(cache)
    except ValueError:  # Usually caused by CSRF
        pass  # Simply ignore them
        return render_template("auth_error.html", result={"error" : "Value Error", "error_description":"Not signed in yet !!"})    
    return redirect(url_for("entry.index"))

  

@entry.route("/logout")
def logout():
    session.clear()  # Wipe out user and its token cache from session
    return redirect(  # Also logout from your tenant's web session
        os.environ['AUTHORITY'] + "/oauth2/v2.0/logout" +
        "?post_logout_redirect_uri=" + url_for("entry.index", _external=True))


@entry.route("/graphcall")
@checkLogged.check_logged
def graphcall():
    token = _get_token_from_cache(json.loads(os.environ['SCOPE']))
    if not token:
        return redirect(url_for("entry.login"))
    graph_data = requests.get(  # Use token to call downstream service
        os.environ['ENDPOINT'],
        headers={'Authorization': 'Bearer ' + token['access_token']},
        ).json()
    return render_template('display.html', result=graph_data)



## below for Reacj JS

def getTodayDate():
    return date.today().strftime("%m/%d/%y")  ## get today's date 


@entry.route('/api/getUserProfile',methods=['POST'])
@checkLogged.check_logged
def getUserProfile():            
    sessionData = establishSessionData()
    if (sessionData):
       return  jsonify(sessionData), 200 
    else:
       return "Error", 404 
   

def establishSessionData():

    sessionData={}     

    email =""
   
    if (os.environ['ENVIRONMENT']=="HEROKU"):            
            
            endpoint = "https://graph.microsoft.com/beta/me"                    
            token = _get_token_from_cache(json.loads(os.environ['SCOPE']))

            if not token and not os.environ:
                return redirect(url_for("entry.login"))

            racf_response = requests.get(  # Use token to call downstream service
                endpoint,
                headers={'Authorization': 'Bearer ' + token['access_token']}, stream=True
                ) 
            status_code = racf_response.status_code                        

            if status_code == 200:                
               email = session["email"]
            else:
               raise Exception("It fails to validate your email.  Please contact regional PBT for assistance!")    

    else:
            email = current_app.config['APP_EMAIL'] 
    
             
    col = db["userProfile"]

    query =  { "email": email}

    results = col.find_one(query)
    #print('results', results["email"])
    # print('results', results["mf_list"])
    
    #this forces the mf_list to be generated from profile only, not API requests.   
    
    session["userName"] = f"{results['first_name']} {results['last_name']}"
    session["mfList"] = results["mf_list"]   
        
    sessionData["userProfile"] = {"email" : results["email"], "userName" : session["userName"], "ignore_submit": results["ignore_submit"],
    "environment":  os.environ["ENVIRONMENT"], 
    "databaseSchema":  "dev" if database[:3].lower() == "dev" else "prod"    
     }    

    sessionData["mfList"] = results["mf_list"]   


    sessionData["mcTable"] = ""  
    session["mcTable"] = ""  

    #get Party Table 
    search = []        
    for _filter in session['mfList']:
        search.append(   {  '$or': [ { '_id': { '$eq': _filter['SU'] } }, { '_id' : { '$eq': _filter['MF'] } } ]  } )         

    col = db["partyTable"]
    query = {'$or' : search}   
            
    party = []     
    results = col.find(query)
    for result in results:
        result["_id"] = str(result["_id"])
        party.append(result)

       
    partyTable = []     
    for pair in session['mfList']:
        for x in party:
            if x['_id'] == pair['SU']:
                pair['SU_NAME'] = x['party_name']
            if x['_id'] == pair['MF']:
                pair['MF_NAME'] = x['party_name']
        partyTable.append(pair)             

    sessionData["partyTable"] = partyTable     
    #print(partyTable)        


    col = db["qcAQL"]
    query = {}   
    aqlTable = []     
    results = col.find(query)
    for result in results:        
        aqlTable.append(result)
    
    sessionData["aqlTable"] = aqlTable


    col = db["checkList"]
    query = {}   
    checkList = []     
    results = col.find(query)
    for result in results:    
        #remove this _id as this is an object not serializable 
        result.pop('_id')    
        checkList.append(result)
    
    sessionData["checkListTemplate"] = checkList 
              
    #print("Established Session Data")

    #get QA members list     

    col = db["metaTable"]
    query = {'category': "qaList"}
    results = col.find_one(query)
 
    group = []
    for rec in results['selectionList']:          
        lead_email = rec["QALead"]
        for qa_email in rec['QAList']:            
            if qa_email["mqa"] == email:                
                group = rec['QAList']   

    sessionData["mqaMembers"] = group   

    #get Pack Type  
    col = db["metaTable"]
    query = {'category': "packType"}
    results = col.find_one(query)
    packTypes = results['selectionList']     
    sessionData["packTypes"] = packTypes     


    #get Product Category list     
    col = db["metaTable"]
    query = {'category': "productCategory"}
    results = col.find_one(query)
    productCategories = results['selectionList']     
    sessionData["productCategories"] = productCategories


    #get Inspection Type 
    col = db["metaTable"]
    query = {'category': "inspType"}
    results = col.find_one(query)    
    sessionData["inspType"] = results['selectionList']        
        
    return sessionData