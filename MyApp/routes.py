from MyApp import app
from flask import render_template, redirect, url_for, flash, request, jsonify, send_file, session, Response
from MyApp.models import  sup_nominated_count,TMStatusByAdmin,UploadByAdmin_obj,sup_topic_count,AttendanceUploadByAdmin,UserDataEmail,UserDataNew,date_table,SupResultStatusByAdmin,TrainingRequest,IncidentTbl,training_topic_count,UploadByAdmin,training_topic_nomination,  SUP_topic, topic_question, Nominate_topic_by, calendar_events, Attendance, Upload, answerPerQnr, result_test, OJTNew,GDCList
from MyApp.forms import SubmitEvent,UploadFile,EnableMail,Training_Initiate,TrainingRequestForm,CloseInc,SendMail,RmAccess,getAccess,RegisterForm, LoginForm, PurchaseItemForm, SellItemForm, Feedback_form, DeleteForm, add_topic, add_question, getInfo, DeleteFormQNR, Update_add_question, TakeSup, EventForm_add, DeleteEvent, AttendanceForm, OJTForm
from MyApp import db
from flask_login import login_user, logout_user, login_required, current_user
import pdb
from sqlalchemy.orm import aliased
from sqlalchemy import select, func, case,intersect
from io import BytesIO
from datetime import date
import io,xlwt
import datetime
import win32com.client
from win32com.client import Dispatch, constants
from email.message import EmailMessage
import smtplib
import pythoncom
from os.path import join, dirname, realpath
import os
from sqlalchemy import create_engine, Column, Integer, String, ForeignKey 
from sqlalchemy.orm import declarative_base 
from sqlalchemy.orm import relationship 
from sqlalchemy.orm import sessionmaker 
import pandas as pd
from datetime import datetime
import random
import zipfile
import win32api
import win32security
import win32net
import re
import getpass

# import psutil
# import logging
# import commands
# logging.basicConfig(
#     level=logging.DEBUG,
#     format="{asctime} {levelname:<8} {message}",
#     style='{',
#     filename='%slog'% __file__[:-2],
#     filemode='a'
#     )
# logging.debug('Debug')
# logging.info('info')
# logging.warning('warning')
# logging.error('error')
# logging.critical('critical')

# Upload folder
UPLOAD_FOLDER = 'static/files'
app.config['UPLOAD_FOLDER'] =  UPLOAD_FOLDER

# Returns the current logged-in user or None if no user is logged in - USER
def get_current_user():
  return current_user

# Get user email - USER
def get_user_email (username):    
    try:        
        user_info = win32net.NetUserGetInfo(None, username, 2)
        email = user_info.get('email','No email found')        
        return email    
    except Exception as e:
            return str(e)

#Default page and this will add the new user into email table - USER
@app.route("/", methods=['GET', 'POST'])
@app.route('/UserEmail', methods=['GET', 'POST'])
def UserEmailCapture():
    user_email = request.args.get('user_email')
    user_name = request.args.get('user_name')
   
    
    if user_email is None or user_name is None:
        return render_template('IncorrectLogin.html')
 
    # Store user_data in session
    session['user_data_new'] = {
        'username': user_name,
        'email': user_email
    }
   
    user_data_new = UserDataEmail.query.filter(func.lower(UserDataEmail.email_add)==func.lower(user_email)).first()
   
    print("Session",session['user_data_new'])
   
    if user_data_new is None:
        print("Inside the None")
        user_to_create = UserDataEmail(
            username=user_name,
            email_add=user_email
        )
        db.session.add(user_to_create)
        db.session.commit()
       
    return redirect(url_for('adm'))

#Page to get the user basic details, only for the first time - USER 
@app.route('/SignIn/', methods=['GET', 'POST'])
def adm():
    
    # Retrieve user_data from session
    user_data_new = session.get('user_data_new')
    
    user_name = user_data_new.get('username')
    user_email = user_data_new.get('email')
    
    
    user_data_new = UserDataNew.query.filter(func.lower(UserDataNew.email_add)==func.lower(user_email)).first()  
    
    GDCSelect1 = request.args.get('GDCSelect')
    
    getGDC = GDCList.query.filter_by(GDCSelect=GDCSelect1).all()
    
    if user_data_new:
        login_user(user_data_new)  
        return redirect(url_for('admSide'))
    else:    
        formss = RegisterForm()
            
        if formss.validate_on_submit() or request.method=="POST":
            
            GDCSelect1 = request.args.get('GDCSelect')
            getGDC = GDCList.query.filter_by(GDCSelect=GDCSelect1).all()
                
            if "@kantar.com" in formss.MngID.data:
                print("success1")
            else:    
                flash("Please enter correct email ID.",category="danger")
                return render_template('deptpage.html', formss=formss,getGDC=getGDC,GDCSelect=GDCSelect1)
            
            if user_data_new:
                print("In-HHS")
            else:
                # User is new, create their profile
                
                GDCSelect1 = request.args.get('GDCSelect')
                getGDC = GDCList.query.filter_by(GDCSelect=GDCSelect1).all()
                
                user_to_create = UserDataNew(
                    username = user_name,
                    email_add = user_email,                                        
                    user_dept=formss.Dept.data,
                    GDCSelect=formss.GDCSelect.data,
                    Cluster=formss.Cluster.data,
                    MngID=formss.MngID.data,
                    user_type="user"
                )
                db.session.add(user_to_create)
                db.session.commit()
                
                login_user(user_to_create)
                flash('Account created successfully!', category='success')
                return redirect(url_for('admSide'))
        
        # Handle GET request or invalid form submission (GET request)
        return render_template('deptpage.html', formss=formss,getGDC=getGDC,GDCSelect=GDCSelect1)
    
    
 
#Page to display the calendar events filtered by department - USER/ADMIN
@app.route("/Dashboard", methods=['GET', 'POST'])
def admSide():
   
    user_data_new = session.get('user_data_new')
    user_email = user_data_new.get('email')
   
    NavActive = f'Home'        
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    #user_email = user_email
    user_data_new = UserDataNew.query.filter_by(email_add=user_email).first()  
    user_name = user_data_new.username
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    user_gdc = user_data_new.GDCSelect
    user_id = user_data_new.id
   
   
    if "admin" in user_type:
       calEvents_SP = calendar_events.query.filter_by(dept="SP").all()
       calEvents_DP = calendar_events.query.filter_by(dept="DP").all()
       calEvents_PM = calendar_events.query.filter_by(dept="PM").all()
       calEvents_CH = calendar_events.query.filter_by(dept="CH").all()
       calEvents_CO = calendar_events.query.filter_by(dept="CO").all()
    else:    
       calEvents_SP = calendar_events.query.filter_by(dept="SP").filter(calendar_events.NoOfGDC.like(f"%{user_gdc}%")).all()
       calEvents_DP = calendar_events.query.filter_by(dept="DP").filter(calendar_events.NoOfGDC.like(f"%{user_gdc}%")).all()
       calEvents_PM = calendar_events.query.filter_by(dept="PM").filter(calendar_events.NoOfGDC.like(f"%{user_gdc}%")).all()
       calEvents_CH = calendar_events.query.filter_by(dept="CH").filter(calendar_events.NoOfGDC.like(f"%{user_gdc}%")).all()
       calEvents_CO = calendar_events.query.filter_by(dept="CO").filter(calendar_events.NoOfGDC.like(f"%{user_gdc}%")).all()

    calEvents = training_topic_nomination.query.filter_by(year=year,period=period,NominatedBy=user_id).all()
 
    calEvents_SP_filtered = []
    calEvents_DP_filtered = []
    calEvents_PM_filtered = []
    calEvents_CH_filtered = []    
    calEvents_CO_filtered = []
 
    # Initialize lists to store filtered events
    calEvents_SP_filtered_Matched = []
    calEvents_DP_filtered_Matched = []
    calEvents_PM_filtered_Matched = []
    calEvents_CH_filtered_Matched = []
    calEvents_CO_filtered_Matched = []
    
   
    # Collect titles of events in calEvents
    matched_titles = [event.topics_added for event in calEvents]
 
    # Filter calEvents_SP
    for event in calEvents_SP:
        if event.title not in matched_titles:
            calEvents_SP_filtered.append(event)
        else:
            calEvents_SP_filtered_Matched.append(event)
   
   
    # Filter calEvents_DP
    for event in calEvents_DP:
        if event.title not in matched_titles:
            calEvents_DP_filtered.append(event)
        else:
            calEvents_DP_filtered_Matched.append(event)
            
    # Filter calEvents_PM
    for event in calEvents_PM:
         if event.title not in matched_titles:
             calEvents_PM_filtered.append(event)
         else:
             calEvents_PM_filtered_Matched.append(event)
    
    # Filter calEvents_CH
    for event in calEvents_CH:
         if event.title not in matched_titles:
             calEvents_CH_filtered.append(event)
         else:
             calEvents_CH_filtered_Matched.append(event)
    
    # Filter calEvents_CO
    for event in calEvents_CO:
         if event.title not in matched_titles:
             calEvents_CO_filtered.append(event)
         else:
             calEvents_CO_filtered_Matched.append(event)
             
    date_val = date_table.query.first()
   
    print("admin" in user_type)            
    if "admin" in user_type:
        return render_template('all_calendar.html', calEvents_CO_filtered_Matched=calEvents_CO_filtered_Matched,calEvents_CO_filtered=calEvents_CO_filtered,calEvents_CH_filtered_Matched=calEvents_CH_filtered_Matched,calEvents_CH_filtered=calEvents_CH_filtered,calEvents_PM_filtered_Matched=calEvents_PM_filtered_Matched,calEvents_PM_filtered=calEvents_PM_filtered,calEvents_DP_filtered_Matched=calEvents_DP_filtered_Matched,calEvents_SP_filtered_Matched=calEvents_SP_filtered_Matched,calEvents_DP_filtered=calEvents_DP_filtered,calEvents_SP_filtered=calEvents_SP_filtered,user=user_data_new, user_type=user_type,user_dept=user_dept,user_gdc=user_gdc,date_val=date_val,user_name=user_name)
    else:
        return render_template('all_calendar.html', calEvents_CO_filtered_Matched=calEvents_CO_filtered_Matched,calEvents_CO_filtered=calEvents_CO_filtered,calEvents_CH_filtered_Matched=calEvents_CH_filtered_Matched,calEvents_CH_filtered=calEvents_CH_filtered,calEvents_PM_filtered_Matched=calEvents_PM_filtered_Matched,calEvents_PM_filtered=calEvents_PM_filtered,calEvents_DP_filtered_Matched=calEvents_DP_filtered_Matched,calEvents_SP_filtered_Matched=calEvents_SP_filtered_Matched,calEvents_DP_filtered=calEvents_DP_filtered,calEvents_SP_filtered=calEvents_SP_filtered,user_gdc=user_gdc,user=user_data_new, user_type=user_type,user_dept=user_dept,date_val=date_val,user_name=user_name)
    
#Render the CapDev tools - USER
@app.route("/talentMeter", methods=['GET', 'POST']) 
def tmeter():
 return render_template('talentmeter.html')

#Report for SME - ADMIN
@app.route("/SMEReport", methods=['GET', 'POST']) 
def SMEReport():
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
   
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    cnt_items = calendar_events.query.all()
    if cnt_items==0:
        flash('There are no items in data',category='danger')
        return redirect(url_for('admSide'))
    
    dept = request.args.get('dept', 'none')
    page = request.args.get('page', 1, type=int)
     
    filters = []
 
    if dept == 'none':
      odata = calendar_events.query.paginate(page=page, per_page=10)
    else:
      filters.append(calendar_events.dept == dept)
      odata = calendar_events.query.filter(*filters).paginate(page=page, per_page=10)
  
    
    # Pass week_numbers and submitters to the template
    depts = calendar_events.query.with_entities(calendar_events.dept).distinct().all()
   
   # Get distinct week numbers and submitters
    depts = [item[0] for item in calendar_events.query.with_entities(calendar_events.dept).distinct().all()]
   
    return render_template('SMEReport.html',odata=odata,depts=depts,dept=dept,user_type=user_type)

#Report page for User Feedback on Training - ADMIN
@app.route("/TrainingReport", methods=['GET', 'POST']) 
def TrainingReport():
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    cnt_items = Attendance.query.filter_by(year=year,period=period).count()
    if cnt_items==0:
        flash('There are no feedbacks in data',category='danger')
        return redirect(url_for('admSide'))
    
    topic = request.args.get('topic', None)
    dept = request.args.get('dept', None)
    page = request.args.get('page', 1, type=int)
     
    filters = []
 
    if topic:
         filters.append(Attendance.Topicname == topic)
    if dept:
         filters.append(Attendance.user_dept == dept)
    odata = Attendance.query.filter(*filters).paginate(page=page, per_page=10)
    
    # Pass week_numbers and submitters to the template
    topics = Attendance.query.with_entities(Attendance.Topicname).distinct().all()
    depts = Attendance.query.with_entities(Attendance.user_dept).distinct().all()
   
   # Get distinct week numbers and submitters
    topics = [str(item[0]) for item in Attendance.query.with_entities(Attendance.Topicname).distinct().all()]
    depts = [item[0] for item in Attendance.query.with_entities(Attendance.user_dept).distinct().all()]
   
    return render_template('TrainingReport.html',odata=odata,topics=topics,depts=depts,user_type=user_type)

#report page for user nomination assessment  - ADMIN
@app.route("/NominationReport", methods=['GET', 'POST']) 
def NominationReport():
   
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    cnt_items = Nominate_topic_by.query.filter_by(year=year,period=period).count()
  
    if cnt_items==0:
        flash('There are no nomination in data',category='danger')
        return redirect(url_for('admSide'))
    q= request.args.get('q')
    if q:
        page = request.args.get('page', 1, type=int)
        odata = Nominate_topic_by.query.filter_by(topics_added=q,year=year,period=period).paginate(page=page, per_page=10)
    else:
        page = request.args.get('page', 1, type=int)
        odata = Nominate_topic_by.query.filter_by(year=year,period=period).paginate(page=page, per_page=10)
        
    user_id = ""
    username = get_username_by_id(user_id)
   
    get_useremail_by_id(user_id)
   
    return render_template('NominationReport.html',odata=odata,get_username_by_id=get_username_by_id,get_useremail_by_id=get_useremail_by_id)

#report page for user assessment nominated Vs completed count - ADMIN
@app.route("/NominationVsCompletedReport", methods=['GET', 'POST']) 
def NominationVsCompletedReport():
    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    cnt_items = sup_nominated_count.query.filter_by(year=year,period=period).count()
   
    if cnt_items==0:
        flash('There are no nomination in data',category='danger')
        return redirect(url_for('admSide'))
    
    dept = request.args.get('dept', None)
    page = request.args.get('page', 1, type=int)
    
    filters = []

    if dept:
        filters.append(sup_nominated_count.Dept == dept)
    odata = sup_nominated_count.query.filter(*filters).paginate(page=page, per_page=10)
   
   # Pass week_numbers and submitters to the template
    depts = sup_nominated_count.query.with_entities(sup_nominated_count.Dept).distinct().all()
  
  # Get distinct week numbers and submitters
    depts = [item[0] for item in sup_nominated_count.query.with_entities(sup_nominated_count.Dept).distinct().all()]
   
    user_id = ""
    username = get_username_by_id(user_id)
   
    get_useremail_by_id(user_id)
   
    return render_template('NominationVsCompletedReport.html',odata=odata,get_username_by_id=get_username_by_id,get_useremail_by_id=get_useremail_by_id,depts=depts)


#Report page for user training nomination count - ADMIN
@app.route("/NominationReportCount", methods=['GET', 'POST']) 
def NominationReportCount():
   
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    total_count = training_topic_count.query.filter_by(year=year,period=period).all()
   
    training_items = training_topic_count.query.filter_by(year=year,period=period).count()
    if training_items==0:
     flash('There are no data in this report',category='danger')
     return redirect(url_for('admSide'))
    
    dept = request.args.get('dept', 'none')
    page = request.args.get('page', 1, type=int)
  
    filters = []

    if dept == 'none':
        odata = training_topic_count.query.filter_by(year=year,period=period).paginate(page=page, per_page=10)
    else:
        filters.append(training_topic_count.Dept == dept)
        odata = training_topic_count.query.filter(*filters).paginate(page=page, per_page=10)
   
   # Pass week_numbers and submitters to the template
    depts = training_topic_count.query.with_entities(training_topic_count.Dept).distinct().all()
  
  # Get distinct week numbers and submitters
    depts = [item[0] for item in training_topic_count.query.with_entities(training_topic_count.Dept).distinct().all()]
    
    return render_template('NominationReportCount.html',dept=dept,odata=odata,depts=depts,user_type=user_type)

#report page to view the raised incidents  - USER
@app.route("/my_incident", methods=['GET', 'POST']) 
def my_incident():
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    page = request.args.get('page', 1, type=int)
    odata = IncidentTbl.query.filter_by(user_name=current_user.email_add,Status="Open").paginate(page=page, per_page=10)
    odata_len = IncidentTbl.query.filter_by(user_name=current_user.email_add,Status="Open").count()
    if odata_len==0:
        flash(f'You have no incidents to view',category='success')
        return redirect(url_for('admSide'))
    return render_template('MyIncidents.html',odata=odata,user_type=user_type)

#report page to review the assessment question and answer  - ADMIN
@app.route("/QuestionBank", methods=['GET', 'POST']) 
def QuestionBank():
 
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    cnt_items = topic_question.query.count()
    if cnt_items==0:
        flash('There are no questions in data',category='danger')
        return redirect(url_for('admSide'))
    q= request.args.get('q')
    if q:
        page = request.args.get('page', 1, type=int)
        odata = topic_question.query.filter(topic_question.q_topic.contains(q)).paginate(page=page, per_page=10)
    else:
        page = request.args.get('page', 1, type=int)
        odata = topic_question.query.paginate(page=page, per_page=10)
   
    return render_template('QuestionBank.html',odata=odata,user_type=user_type)

#report page to view the user answers for assessment - ADMIN
@app.route("/UserAnswerReport", methods=['GET', 'POST']) 
def UserAnswerReport():
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
  
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    cnt_items = answerPerQnr.query.count()
    if cnt_items==0:
        flash('There are no nomination in data',category='danger')
        return redirect(url_for('admSide'))
    q= request.args.get('q')
    if q:
        page = request.args.get('page', 1, type=int)
        odata = answerPerQnr.query.filter_by(session_id=q,year=year,period=period).paginate(page=page, per_page=10)
    else:
        page = request.args.get('page', 1, type=int)
        odata = answerPerQnr.query.filter_by(year=year,period=period).paginate(page=page, per_page=10)
   
    user_id = ""
    username = get_username_by_id(user_id)
    
    get_useremail_by_id(user_id)
    qn_id = ""
    get_qnname_by_id(qn_id)
   
    return render_template('UserAnswerReport.html',odata=odata,get_username_by_id=get_username_by_id,get_useremail_by_id=get_useremail_by_id,get_qnname_by_id=get_qnname_by_id,user_type=user_type)

#Report to view all the user raised incident - ADMIN
@app.route("/all_incident", methods=['GET', 'POST']) 
def all_incident():
   
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    get_info = getInfo()
    if get_info.validate_on_submit():
     inc_id = request.form.get('inc_id')
     return redirect(url_for('inc_details',inc_id=inc_id))
    formToUpdate = CloseInc()
    page = request.args.get('page', 1, type=int)
    odata = IncidentTbl.query.filter_by(Status="Open").paginate(page=page, per_page=10)
    odata_len = IncidentTbl.query.filter_by(Status="Open").count()
    if odata_len == 0:
        flash(f'you have no incidents to view',category='success')
        return redirect(url_for('admSide'))
    return render_template('AllIncidents.html',get_info=get_info,odata=odata,formToUpdate=formToUpdate,user_type=user_type)

# Report to take actions on user raised incident - ADMIN
@app.route("/all_incident_report", methods=['GET', 'POST']) 
def all_incident_report():
  
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    formToUpdate = CloseInc()
    mod = request.args.get('mod', None)
    page = request.args.get('page', 1, type=int)
    
    filters = []
    if mod:
       filters.append(IncidentTbl.modulename == mod)
    
    odata = IncidentTbl.query.filter(*filters).paginate(page=page, per_page=10)
   
    # Pass week_numbers and submitters to the template
    mods = IncidentTbl.query.with_entities(IncidentTbl.modulename).distinct().all()
  
  # Get distinct week numbers and submitters
    mods = [item[0] for item in IncidentTbl.query.with_entities(IncidentTbl.modulename).distinct().all()]
    
    odata_len = IncidentTbl.query.count()
    if odata_len == 0:
        flash(f'you have no incidents to view',category='success')
        return redirect(url_for('admSide'))
    get_info = getInfo()
    if get_info.validate_on_submit():
     inc_id = request.form.get('inc_id')
     return redirect(url_for('inc_details',inc_id=inc_id))
    return render_template('AllIncidents_Report.html',get_info=get_info,odata=odata,formToUpdate=formToUpdate,user_type=user_type,mods=mods)

#To close an incident - ADMIN
@app.route('/CloseIncident',methods=['GET','POST'])
def CloseIncident():
    formToUpdate = CloseInc()
    if request.method == "POST" and formToUpdate.validate_on_submit():
        insert = request.form.get('Getinc_ID') 
        quest = IncidentTbl.query.filter_by(inc_id=insert).first()
        quest.Status="Closed"
        quest.Closuredate = date.today()
        db.session.commit()
        flash(f'Incident closed successfully', category='success')
        return redirect(url_for('all_incident'))   

#To download the user uploaded file for an incident - ADMIN
@app.route('/downloadIncfile/<incid>',methods=['GET','POST'])
def downloadIncfile(incid):
    upload = IncidentTbl.query.filter_by(id=incid).first()  
    if upload.filename=='':
        flash(f'No files uploaded', category='success')
        return redirect(url_for('all_incident'))   
    return send_file(BytesIO(upload.data),download_name=upload.filename,as_attachment=True)

#Page to collect the feedback on training - USER
@app.route("/TrainingFeedback", methods=['GET', 'POST']) 
def attendance_page():
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    user_location = user_data_new.GDCSelect
    
    form = AttendanceForm()    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    topic_select = AttendanceUploadByAdmin.query.filter_by(Year=year, Period=period, Email=current_user.email_add,Feedback="Not Done").all()
 
    if form.validate_on_submit() or request.method == "POST":
        Topic_value = request.form['TopicSelect']
        data_exist = Attendance.query.filter_by(year=year,period=period,user_name=current_user.email_add,Topicname=Topic_value).first() 
        if data_exist:
            flash('Feedback has already been submitted for this topic.',category='danger')
            return redirect(url_for('attendance_page'))
        user_to_create = Attendance(user_name=current_user.email_add,
                              user_dept=Topic_value[-2:],
                              location=user_location,
                              name=form.name.data,
                              confirm_attendance=form.confirm_attendance.data,
                              Topicname = request.form['TopicSelect'],
                              level = Topic_value[-5:-3],
                              rate_1=form.rate_1.data,
                              rate_2=form.rate_2.data,
                              rate_3=form.rate_3.data,
                              rate_4=form.rate_4.data,
                              session_like=form.session_like.data,
                              session_better=form.session_better.data,
                              session_do_well=form.session_do_well.data,
                              session_even_better=form.session_even_better.data,
                              suggestion=form.suggestion.data,
                              FeedbackDate = date.today(),
                              year=year,
                              period=period)
        
        filterbynomiated = AttendanceUploadByAdmin.query.filter_by(Year=year,Period=period,Subject=request.form['TopicSelect'],Email=current_user.email_add).first()
        filterbynomiated.Feedback = "Done"
        
        db.session.add(user_to_create)
        db.session.commit()        
        flash('Attendance marked successfully!',category='success')
        return redirect(url_for('attendance_page'))
    # else:
    #     print(form.errors)
    if form.errors != {}:
        for err_msg in form.errors.values():
            flash(f'There was an error while filling attendance :{err_msg}',category='danger')
           
    return render_template('user_feedback.html',form=form,topic_select=topic_select,user_type=user_type)

#Page to send an email for the adhoc training request - USER
@app.route("/training_initiate", methods=['GET', 'POST']) 
def training_initiate():
     form = Training_Initiate()
     if request.method == "POST":
        To = request.form['To']
        Cc = request.form['Cc']
        Subject = request.form['Subject']
        Description = request.form['Description']
   
        const=win32com.client.constants
        pythoncom.CoInitialize()
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        mailItem = obj.CreateItem(olMailItem)
        mailItem.Subject = Subject
        mailItem.BodyFormat = 2
        mailItem.HTMLBody = Description
        mailItem.To = To
        mailItem.CC = Cc
        # mailItem.display()
        # mailItem.Send()
        flash(f'Mail sent successfully!',category='success')
        return redirect(url_for('admSide'))
    
     return render_template('training_initiate.html',form=form)    

#To raise an incident for any issues - USER
@app.route("/incident_page", methods=['GET', 'POST']) 
def incident_page():
  
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    user_gdc = user_data_new.GDCSelect

    if request.method == "POST":
        file = request.files['file'] 
        subject = request.form['Subject']
        issuetype = request.form['gdc']
        location = user_gdc
        modulename = request.form['md']
        description = request.form.get('Description')
        currentDay = datetime.now().day
        currentMonth = datetime.now().month
        if currentMonth<10:
            currentMonth = "0" + str(currentMonth)
        currentYear = datetime.now().year
        currentYear = currentYear % 100 
        number = random.randint(1000,9999)
        inc_id = "INC" + str(currentDay) + str(currentMonth) + str(currentYear) + str(number)

        
        incident_to_create = IncidentTbl(user_name=current_user.email_add,
                              user_dept=user_dept,
                              location=location,
                              issuetype = issuetype,
                              modulename=modulename,
                              Subject=subject,
                              Description=description,
                              inc_id = "I" + str(currentDay) + str(currentMonth) + str(currentYear) + str(number),
                              filename=file.filename,
                              Raiseddate = date.today(), 
                              data=file.read())
        db.session.add(incident_to_create)
        db.session.commit()        
        
        const=win32com.client.constants
        pythoncom.CoInitialize()
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        mailItem = obj.CreateItem(olMailItem)
        mailItem.Subject = str(inc_id) + ' - Incident raised for your query'
        mailItem.BodyFormat = 2
        mailItem.HTMLBody = '<HTML><BODY>Dear Team, <br><br> We have a query from one of the users, please find it below. <br/><br/>'+ str(description) + '</BODY></HTML>'
        
        if issuetype == "Technical":
            mailItem.To = 'deepak.sartape@kantar.com' 
            mailItem.CC = 'deepak.sartape@kantar.com'
        else:
            mailItem.To = 'harish.satyan@kantar.com' 
            mailItem.CC = 'harish.satyan@kantar.com'    
        
        # mailItem.display()
        # mailItem.Send()
        flash(f'Incident marked successfully!',category='success')
        return redirect(url_for('admSide'))
    
    return render_template('incident_page.html',user_type=user_type)       

#logout page - USER
@app.route('/logout')
def logout_page():
     logout_user()
     flash(f'You have been logged out', category='info')
     return render_template('logout.html')

#Login back from the logout page - USER
@app.route('/LoginBack',methods=['GET','POST'])
def LoginBack():
    iframe = "https://apps.powerapps.com/play/e/default-1e355c04-e0a4-42ed-8e2d-7351591f0ef1/a/8f0069e2-5d97-4b95-b079-4bab0f349392?tenantId=1e355c04-e0a4-42ed-8e2d-7351591f0ef1&hint=5be05465-9755-417b-ac5d-97a16515510c&sourcetime=1721805154263"
    return render_template('PowerAppsPage.html', iframe=iframe)

#Download the talent meter uploaded data from the backend - ADMIN
@app.route('/download/report/excel12', methods=['GET', 'POST'])
def download_tm_result(**kwargs):
       
        new_var_7 = TMStatusByAdmin.query.all()
        df = pd.DataFrame(columns=['ID','Function','Level','Topic','User Mail','Score','Remarks','Year','Period'])
        for i in new_var_7:
            df = df._append({'ID': i.id,
                    'Function': i.department,
                    'Level': i.level,
                    'Topic': i.topic,
                    'User Mail': i.mailid,
                    'Score': i.score,
                    'Remarks': i.remarks,
                    'Year': i.year,
                    'Period': i.period,
                    },ignore_index=True)
          # Save DataFrame to a BytesIO object
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
           df.to_excel(writer, index=False, sheet_name='Sheet1')
       
       # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=TalentMeter_Result_File.xlsx"})
        
        
#download the template for Talent Meter uploading process - ADMIN
@app.route('/download_TMtemplate', methods=['GET', 'POST'])
def download_TMtemplate(**kwargs):
        df = pd.DataFrame({'Function': ['SP', 'DP'],
                         'Level': ['L1', 'L1'],
                         'Topic': ['Priority_L1_SP','Priority_L1_DP'],
                         'User Email': ['abc.xyz@kantar.com','def.xyz@kantar.com'],
                         'Score': [90,98],
                         'Remarks': ['Pass', 'Fail'],
                         'Year': [2024, 2024],
                         'Period': ['H1', 'H1']
                         })

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
           df.to_excel(writer, index=False, sheet_name='Sheet1')
       # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=TMTemplate.xlsx"})
        
# Upload the data for talent meter - ADMIN
@app.route("/TMFileSub",methods=['GET','POST'])
def TMFileSub():
   
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
  
    if request.method == 'POST':
        file = request.files['file']       
     
        if file.filename != '':
           
           file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
          # set the file path
           file.save(file_path)
           col_names = ['Function','Level','Topic','User Email', 'Score','Remarks']
            # Use Pandas to parse the CSV file
           csvData = pd.read_excel(file)
            # Loop through the Rows
           for i,row in csvData.iterrows():
                    result_to_create = TMStatusByAdmin(
                                    mailid = row['User Email'],
                                  topic=row['Topic'],
                                  level=row['Level'],
                                  department=row['Function'],
                                  score=row['Score'],
                                  remarks=row['Remarks'],
                                  year=row['Year'],
                                  period=row['Period'],
                                  uploadby = get_current_user().email_add
                                 )
                    db.session.add(result_to_create)
                    db.session.commit()     
                   
                    
        flash(f'File uploaded successfully', category='success')
        return redirect(url_for('TMFileSub'))
    
    q= request.args.get('q')
    if q:
        page = request.args.get('page', 1, type=int)
        odata = TMStatusByAdmin.query.filter(TMStatusByAdmin.Topicname.contains(q)).paginate(page=page, per_page=8)
    else:
        page = request.args.get('page', 1, type=int)
        odata = TMStatusByAdmin.query.paginate(page=page, per_page=8)
        
    return render_template('TM_upload.html',odata=odata)


#Page to create assessment topic - ADMIN
@app.route('/adminNomination', methods=['GET', 'POST'])
def adm_nomination():
   form = add_topic()
 
   dept = request.args.get('dept', None)
   filters = []
   if dept:
       # filters.append(calendar_events.dept == dept,IsDeleted="No")
       filters.append((SUP_topic.Department == dept) & (SUP_topic.IsDeleted == "No"))
    # Pass week_numbers and submitters to the template
   SUP_topics = SUP_topic.query.filter(*filters).order_by(SUP_topic.topics_added.asc()).all()
   depts = SUP_topic.query.with_entities(calendar_events.dept).distinct().all()
 
 # Get distinct week numbers and submitters
   depts = [item[0] for item in SUP_topic.query.with_entities(SUP_topic.Department).distinct().all()]
   
   # SUP_topics = SUP_topic.query.filter_by(IsDeleted="No").order_by(SUP_topic.topic_ID.desc()).all()
   delete_form = DeleteForm()
   get_info = getInfo()
  
   deleteEvent = DeleteEvent()
   
   def validate_NoOfGDC(self, field):
        selected_values = field.data  # This will be a list of selected values
        
        # Check if 'All' is selected and if there are other selections
        if 'All' in selected_values and len(selected_values) > 1:
            flash('The option "All" cannot be combined with other GDC locations.', category='success')
            return redirect(url_for('adm_nomination'))
           
   if deleteEvent.validate_on_submit():
         deleteThisTopic = request.form.get('deleteThisTopic')
         print(f"Item to delete: {deleteThisTopic}")
         delete_object = SUP_topic.query.filter_by(topic_ID=deleteThisTopic).first()
         print(f"Query result: {delete_object}")
         if delete_object:
            delete_object.IsDeleted = "Yes" 
            topic_data = delete_object.topics_added
            db.session.commit()
            flash('Deleted successfully', category='success')
            return redirect(url_for('adm_nomination'))
   #add topic 
   if form.validate_on_submit():
       
       if 'All' in form.NoOfGDC.data:
            gdc_locations = ['GDC-India', 'GDC-Philippines', 'GDC-Colombia', 'GDC-Egypt', 'GDC-Bratislava', 'DRC-Poland', 'GRC-Czech']
            no_of_gdc_data = ','.join(gdc_locations)
       else:
            no_of_gdc_data = ','.join(form.NoOfGDC.data) 
       
       if 'SP' in form.Department.data:
           spsubtopic = form.SPSubTopic.data
       
       if 'SP' in form.Department.data:      
         topic_value = form.SPSubTopic.data + "_" + form.topic_added.data.strip() + "_" + form.Level.data + "_" + form.Department.data
       else:
         topic_value = form.topic_added.data.strip() + "_" + form.Level.data + "_" + form.Department.data
       topic_exist = SUP_topic.query.filter_by(topics_added=topic_value).first()
       if topic_exist:
           flash('This SUP Assessment topic is already exist in our system.', category='success')
           return redirect(url_for('adm_nomination'))
       if 'SP' in form.Department.data:       
           topic_to_create = SUP_topic(topics_added=topic_value,
                             question_type=form.question_types.data,
                             Department = form.Department.data,
                             Level = form.Level.data,
                             Mandatory = form.Mandatory.data,
                             TimeInMin = form.TimeInMin.data,
                             PassMark = form.PassMark.data,
                             NoOfGDC=no_of_gdc_data,
                             SPSubTopic = spsubtopic,
                             uploadby=current_user.email_add
                             )
       else:
           topic_to_create = SUP_topic(topics_added=topic_value,
                             question_type=form.question_types.data,
                             Department = form.Department.data,
                             Level = form.Level.data,
                             Mandatory = form.Mandatory.data,
                             TimeInMin = form.TimeInMin.data,
                             PassMark = form.PassMark.data,
                             NoOfGDC=no_of_gdc_data,
                             uploadby=current_user.email_add
                             )
      
       db.session.add(topic_to_create)
       db.session.commit()
       flash('SUP Assessment topic added successfully.', category='success')
       return redirect(url_for('adm_nomination'))
   
   
   user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()   
   user_type = user_data_new.user_type
   
# get info
   if get_info.validate_on_submit():
           # Get info
       getInfoTopic = request.form.get('info_topic')
       getTypeQ = request.form.get('info_typeq')
       getLevel = request.form.get('info_level')
       print(f"Item to get: {getInfoTopic}")
       return redirect(url_for('info_topic',user_type=user_type, getInfoTopic=getInfoTopic, getTypeQ=getTypeQ,getLevel=getLevel))
           
       
   
   if form.errors != {}: #no error
       for err_msg in form.errors.values():
          flash(f'Error in creating user: {err_msg}', category='danger')

   return render_template('admin_nomination.html', depts=depts,user_type=user_type,form=form, SUP_topics=SUP_topics, delete_form=delete_form, get_info=get_info,deleteEvent=deleteEvent)


#Add question per topic for SUP assessment - ADMIN
@app.route('/TopicInfo/<getInfoTopic>/<getTypeQ>/<getLevel>', methods=['POST', 'GET'])
def info_topic(**kwargs):
   add_q_form = add_question()
   topic_q = topic_question.query.all()
   getInfoTopic = kwargs.get('getInfoTopic')
   getTypeQ = kwargs.get('getTypeQ')
   getLevel = kwargs.get('getLevel')
   deleteEvent = DeleteFormQNR()
   topicINF = topic_question.query.filter_by(q_topic=getInfoTopic,IsDeleted="No").all()
   print(topicINF)
   topicINF_dept = SUP_topic.query.filter_by(topics_added=getInfoTopic).first()
   dept_name = topicINF_dept.Department
   
   user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()   
   user_type = user_data_new.user_type

   if add_q_form.validate_on_submit():
         if add_q_form.choices_nos.data=="2" and (add_q_form.q_choice3s.data!="" or add_q_form.q_choice4s.data!=""):
             flash(f'Please keep the option 3 and option 4 text boxes blank for dual choice question', category='danger')
             return render_template('admin_addQ.html', user_type=user_type,getLevel=getLevel,dept_name=dept_name,add_q_form=add_q_form, topic_q=topic_q, getInfoTopic=getInfoTopic, topicINF=topicINF, getTypeQ=getTypeQ, deleteEvent=deleteEvent)
         if add_q_form.choices_nos.data=="2" and (add_q_form.q_choice3s.data=="Option 3" or add_q_form.q_choice4s.data=="Option 4"):
             flash(f'You cannot chose option 3 or option 4 text boxes for dual choice question', category='danger')
             return render_template('admin_addQ.html', user_type=user_type,getLevel=getLevel,dept_name=dept_name,add_q_form=add_q_form, topic_q=topic_q, getInfoTopic=getInfoTopic, topicINF=topicINF, getTypeQ=getTypeQ, deleteEvent=deleteEvent) 
         #if add_q_form.choices_nos.data=="4" and (add_q_form.q_choice3s.data=="" or add_q_form.q_choice4s.data==""):
         if add_q_form.choices_nos.data=="2" and (add_q_form.q_choice1s.data==add_q_form.q_choice2s.data):
              flash(f'Please keep the option 1 and option 2 labels unique', category='danger')
              return render_template('admin_addQ.html', user_type=user_type,getLevel=getLevel,dept_name=dept_name,add_q_form=add_q_form, topic_q=topic_q, getInfoTopic=getInfoTopic, topicINF=topicINF, getTypeQ=getTypeQ, deleteEvent=deleteEvent)
         if add_q_form.choices_nos.data=="4" and (add_q_form.q_choice3s.data=="" or add_q_form.q_choice4s.data==""):
              flash(f'Please provide value for option 3 and option 4 text boxes', category='danger')
              return render_template('admin_addQ.html', user_type=user_type,getLevel=getLevel,dept_name=dept_name,add_q_form=add_q_form, topic_q=topic_q, getInfoTopic=getInfoTopic, topicINF=topicINF, getTypeQ=getTypeQ, deleteEvent=deleteEvent)
         if add_q_form.choices_nos.data=="4" and (add_q_form.q_choice1s.data==add_q_form.q_choice2s.data or add_q_form.q_choice1s.data==add_q_form.q_choice3s.data or add_q_form.q_choice1s.data==add_q_form.q_choice4s.data or add_q_form.q_choice2s.data==add_q_form.q_choice3s.data or add_q_form.q_choice2s.data==add_q_form.q_choice4s.data or add_q_form.q_choice3s.data==add_q_form.q_choice4s.data):
              flash(f'Please provide unique value for the options', category='danger')
              return render_template('admin_addQ.html', user_type=user_type,getLevel=getLevel,dept_name=dept_name,add_q_form=add_q_form, topic_q=topic_q, getInfoTopic=getInfoTopic, topicINF=topicINF, getTypeQ=getTypeQ, deleteEvent=deleteEvent)
         correctans = request.form.getlist('mycheckbox') 
         if len(correctans)==0:
             flash(f'Please select at least an answer to proceed', category='danger')
             return render_template('admin_addQ.html', user_type=user_type,getLevel=getLevel,dept_name=dept_name,add_q_form=add_q_form, topic_q=topic_q, getInfoTopic=getInfoTopic, topicINF=topicINF, getTypeQ=getTypeQ, deleteEvent=deleteEvent)
         if add_q_form.q_type.data=="Single" and len(correctans)>1:
              flash(f'Please select only one answer for single punch question', category='danger')
              return render_template('admin_addQ.html', user_type=user_type,getLevel=getLevel,dept_name=dept_name,add_q_form=add_q_form, topic_q=topic_q, getInfoTopic=getInfoTopic, topicINF=topicINF, getTypeQ=getTypeQ, deleteEvent=deleteEvent)
         if add_q_form.q_type.data=="Multi" and len(correctans)==1:
               flash(f'Please select multiple answer for multi punch question', category='danger')
               return render_template('admin_addQ.html', user_type=user_type,getLevel=getLevel,dept_name=dept_name,add_q_form=add_q_form, topic_q=topic_q, getInfoTopic=getInfoTopic, topicINF=topicINF, getTypeQ=getTypeQ, deleteEvent=deleteEvent)
         correctans = listToString(correctans) 
         
         q_to_create = topic_question(
                                      question_type=add_q_form.question_types.data,
                                      q_topic=add_q_form.q_topics.data,
                                      choices_no=add_q_form.choices_nos.data,
                                      q_type=add_q_form.q_type.data,
                                      q_text=add_q_form.q_texts.data,
                                      q_choice1=add_q_form.q_choice1s.data,
                                      q_choice2=add_q_form.q_choice2s.data,
                                      q_choice3=add_q_form.q_choice3s.data,
                                      q_choice4=add_q_form.q_choice4s.data,
                                      q_ans=correctans,
                                      q_points=add_q_form.q_pointss.data,
                                      level = getLevel,
                                      dept_Func=dept_name,
                                      uploadby=current_user.email_add)
         db.session.add(q_to_create)
         db.session.commit()
         flash(f'Successfully submitted', category='success')
         return redirect(url_for('info_topic',user_type=user_type,dept_name=dept_name,getInfoTopic=getInfoTopic, getTypeQ=getTypeQ,getLevel=getLevel))
         
   else:
       # return render_template('admin_addQ.html', add_q_form=add_q_form, topic_q=topic_q, getInfoTopic=getInfoTopic, topicINF=topicINF, getTypeQ=getTypeQ, deleteEvent=deleteEvent)
       print(add_q_form.errors)


   if request.method == "POST":
      # Delete topic
       deleteThisQnr = request.form.get('deleteThisQnr')      
       print(f"Item to delete: {deleteThisQnr}")
       delete_object = topic_question.query.filter_by(q_ID=deleteThisQnr).first()
       print(f"Query result: {delete_object}")
       if delete_object:
          delete_object.IsDeleted = "Yes"
          db.session.commit()
          flash(f'Deleted successfully', category='success')
          return redirect(url_for('info_topic',user_type=user_type,getLevel=getLevel,getInfoTopic=getInfoTopic, getTypeQ=getTypeQ))
       else:
          flash(f"Something went error ", category='danger')
       return redirect(url_for('info_topic',user_type=user_type,getLevel=getLevel,getInfoTopic=getInfoTopic, getTypeQ=getTypeQ))
  
   return render_template('admin_addQ.html', user_type=user_type,getLevel=getLevel,dept_name=dept_name,add_q_form=add_q_form, topic_q=topic_q, getInfoTopic=getInfoTopic, topicINF=topicINF, getTypeQ=getTypeQ, deleteEvent=deleteEvent)

#Add comments for the incident request raised from user - ADMIN
@app.route('/IncInfo/<inc_id>', methods=['POST', 'GET'])
def inc_details(inc_id):
    
   odata = IncidentTbl.query.filter_by(inc_id=inc_id).first()
   if request.method == "POST":
      # Delete topic
       AdminComments = request.form.get('Acomments')      
       UpdatedComment = IncidentTbl.query.filter_by(inc_id=inc_id).first()
       UpdatedComment.Admincomment = AdminComments
       db.session.commit()
       flash(f'Updated successfully', category='success')
       return redirect(url_for('all_incident_report'))
      
   return render_template('Incident_details.html', odata=odata,inc_id=inc_id)
  
#Page to download the template for bulk question creation - ADMIN
@app.route('/download_Qtemplate', methods=['GET', 'POST'])
def download_Qtemplate(**kwargs):
       
       df = pd.DataFrame({'question_type': ['Objective', 'Objective'],
                          'q_topic': ['Conjoint_L1_SP', 'Matrix_L1_SP'],
                          'choices_no': [4,4],
                          'q_text': ['What is Conjoint','What is Matrix'],
                          'type': ['Single','Multi'],
                          'q_choice1': ['A', 1],
                          'q_choice2': ['B', 2],
                          'q_choice3': ['C', 3],
                          'q_choice4': ['D', 4],
                          'q_ans': ['Option 2', 'Option 1,Option 2'],
                          'q_points': [5, 10],
                          'Function': ['SP', 'SP'],
                          'level': ['L1', 'L1']})
   
   # Save DataFrame to a BytesIO object
       output = BytesIO()
       with pd.ExcelWriter(output, engine='openpyxl') as writer:
           df.to_excel(writer, index=False, sheet_name='Sheet1')
       
       # Seek to the beginning of the BytesIO object
       output.seek(0)
       return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=QuestionCreationTemplate.xlsx"})

#Page to nominate for the training specifically for SP department - USER       
@app.route("/UserNominationSP/<SPSubTopic>", methods=['GET', 'POST']) 
def user_nominationSP(SPSubTopic):
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    GDCSelect = user_data_new.GDCSelect
    
    SUP_topics = SUP_topic.query.all()
    formToUpdate = Update_add_question()
    
    level_order = case(
        (SUP_topic.Level == 'L1', 1),
        (SUP_topic.Level == 'L2', 2),
        (SUP_topic.Level == 'L3', 3),
    )
   
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    todays_date = date.today() 
    nominate_val = date_val.start2
    nominate_Eval = date_val.end2
    IDlistE = nominate_Eval.split('-')
    N_Eyear = int(IDlistE[0])
    N_Emonth = int(IDlistE[1])
    N_Eday = int(IDlistE[2])
    IDlist = nominate_val.split('-')
    N_year = int(IDlist[0])
    N_month = int(IDlist[1])
    N_day = int(IDlist[2])
    day = todays_date.day
    month = todays_date.month
    year1 = todays_date.year
    d1 = date(year1,month,day)
    d2 = date(N_year,N_month,N_day)
    d3 = date(N_Eyear,N_Emonth,N_Eday)
    if d1 < d2:
        flash('The nominations are yet to open, please wait.', category='success')
        return redirect(url_for('admSide'))  
    if d1 > d3:
        flash('The nominations closed, please try for next period.', category='success')
        return redirect(url_for('admSide'))  


    SUP_topicsNominated = SUP_topic.query.filter(SUP_topic.SPSubTopic == SPSubTopic,func.instr(SUP_topic.NoOfGDC, GDCSelect) > 0).order_by(level_order).all()
    tmdata = TMStatusByAdmin.query.filter_by(mailid=get_current_user().email_add,remarks="Pass").all()
    topics_to_remove = {i.topic for i in tmdata}

    filtered_SUP_topicsNominated = [item for item in SUP_topicsNominated if item.topics_added not in topics_to_remove]
    
    topics_nominatedBY = Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id)
    CurrentUserId = user_id

    nominated_topics = Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id).all()
    
    date_obj = datetime.strptime(nominate_Eval, "%Y-%m-%d")
        
    # Format the datetime object as "16th June 2024"
        
    formatted_date = date_obj.strftime("%d-%B-%Y")
    
    if request.method == 'POST' and formToUpdate.validate_on_submit():
        
         insert = request.form.getlist('mycheckbox')
         Department = request.form.get('Department')
         # flash("request.form.get('Department')",request.form.get('Department'))
         if len(insert) == 0:
             flash('Please select the options to proceed.', category='danger')
             return render_template('user_nomination.html', nominate_Eval=formatted_date,nominated_topics=nominated_topics,SUP_topics=SUP_topics, formToUpdate=formToUpdate, SUP_topicsNominated=SUP_topicsNominated, topics_nominatedBY=topics_nominatedBY, CurrentUserId=CurrentUserId,Department="SP",user_type=user_type)
         
         for ele in insert:
             
             SUP_type = SUP_topic.query.filter_by(topics_added=ele).first()
             date_val = date_table.query.first()
             period = date_val.period
             year = date_val.year
             
             userNominated_to_create = Nominate_topic_by(topics_added=ele,
                             question_type=SUP_type.question_type,
                             Department = ele[-2:],
                             nominate_topic=formToUpdate.nominated_topic.data,
                             nominatedBY=user_id,
                             NominationDate = date.today(),
                             TimeInMin=SUP_type.TimeInMin,
                             Level = SUP_type.Level,
                             year=year,
                             period=period,
                             uploadby=current_user.email_add
                             )
        
             db.session.add(userNominated_to_create)
             db.session.commit()
             
             # Update topic counts
         test_topics = Nominate_topic_by.query.with_entities(Nominate_topic_by.topics_added, Nominate_topic_by.year, Nominate_topic_by.period,Nominate_topic_by.Department).distinct().all()
        
         for i in test_topics:
            nominated_count = Nominate_topic_by.query.filter_by(year=year, period=period, topics_added=i.topics_added).count()
            already_exist = sup_topic_count.query.filter_by(topic_name=i.topics_added, year=year, period=period).all()
            if already_exist:
                count_val = sup_topic_count.query.filter_by(topic_name=i.topics_added, year=year, period=period).first()
                count_val.Count = nominated_count
            else:
                topiccount_to_add = sup_topic_count(topic_name=i.topics_added, Count=nominated_count, Dept=i.Department, year=year, period=period)
                db.session.add(topiccount_to_add)
            db.session.commit()
           
    # Parse the date string into a datetime object
         flash('Successfully nominated.', category='success')
         return redirect(url_for('UserNomination_Before'))   
    
    return render_template('user_nomination.html', filtered_SUP_topicsNominated=filtered_SUP_topicsNominated,nominate_Eval=formatted_date,nominated_topics=nominated_topics,SUP_topics=SUP_topics, formToUpdate=formToUpdate, SUP_topicsNominated=SUP_topicsNominated, topics_nominatedBY=topics_nominatedBY, CurrentUserId=CurrentUserId,Department="SP",user_type=user_type)

#Page to nominate for the training except SP department - USER  
@app.route("/UserNomination/<Department>", methods=['GET', 'POST']) 
def user_nomination(Department):
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    GDCSelect = user_data_new.GDCSelect
    
    SUP_topics = SUP_topic.query.all()
    formToUpdate = Update_add_question()
    
    level_order = case(
        (SUP_topic.Level == 'L1', 1),
        (SUP_topic.Level == 'L2', 2),
        (SUP_topic.Level == 'L3', 3),
    )
  
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    todays_date = date.today() 
    nominate_val = date_val.start2
    nominate_Eval = date_val.end2
    IDlistE = nominate_Eval.split('-')
    N_Eyear = int(IDlistE[0])
    N_Emonth = int(IDlistE[1])
    N_Eday = int(IDlistE[2])
    IDlist = nominate_val.split('-')
    N_year = int(IDlist[0])
    N_month = int(IDlist[1])
    N_day = int(IDlist[2])
    day = todays_date.day
    month = todays_date.month
    year1 = todays_date.year
    d1 = date(year1,month,day)
    d2 = date(N_year,N_month,N_day)
    d3 = date(N_Eyear,N_Emonth,N_Eday)
    if d1 < d2:
        flash('The nominations are yet to open, please wait.', category='success')
        return redirect(url_for('admSide'))  
    if d1 > d3:
        flash('The nominations closed, please try for next period.', category='success')
        return redirect(url_for('admSide'))  

    SUP_topicsNominated = SUP_topic.query.filter(SUP_topic.Department == Department,func.instr(SUP_topic.NoOfGDC, GDCSelect) > 0).order_by(level_order).all()
   
    tmdata = TMStatusByAdmin.query.filter_by(mailid=get_current_user().email_add,remarks="Pass").all()
   
    topics_to_remove = {i.topic for i in tmdata}
   
    filtered_SUP_topicsNominated = [item for item in SUP_topicsNominated if item.topics_added not in topics_to_remove]
    
    topics_nominatedBY = Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id)
    CurrentUserId = user_id

    nominated_topics = Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id).all()
    
    date_obj = datetime.strptime(nominate_Eval, "%Y-%m-%d")
        
    # Format the datetime object as "16th June 2024"
        
    formatted_date = date_obj.strftime("%d-%B-%Y")
    
    if request.method == 'POST' and formToUpdate.validate_on_submit():
      
         insert = request.form.getlist('mycheckbox')
         Department = request.form.get('Department')
         if len(insert) == 0:
             flash('Please select the options to proceed.', category='danger')
             return render_template('user_nomination.html', nominate_Eval=formatted_date,nominated_topics=nominated_topics,SUP_topics=SUP_topics, formToUpdate=formToUpdate, SUP_topicsNominated=SUP_topicsNominated, topics_nominatedBY=topics_nominatedBY, CurrentUserId=CurrentUserId,Department=Department,user_type=user_type)
         
       
         for ele in insert:
             
             SUP_type = SUP_topic.query.filter_by(topics_added=ele).first()
            
             date_val = date_table.query.first()
             period = date_val.period
             year = date_val.year
             
             userNominated_to_create = Nominate_topic_by(topics_added=ele,
                             question_type=SUP_type.question_type,
                             Department = ele[-2:],
                             nominate_topic=formToUpdate.nominated_topic.data,
                             nominatedBY=user_id,
                             NominationDate = date.today(),
                             TimeInMin=SUP_type.TimeInMin,
                             Level = SUP_type.Level,
                             year=year,
                             period=period,
                             )
         
             db.session.add(userNominated_to_create)
             db.session.commit()
             
             # Update topic counts
         test_topics = Nominate_topic_by.query.with_entities(Nominate_topic_by.topics_added, Nominate_topic_by.year, Nominate_topic_by.period,Nominate_topic_by.Department).distinct().all()
         for i in test_topics:
            nominated_count = Nominate_topic_by.query.filter_by(year=year, period=period, topics_added=i.topics_added).count()
            already_exist = sup_topic_count.query.filter_by(topic_name=i.topics_added, year=year, period=period).all()
            if already_exist:
                count_val = sup_topic_count.query.filter_by(topic_name=i.topics_added, year=year, period=period).first()
                count_val.Count = nominated_count
            else:
                topiccount_to_add = sup_topic_count(topic_name=i.topics_added, Count=nominated_count, Dept=i.Department, year=year, period=period)
                db.session.add(topiccount_to_add)
            db.session.commit()
          
    # Parse the date string into a datetime object
         flash('Successfully nominated.', category='success')
         return redirect(url_for('UserNomination_Before'))   
    
    return render_template('user_nomination.html', filtered_SUP_topicsNominated=filtered_SUP_topicsNominated,nominate_Eval=formatted_date,nominated_topics=nominated_topics,SUP_topics=SUP_topics, formToUpdate=formToUpdate, SUP_topicsNominated=SUP_topicsNominated, topics_nominatedBY=topics_nominatedBY, CurrentUserId=CurrentUserId,Department=Department,user_type=user_type)

#Instruction Page for Objective assessment - USER  
@app.route("/ObjInstruction", methods=['GET', 'POST'])
def ObjInstruction():
    return render_template('objInstructiontext.html') 

#Page to list the assessment topics - USER
@app.route("/Subjects", methods=['GET', 'POST'])
def sup_subj():
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
     
    session['result']=""
    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    level_order = case(
        (Nominate_topic_by.Level == 'L1', 1),
        (Nominate_topic_by.Level == 'L2', 2),
        (Nominate_topic_by.Level == 'L3', 3),
    )
    
    topics_nominatedBY = Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id, question_type="obj").order_by(level_order).all()
    topic_names = UploadByAdmin_obj.query.all()
    nominated_topics = Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id, question_type="obj").count()
    if nominated_topics == 0:
        flash('You are yet to nominate for this period.', category='danger')
        return redirect(url_for('admSide')) 
   
    completed_topics = set(
        answer.topics_added for answer in Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id,SUP_status="Completed")
    )
    get_takeSup = TakeSup()
   
    if request.method == "POST":
        if get_takeSup.validate_on_submit():
            
            getInfoTopic = request.form.get('info_topic')
            
            
            return redirect(url_for('quiz', getInfoTopic=getInfoTopic))

    return render_template('user_obj_sup.html', topic_names=topic_names,topics_nominatedBY=topics_nominatedBY, get_takeSup=get_takeSup,completed_topics=completed_topics)                                                                               

#Download the input file for Objective assessment - USER
@app.route('/download_file_assess/<getInfoTopic>',methods=['GET','POST'])
def download_file_assess(getInfoTopic):
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
   
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    level_order = case(
        (Nominate_topic_by.Level == 'L1', 1),
        (Nominate_topic_by.Level == 'L2', 2),
        (Nominate_topic_by.Level == 'L3', 3),
    )
    completed_topics = set(
       answer.topics_added for answer in Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id,SUP_status="Completed")
    )
    get_takeSup = TakeSup()
    topics_nominatedBY = Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id, question_type="obj").order_by(level_order).all()
    upload = UploadByAdmin_obj.query.filter_by(topic_added=getInfoTopic).first()    
    num_results = UploadByAdmin_obj.query.filter_by(topic_added=getInfoTopic).count()  
  
    if num_results==0:
        flash('No files to download.', category='danger')
        return redirect(url_for('sup_subj'))
    else:
        return send_file(BytesIO(upload.data),download_name=upload.filename,as_attachment=True)
    

#Instruction page for Subjective Assessment - USER
@app.route("/subjInstruction", methods=['GET', 'POST'])
def subjInstruction():
    return render_template('subjInstructiontext.html')

#Listing the topics of Subjective assessment  - USER
@app.route("/subjective assessment", methods=['GET', 'POST'])
def subjective_upl():
 
 user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
 user_id = user_data_new.id
 user_dept = user_data_new.user_dept
 user_type = user_data_new.user_type   
 topic_names = UploadByAdmin.query.all()
 date_val = date_table.query.first()
 period = date_val.period
 year = date_val.year
 get_UploadFile = UploadFile()
 topics_nominatedBY = Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id,question_type="subj",SUP_status="Incomplete")
 cnt_topics_nominatedBY = Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id,question_type="subj").count()
 if cnt_topics_nominatedBY == 0:
     flash('You have no subjective assessment nominated for this period.', category='success')
     return redirect(url_for('admSide'))  
       
 return render_template('user_subj_sup.html', get_UploadFile=get_UploadFile,topic_names=topic_names,topics_nominatedBY=topics_nominatedBY,user_type=user_type)

#Listing the topics of Subjective assessment  - ADMIN
@app.route("/subjectivefiles", methods=['GET', 'POST'])
def subjectivefiles_download():
  
 user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
 user_id = user_data_new.id
 user_dept = user_data_new.user_dept
 user_type = user_data_new.user_type

 date_val = date_table.query.first()
 period = date_val.period
 year = date_val.year
 filesuploadebyadmin = UploadByAdmin.query.filter_by(year=year,period=period).all()
 if len(filesuploadebyadmin)==0:
    flash("There are no files to view.",category="danger")    
    return redirect(url_for('admSide'))  
 return render_template('admin_upl_subjfiles.html', filesuploadebyadmin=filesuploadebyadmin,user_type=user_type)

#Assessment page - USER
@app.route('/quiz/<getInfoTopic>',methods=['GET', 'POST'])
def quiz(getInfoTopic): 
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    session['current_question_index'] = 0
    
    current_user_account = user_id
    current_session_id = answerPerQnr.query.filter_by(session_id=getInfoTopic).first()
    
    subval = None # Initialize subval with a default value
    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    if current_session_id is not None:
        subval=current_session_id.session_id
        print("current_session_id",subval)
    
    existing_response = answerPerQnr.query.filter_by(
        useraccount=current_user_account,        
        session_id=subval,
        year=year,
        period=period
    ).all()
    
    if existing_response is not None:
        # Delete each record
        for response in existing_response:
            db.session.delete(response)
        # Commit the changes to the database
        db.session.commit()
    
    questList = topic_question.query.filter_by(q_topic=getInfoTopic).all()
    quest = topic_question.query.filter_by(q_topic=getInfoTopic).first()
    if len(questList)==0:
        flash("There are no questions added for this topic.",category="danger")
        return redirect(url_for('sup_subj'))  
    # Shuffle the list of questions
    if quest:
        # Retrieve the choices into a list
        choices = [quest.q_choice1, quest.q_choice2, quest.q_choice3, quest.q_choice4]
        #random.shuffle(choices)
        # Shuffle the choices
    
    random.shuffle(questList)

    # Shuffle the list of questions
    
    submitEvent=SubmitEvent()
    assignment = Nominate_topic_by.query.filter_by(year=year,period=period,topics_added=getInfoTopic).first()
 
    return render_template("user_qnr2.html",submitEvent=submitEvent,questList=questList, quest=choices,getInfoTopic=getInfoTopic,assignment=assignment,user_type=user_type) 

#Add answers to the database whe user submit the assessment - USER
@app.route('/qnr/<getInfoTopic>', methods=['GET', 'POST']) 
def sup_qnr(getInfoTopic):
    
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    if 'current_question_index' not in session:
        session['current_question_index'] = 0
      
    else:
        current_question_index = session.get('current_question_index', 0)
    
    submitEvent = SubmitEvent()
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    
    if submitEvent.validate_on_submit():
        submitThisEvent = request.form.get('submitThisEvent')
        reload = request.form.get('reload')
       
        if reload=='1':
            return redirect(url_for('quiz', getInfoTopic=submitThisEvent)) 
        starttime = request.form.get('starttime')
        stoptime = request.form.get('stoptime')
        start = request.form.get('start')
        stop = request.form.get('stop')
        final_val = int(stop) - int(start)
        final_val = round(final_val/60)
        submitobject = Nominate_topic_by.query.filter_by(year=year, period=period, topics_added=submitThisEvent, nominatedBY=user_id).first()
        if submitobject.SUP_status == "Completed":
            flash("This assessment is already completed",category="danger")
            return redirect(url_for('sup_subj'))  
        if submitobject:
            submitobject.SUP_status = "Completed"
            submitobject.AssessmentTime = final_val
            submitobject.StartTime = starttime
            submitobject.StopTime = stoptime
            db.session.commit()
            flash(f"Submitted Successfully",category='success')
            # return redirect(url_for('sup_qnr',getInfoTopic=submitobject))
            form_data = request.form.to_dict()
            
            answered_questions = False  # Flag to check if any questions are answered
            sub = None  # Initialize sub outside the loop
            
            for key, value in form_data.items():
                if key.startswith('answer_'):
                    qid = key.split('_')[1]
                    # ans = value
                    list_to_str = request.form.getlist(f'answer_{qid}')
                    list_to_str = listToString(list_to_str)
                    ans = list_to_str
                    sub = form_data.get(f'subject_{qid}')
                    level = form_data.get(f'level_{qid}')
                    correctAns = form_data.get(f'correctans_{qid}')
                    qpoints = form_data.get(f'qpoints_{qid}')
                    qtype = form_data.get(f'q_type_{qid}')
                    print("qtype - ",qtype)
                    
                    answer_to_add = answerPerQnr(
                        useraccount=user_id,
                        useranswer=ans,
                        correctanswer=correctAns,
                        question_id=qid,
                        session_id=sub,
                        q_points=qpoints,
                        level=level,
                        isitdone =  "done",
                        year = year,
                        period=period,
                        qtype = qtype
                    )
                    print(answer_to_add)
                    db.session.add(answer_to_add)
                    answered_questions = True
            
            if not answered_questions:
                # If no questions were answered, extract the subject from the form data
                for key in form_data.keys():
                    if key.startswith('subject_'):
                        sub = form_data.get(key)
                        break
            
            if sub:
                get_sub = Nominate_topic_by.query.filter_by(year=year, period=period, topics_added=sub, nominatedBY=user_id).first()
                if get_sub:
                    get_sub.SUP_status = "Completed"
                    get_cnt= Nominate_topic_by.query.filter_by(year=year, period=period, topics_added=sub,SUP_status = "Completed").count()
                    get_incompletecnt= Nominate_topic_by.query.filter_by(year=year, period=period, topics_added=sub,SUP_status = "Incomplete").count()
                    already_exist = sup_nominated_count.query.filter_by(topic_name=get_sub.topics_added, year=year, period=period).all()
                    if already_exist:
                        already_exist.Count = get_cnt
                        already_exist.IncompleteCount = get_incompletecnt
                    else:
                        topiccount_to_add = sup_nominated_count(topic_name=get_sub.topics_added, Count=get_cnt, IncompleteCount=get_incompletecnt, Dept=get_sub.Department, year=year, period=period)
                        db.session.add(topiccount_to_add)
                    db.session.commit()
                   
            else:
                # Handle the case where sub is still None
                print("Error: Subject not found")
                    
        
    # Fetch the current question if the index is within the valid range
    questList = topic_question.query.filter_by(q_topic=sub).all()
    if len(questList)==0:
        flash("There are no questions added for this topic",category="danger")
        return redirect(url_for('sup_subj'))  
   
    if current_question_index < len(questList):
        quest = questList[current_question_index]
        print(f"Processing question {current_question_index + 1}/{len(questList)}")
    else:
        print("Redirecting to completion page")
    # Reset the session variables and redirect to the completion page
        session.pop('current_question_index', None)
        return redirect(url_for('sup_subj'))  # Replace 'completion_page' with the actual endpoint for the completion page
 
    setStatus(questList)

    return redirect(url_for('sup_subj'))

#Convert list to String - USER/ADMIN
def listToString(s):
 
    # initialize an empty string
    str1 = ""
    i = 1
    # traverse in the string
    for ele in s:
        if i > 1:
            str1 = str1 + ','
        str1 += ele
        i+=1
    # return string
    return str1

#Change the color of the question number box when it is answered - USER
def setStatus(qlist):
    # print(qlist)
    qAttempt=[]
    strval=session['result'].strip()
    # print(strval)
    ans=strval.split(',')
    # print(ans)
    for i in range(int(len(ans)/2)):
        qAttempt.append(int(ans[2*i]))  
        # print(qAttempt)
    for rw in qlist:
        if rw.q_ID in qAttempt:
            rw.bcol='green'
            #rw.status='disabled' # disable

#Extract the list of questions for the user selected assessment - USER
@app.route("/showQuest/<string:getInfoTopic>,<int:qid>")
def showQuest(getInfoTopic,qid):
   
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    questList= topic_question.query.filter_by(q_topic=getInfoTopic).all()
    
    quest=topic_question.query.filter_by(q_ID=qid).first()
    setStatus(questList)
    
    assignment = Nominate_topic_by.query.filter_by(year=year,period=period,topics_added=getInfoTopic).first()
    
    # Assuming you have a way to get the current user's account and session ID
    current_user_account = get_current_user().id
    current_session_id = topic_question.query.filter_by(q_topic=getInfoTopic).first()
    
    existing_response = answerPerQnr.query.filter_by(
        useraccount=current_user_account,
        question_id=qid,
        year=year,
        period=period,
    ).first()
    submitEvent = SubmitEvent()
    user_answer = existing_response.useranswer if existing_response else None
    
    return render_template("user_qnr2.html",submitEvent=submitEvent,questList=questList, quest=quest,getInfoTopic=getInfoTopic,assignment=assignment,user_answer=user_answer)  

#Generate the results for the Objective assessment - USER
@app.route('/test_results')
def test_results():
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    cnt_res = result_test.query.filter_by(year=year,period=period,NominatedBy=user_id).count()
    todays_date = date.today() 
    
    date_val = date_table.query.first()
    nominate_val = date_val.startR
    IDlist = nominate_val.split('-')
    N_year = int(IDlist[0])
    N_month = int(IDlist[1])
    N_day = int(IDlist[2])
    day = todays_date.day
    month = todays_date.month
    year1 = todays_date.year
    d1 = date(year1,month,day)
    d2 = date(N_year,N_month,N_day)
    
    result_entry = result_test.query.filter_by(
        year=year,
        period=period,        
        NominatedBy=user_id
    ).first()

    
    if d1 < d2:
        flash('The results are yet to published, please wait.', category='success')
        return redirect(url_for('admSide'))
    
    if result_entry is None and d1 < d2:                    
        flash('Please take the assignment first then only result will publish.', category='success')
        return redirect(url_for('admSide')) 
      
    if d1 >= d2:
        
        getResult = answerPerQnr.query.filter_by(year=year,period=period,useraccount=user_id).all()
        q_points = []  # Initialize an empty list for q_points
      #  result_data = {}
    
        # Fetch results from answerPerQnr table
        results = answerPerQnr.query.with_entities(
            answerPerQnr.session_id,
            func.sum(
                case(
                    (answerPerQnr.useranswer == answerPerQnr.correctanswer, answerPerQnr.q_points),
                    else_=0
                )
            ).label('total_points')
        ).filter(
            answerPerQnr.useraccount == user_id,
            answerPerQnr.year == year,
            answerPerQnr.period == period
        ).group_by(answerPerQnr.session_id).all()
        
        # Process the results
        for session_id, total_points in results:
            # Fetch passing marks for the session
            passing_marks = SUP_topic.query.filter_by(topics_added=session_id).first()
            leveldata = answerPerQnr.query.filter_by(year=year,period=period,session_id=session_id).first()
            level=leveldata.level
            
            
            if passing_marks:
                passing_marks = passing_marks.PassMark
                
                # Check if total points are greater than passing marks
                if total_points >= passing_marks:
                    status = "Passed"
                else:
                    status = "Failed"
                
                result_entry = result_test.query.filter_by(
                    year=year,
                    period=period,        
                    NominatedBy=user_id,
                    session_id=session_id
                ).first()   
                
                
                if result_entry:
                    # Update existing entry
                    result_entry.overall_score = total_points
                    result_entry.remarks = status
                    result_entry.PassMark = passing_marks
                    result_entry.level = level
                else:
                    # Create new entry
                    result_entry = result_test(
                        session_id=session_id,
                        overall_score=total_points,
                        remarks=status,
                        NominatedBy=user_id,
                        year=year,
                        period=period,
                        PassMark=passing_marks,
                        level = level,
                        uploadby=current_user.email_add
                    )

                db.session.add(result_entry)
                db.session.commit()

        result = result_test.query.filter_by(year=year,period=period,NominatedBy=user_id).all()
        return render_template('userresult.html', result=result,q_points=q_points, getResult=getResult, result_data=result_entry,user_type=user_type)

#Page to add the trainin topic - ADMIN
@app.route('/addEventCalendar', methods=['GET', 'POST'])
def addEvent_calendar():
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type

    dept_order = case(
        (calendar_events.dept == 'SP', 1),
        (calendar_events.dept == 'DP', 2),
        (calendar_events.dept == 'CO', 3),
        (calendar_events.dept == 'PM', 4),
        (calendar_events.dept == 'CH', 5),
    )
    dept = request.args.get('dept', None)
    filters = []
    if dept:
        # filters.append(calendar_events.dept == dept,IsDeleted="No")
        filters.append((calendar_events.dept == dept) & (calendar_events.IsDeleted == "No"))
     # Pass week_numbers and submitters to the template
    calEvents = calendar_events.query.filter(*filters).order_by(calendar_events.title.asc()).all()
    depts = calendar_events.query.with_entities(calendar_events.dept).distinct().all()
  
  # Get distinct week numbers and submitters
    depts = [item[0] for item in calendar_events.query.with_entities(calendar_events.dept).distinct().all()]
    
    addEventForm = EventForm_add()  # Ensure this matches the form used in the template
    deleteEvent = DeleteEvent()

    if deleteEvent.validate_on_submit():
        deleteThisEvent = request.form.get('deleteThisEvent')
        delete_object = calendar_events.query.filter_by(eventid=deleteThisEvent).first()
        if delete_object:
            delete_object.IsDeleted = "Yes"
            db.session.commit()
            flash(f'Deleted successfully', category='success')
            return redirect(url_for('addEvent_calendar'))

    if request.method == "POST":
        title = request.form['title']
        dept = request.form['dept']
        start = request.form['start']
        Stime = request.form['Sappt']
        Etime = request.form['Eappt']
        url = request.form['url']
        level = request.form['level']
        sme = request.form['sme']
        if 'SP' in dept:
            SPSubTopic = request.form['spsubtopic']
        
        # Retrieve and process the NoOfGDC field
        NoOfGDC = request.form.getlist('topiccenter')  # Get the list of selected values
        if 'All' in NoOfGDC:
            # Define all possible centers
            all_centers = ["GDC-India", "GDC-Philippines", "GDC-Colombia", "GDC-Egypt", "GDC-Bratislava", "DRC-Poland", "GRC-Czech"]
            NoOfGDC = all_centers
        
        # Convert the list to a comma-separated string
        NoOfGDC_str = ','.join(NoOfGDC)

        startT = datetime.strptime(Stime, "%H:%M") 
        endT = datetime.strptime(Etime, "%H:%M") 
        difference = endT - startT 
        seconds = difference.total_seconds()
        if 'SP' in dept: 
            topic_value = SPSubTopic + "_" + title + "_" + level + "_" + dept
        else:
            topic_value = title + "_" + level + "_" + dept
                
        if seconds < 0:
            flash(f'Please provide the end time greater than the start time', category='success')
            return render_template('admin_addEvent.html', depts=depts,calEvents=calEvents, addEventForm=addEventForm, deleteEvent=deleteEvent, user_type=user_type)
        topic_exist = calendar_events.query.filter_by(title=topic_value).first()
        if topic_exist:
            flash(f'This topic already exist in our system', category='success')
            return redirect(url_for('addEvent_calendar'))
        if 'SP' in dept: 
            event_to_create = calendar_events(
                title=topic_value,
                dept=dept,
                date=start,
                Time=Stime,
                Etime=Etime,
                url=url,
                level=level,
                NoOfGDC=NoOfGDC_str,
                SPSubTopic = SPSubTopic,
                sme = sme,
                uploadby=current_user.email_add
            )
        else:
            event_to_create = calendar_events(
                title=topic_value,
                dept=dept,
                date=start,
                Time=Stime,
                Etime=Etime,
                url=url,
                level=level,
                NoOfGDC=NoOfGDC_str,
                sme = sme,
                uploadby=current_user.email_add
            )
        db.session.add(event_to_create)
        db.session.commit()
        flash(f'Successfully submitted', category='success')
        return redirect(url_for('addEvent_calendar'))

    return render_template('admin_addEvent.html', depts=depts,calEvents=calEvents, addEventForm=addEventForm, deleteEvent=deleteEvent, user_type=user_type)

#Download the template for adding the training topics - ADMIN
@app.route('/download_Etemplate', methods=['GET', 'POST'])
def download_Etemplate(**kwargs):
        df = pd.DataFrame({'title': ['Conjoint_L1_DP', 'NIPO_MaxDiff_L1_SP'],
                         'function': ['SP', 'DP'],
                         'level': ['L1','L1'],
                         'Time': ['18:00','18:00'],
                         'Etime': ['19:00','19:00'],
                         'url': ['', ''],
                         'date': ['6/30/2024', '6/30/2024'],
                         'gdc': ['GDC-India,GDC-Philippines,GDC-Colombia,GDC-Egypt', 'GDC-India'],
                         'spsubtopic': ['', 'NIPO']
                         })
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
         df.to_excel(writer, index=False, sheet_name='Sheet1')
     # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=CalendarTemplate.xlsx"})
        
#Download the template for adding the assessment topics - ADMIN       
@app.route('/download_SUPTemplate', methods=['GET', 'POST'])
def download_SUPTemplate(**kwargs):
        df = pd.DataFrame({'topics_added': ['TestTopicABC_L1_SP', 'NIPO_MaxDiff_L2_SP'],
                   'question_type': ['obj', 'obj'],
                   'Function': ['SP','SP'],
                   'Level': ['L1','L2'],
                   'Mandatory': ['Yes','No'],
                   'TimeInMin': [60, 90],
                   'PassMark': [80, 75],
                   'NoOfGDC': ['GDC-India,GDC-Philippines,GDC-Colombia,GDC-Egypt','GDC-India,GDC-Philippines,GDC-Colombia,GDC-Egypt,GDC-Bratislava,DRC-Poland,GRC-Czech'],
                   'SPSubTopic': ['', 'NIPO']
                   })
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
         df.to_excel(writer, index=False, sheet_name='Sheet1')
     # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=SUPTemplate.xlsx"})
      
        
#Update the training topic - ADMIN
@app.route('/UpdateEvent/<eventid>',methods=['GET','POST'])
def UpdateEvent(eventid):

    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type	
    
    event = calendar_events.query.filter_by(eventid=eventid).first()
    level = event.level
   
    if request.method == 'POST':
       if event:
           db.session.delete(event)
           db.session.commit()
           eventid = request.form['eventid']
           title = request.form['title']
           date = request.form['start']
           url = request.form['url']
           dept = request.form['dept']
           time = request.form['Sappt']
           Etime = request.form['Eappt']
           NoOfGDC = request.form['topiccenter']
           level = level
           event = calendar_events(level=level,eventid=eventid, title=title, date=date, url=url, dept = dept,Time=time,Etime=Etime,NoOfGDC=NoOfGDC,uploadby=current_user.email_add)
           db.session.add(event)
           db.session.commit()
           flash('Data updated successfully',category='success')
           return redirect(url_for('addEvent_calendar'))
    return render_template('admin_UpdateEvent.html',event=event,user_type=user_type)


#Upload the file for subjective assessment - ADMIN
@app.route("/UploadFile",methods=['GET','POST'])
def upload_file():
        
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_type = user_data_new.user_type    
    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    DeptSelect = request.args.get('DeptSelect')
    getSubj = SUP_topic.query.filter_by(question_type="subj",Department=DeptSelect)
    new_var = UploadByAdmin.query.filter_by(year=year,period=period)
    new_var2 = UploadByAdmin.query.filter_by(year=year,period=period).all()
    

    deleteEvent = DeleteEvent()

    if deleteEvent.validate_on_submit:
    # Delete topic
     deleteThisUpload = request.form.get('deleteThisUpload')      
    
    # Query the Upload model using db.session.query
    delete_object = db.session.query(UploadByAdmin).filter_by(id=deleteThisUpload).first()
    
    if delete_object:
        db.session.delete(delete_object)
        db.session.commit()
        flash(f'Deleted successfully', category='success')
        return redirect(url_for('upload_file'))
    else:
        print(f"Query result: {delete_object}")

    if request.method == 'POST':
        file = request.files['file']       
        topicSubj = request.form['topic_subj'] 
        
        parts = topicSubj.split('_')
        if len(parts) > 1:
            LevelSelect = parts[1]  # This will extract "L1"
        else:
            LevelSelect = "" 
        
        DeptSelect = request.form['DeptSelect'] 
        
        date_val = date_table.query.first()
        period = date_val.period
        year = date_val.year
        username1 = os.environ.get('USERNAME')
        fname = topicSubj + "_" + username1 + "_" + str(datetime.now()) + "_" + file.filename
        upload = UploadByAdmin(year=year,period=period,filename=fname,data=file.read(),dept=DeptSelect,topic_added=topicSubj,uploadby=user_id,level=LevelSelect)
        exist_file = db.session.query(UploadByAdmin).filter_by(year=year,period=period,topic_added=topicSubj,dept=DeptSelect,level=LevelSelect).first()
        if exist_file:
            flash(f'File already uploaded for this module', category='success')
            return redirect(url_for('upload_file'))
        else:
            db.session.add(upload)
            db.session.commit()     
            flash(f'File uploaded successfully', category='success')
            return redirect(url_for('upload_file'))
    
    return render_template('all_upload.html', user_type=user_type,DeptSelect=DeptSelect,new_var=new_var, getSubj=getSubj,deleteEvent=deleteEvent, new_var2=new_var2)

#Upload the file for Objective assessment - ADMIN
@app.route("/upload_file_Obj",methods=['GET','POST'])
def upload_file_Obj():
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    DeptSelect = request.args.get('DeptSelect')
    
    getSubj = SUP_topic.query.filter_by(question_type="obj",Department=DeptSelect)
    new_var = UploadByAdmin_obj.query.filter_by(year=year,period=period,uploadby=user_id)
    new_var2 = UploadByAdmin_obj.query.filter_by(year=year,period=period).all()
    
    deleteEvent = DeleteEvent()

    if deleteEvent.validate_on_submit:
    # Delete topic
     deleteThisUpload = request.form.get('deleteThisUpload')      
    
    # Query the Upload model using db.session.query
    delete_object = db.session.query(UploadByAdmin_obj).filter_by(id=deleteThisUpload).first()
    
    if delete_object:
        db.session.delete(delete_object)
        db.session.commit()
        flash(f'Deleted successfully', category='success')
        return redirect(url_for('upload_file_Obj'))
    else:
        print(f"Query result: {delete_object}")

    if request.method == 'POST':
        file = request.files['file']       
        topicSubj = request.form['topic_subj'] 
        DeptSelect = request.form['DeptSelect']
        LevelSelect = topicSubj[-2:]
       
        date_val = date_table.query.first()
        period = date_val.period
        year = date_val.year
        username1 = os.environ.get('USERNAME')
        fname = topicSubj + "_" + username1 + "_" + str(datetime.now()) + "_" + file.filename

        upload = UploadByAdmin_obj(year=year,period=period,filename=fname,data=file.read(),dept=DeptSelect,topic_added=topicSubj,uploadby=user_id,level=LevelSelect)
        exist_file = db.session.query(UploadByAdmin_obj).filter_by(year=year,period=period,topic_added=topicSubj,dept=DeptSelect,level=LevelSelect).first()
        if exist_file:
            flash(f'File already uploaded for this module', category='success')
            return redirect(url_for('upload_file_Obj'))
        else:
            db.session.add(upload)
            db.session.commit()     
            flash(f'File uploaded successfully', category='success')
            return redirect(url_for('upload_file_Obj'))
    
    return render_template('all_upload_Obj.html', DeptSelect=DeptSelect,new_var=new_var, getSubj=getSubj,deleteEvent=deleteEvent, new_var2=new_var2,user_type=user_type)

#Assessment Nominated User Count per topic - ADMIN
@app.route("/SUPReportCount", methods=['GET', 'POST']) 
def SUPReportCount():

    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    total_count = sup_topic_count.query.filter_by(year=year,period=period).all()
    if total_count==0:
        flash('There are no data in this report',category='danger')
        return redirect(url_for('admSide'))
    
  
    dept = request.args.get('dept', None)
    page = request.args.get('page', 1, type=int)
      
    filters = []

    if dept:
          filters.append(sup_topic_count.Dept == dept)
    odata = sup_topic_count.query.filter(*filters).paginate(page=page, per_page=10)
    
    depts = sup_topic_count.query.with_entities(sup_topic_count.Dept).distinct().all()
    
    depts = [item[0] for item in sup_topic_count.query.with_entities(sup_topic_count.Dept).distinct().all()]
    
    return render_template('SUPReportCount.html',odata=odata,depts=depts,user_type=user_type)

#Download the Assessment Nomination User Count report - ADMIN  
@app.route('/download/report/excel10', methods=['GET', 'POST'])
def download_SUPNominationCount_report(**kwargs):
        print(f'Form Data: {request.form}')
        new_var_4 = sup_topic_count.query.all()
       
        df = pd.DataFrame(columns=['ID','Topic Name','Function','Nominated Count'])
        for i in new_var_4:
            df = df._append({'ID': i.topic_ID,
                    'Topic Name': i.topic_name,
                    'Function': i.Dept,
                    'Nominated Count': i.Count,
                    },ignore_index=True)
          # Save DataFrame to a BytesIO object
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
   # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=SUPNominationCount_Report.xlsx"})
       

#Report for Adhoc training request page - ADMIN
@app.route("/AdhocRequest", methods=['GET', 'POST']) 
def AdhocRequest():
  
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    dept = request.args.get('dept', None)
    page = request.args.get('page', 1, type=int)
    filters = []


    if dept:
        filters.append(TrainingRequest.Department == dept)
    odata = TrainingRequest.query.filter(*filters).paginate(page=page, per_page=10)
    
    # Pass week_numbers and submitters to the template
    depts = TrainingRequest.query.with_entities(TrainingRequest.Department).distinct().all()
    
    # Get distinct week numbers and submitters
    depts = [item[0] for item in TrainingRequest.query.with_entities(TrainingRequest.Department).distinct().all()]
   
    odata_len = TrainingRequest.query.count()
    if odata_len == 0:
        flash(f'you have no report to view',category='success')
        return redirect(url_for('admSide'))
    return render_template('Adhoc_Training_Report.html',odata=odata,depts=depts)

#Download the Adhoc training report - ADMIN
@app.route('/download/report/excel11', methods=['GET', 'POST'])
def download_Adhoc_report(**kwargs):
        print(f'Form Data11: {request.form}')
        new_var_2 = TrainingRequest.query.all()
      
        df = pd.DataFrame(columns=['ID','Topic','Location','Function','cluster','Submitted By','People count'])
        for i in new_var_2:
            df = df._append({'ID': i.id,
                    'Topic': i.Topicname,
                    'Location': i.location,
                    'Function': i.Department,
                    'cluster': i.cluster,
                    'Submitted By': i.user_name,
                    'People count': i.PeopleCount,
                    },ignore_index=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
           df.to_excel(writer, index=False, sheet_name='Sheet1')
       # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=AdhocTraining_Report.xlsx"})
    
          
#Page to set the start and end date for Nomination, Assessment and Result announcement
@app.route('/setdate',methods=['GET','POST'])
def setdate():
 
   user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
   user_id = user_data_new.id
   user_dept = user_data_new.user_dept
   user_type = user_data_new.user_type
     
   data_items = date_table.query.all()
   if request.method=='POST':
       start1 = request.form['start1']
       end1 = request.form['end1']
       start2 = request.form['start2']
       end2 = request.form['end2']
       start3 = request.form['start3']
       end3 = request.form['end3']
       startR = request.form['startR']
       year = request.form['year']
       period = request.form['period']
       if end1 < start1 or end2 < start2 or end3 < start3:
           flash("Please input the end dates greater than the start date")
           return redirect(url_for('setdate')) 
       date_to_add = date_table(period=period,year=year,start1=start1,end1=end1,start2=start2,end2=end2,start3=start3,end3=end3,startR=startR)
       db.session.add(date_to_add)
       db.session.commit()
       flash("Date added successfully")
       return redirect(url_for('setdate')) 
   return render_template('admin_setdate.html',data_items=data_items,user_type=user_type)

#Page to modify the start and end date for Nomination, Assessment and Result announcement
@app.route('/UpdateDate/<Date_ID>',methods=['GET','POST'])
def UpdateDate(Date_ID):
   
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    event = date_table.query.filter_by(Date_ID=Date_ID).first()
    if request.method == 'POST':
       if event:
           db.session.delete(event)
           db.session.commit()
           Date_ID = request.form['Date_ID']
           start1 = request.form['start1']
           end1 = request.form['end1']
           start2 = request.form['start2']
           end2 = request.form['end2']
           start3 = request.form['start3']
           end3 = request.form['end3']
           startR = request.form['startR']
           year = request.form['year']
           period = request.form['period']
           event = date_table(period=period,year=year,Date_ID=Date_ID,start1=start1,end1=end1,start2=start2,end2=end2,start3=start3,end3=end3,startR=startR)
           db.session.add(event)
           db.session.commit()
           flash('Data updated successfully',category='success')
           return redirect(url_for('setdate'))
    return render_template('admin_Updatedate.html',event=event,user_type=user_type)

#Download the Attendance template - ADMIN
@app.route('/download_Attendance_template', methods=['GET', 'POST'])
def download_Attendance_template(**kwargs):
       
        df = pd.DataFrame({'Subject': ['Conjoint_L1_SP', 'Matrix_L1_SP'],
                  'Function': ['SP', 'SP'],
                  'Email': ['abc.xyz@kantar.com','def.xyz@kantar.com'],
                  'First_Join': ['4/17/24, 1:46:31 PM','4/17/24, 1:50:31 PM'],
                  'Last_Leave': ['4/17/24, 2:46:31 PM','4/17/24, 2:46:31 PM'],
                  'Duration': ['16m 29s', '19m 29s'],
                  'Period': ['H1', 'H1'],
                  'Year': ['2024', '2024']
                  })
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
         df.to_excel(writer, index=False, sheet_name='Sheet1')
     # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=AttendanceTemplate.xlsx"})
        
#Upload the user attendance - ADMIN
@app.route("/uploadAttendence", methods=['POST'])
def uploadAttendence():
      # get the uploaded file
      uploaded_file = request.files['file']
      form = add_question()
      upload_folder = app.config['UPLOAD_FOLDER']
      
      if uploaded_file.filename != '':
           file_path = os.path.join(upload_folder, uploaded_file.filename)
          # set the file path
           uploaded_file.save(file_path)
           col_names = ['Subject','Function', 'Email', 'First_Join' , 'Last_Leave', 'Duration', 'Feedback', 'Period', 'Year']
            # Use Pandas to parse the CSV file
           csvData = pd.read_excel(uploaded_file)
           
           duplicate_rows = csvData[csvData.duplicated(keep=False)]
           
           if not duplicate_rows.empty:
            flash(f'There are some duplicate values present in upload file, please recheck the data and upload the file again', category='danger')
            return redirect(url_for('uploadAttendencePage'))
           
           
           for i,row in csvData.iterrows():
                   
                   week_to_create = AttendanceUploadByAdmin(
                                         Subject = row['Subject'],
                                         Dept=row['Function'],
                                         Email=row['Email'],
                                         First_Join=row['First_Join'],
                                         Last_Leave=row['Last_Leave'],
                                         Duration=row['Duration'],
                                         Period=row['Period'],
                                         Year=row['Year'],
                                         uploadby = current_user.email_add
                                         )
                   db.session.add(week_to_create)
                   db.session.commit()     
           flash(f'Data uploaded successfully', category='success')
           return redirect(url_for('uploadAttendencePage'))
          # save the file
      else:
           flash(f'Please add file to proceed', category='danger')
           return redirect(url_for('uploadAttendencePage'))

# Page to Upload the Attendence - ADMIN
@app.route("/uploadAttendencePage", methods=['GET', 'POST'])
def uploadAttendencePage():
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
 
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    q= request.args.get('q')
    if q:
        page = request.args.get('page', 1, type=int)
        odata = AttendanceUploadByAdmin.query.filter_by(Period=period,Year=year).paginate(page=page, per_page=10)
    else:
        page = request.args.get('page', 1, type=int)
        odata = AttendanceUploadByAdmin.query.filter_by(Year=year,Period=period).paginate(page=page, per_page=10)
    return render_template('admin_attendence_upload.html',odata=odata,user_type=user_type)

#Page to view the list of subjective topics - USER
@app.route("/SubjectsSub", methods=['GET', 'POST'])
def sup_subjective():
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
     
    session['result']=""
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    level_order = case(
        (Nominate_topic_by.Level == 'L1', 1),
        (Nominate_topic_by.Level == 'L2', 2),
        (Nominate_topic_by.Level == 'L3', 3),
    )
    
    topics_nominatedBY = Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id, question_type="subj").order_by(level_order).all()
    topic_names = UploadByAdmin.query.all()
    nominated_topics = Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id, question_type="subj").count()
    if nominated_topics == 0:
        flash(f'You are yet to nominate for this period', category='danger')
        return redirect(url_for('admSide')) 
    completed_topics = set(
        answer.topics_added for answer in Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id,SUP_status="Completed")
    )
    
    get_UploadFile = UploadFile()
    

    if request.method == "POST":
        if get_UploadFile.validate_on_submit():
            getInfoTopic = request.form.get('info_topic')
            return redirect(url_for('USERupload_file', getInfoTopic=getInfoTopic))

    return render_template('user_subj_sup.html', topic_names=topic_names,topics_nominatedBY=topics_nominatedBY, get_UploadFile=get_UploadFile,completed_topics=completed_topics)                                                                               


#Page to upload the bulk training events - ADMIN
@app.route("/uploadEvents", methods=['POST'])
def uploadEvents():
      # get the uploaded file
      uploaded_file = request.files['file']
      form = add_question()
      upload_folder = app.config['UPLOAD_FOLDER']
     
      if uploaded_file.filename != '':
           file_path = os.path.join(upload_folder, uploaded_file.filename)
          # set the file path
           uploaded_file.save(file_path)
           col_names = ['title','function', 'level', 'Time' , 'Etime', 'url', 'date','gdc','spsubtopic']
            # Use Pandas to parse the CSV file
           # csvData = pd.read_csv(file_path,names=col_names, header=0)
           csvData = pd.read_excel(uploaded_file)
           
           for i,row in csvData.iterrows():
                   topic_exist = calendar_events.query.filter_by(title=row['title']).first()  
                   if topic_exist:
                       print("pass")
                   else:
                       date_str = row['date']
                        # Parse the date string into a datetime object
                       parsed_date = datetime.strptime(date_str, '%m/%d/%Y')
                        # Format the datetime object into the desired format
                       formatted_date = parsed_date.strftime('%Y-%m-%d')
                       time_format = '%H:%M'
                       F_Time = datetime.strptime(str(row['Time']), time_format) 
                       F_Time = F_Time.strftime('%H:%M')
                       F_Etime = datetime.strptime(str(row['Etime']), time_format) 
                       F_Etime = F_Etime.strftime('%H:%M')
                       week_to_create = calendar_events(
                                               title = row['title'],
                                             dept=row['function'],
                                             level=row['level'],
                                             Time=F_Time,
                                             Etime=F_Etime,
                                             url=row['url'],
                                             date=formatted_date,
                                             NoOfGDC=row['gdc'],
                                             SPSubTopic=row['spsubtopic'],
                                             uploadby=current_user.email_add)
                       db.session.add(week_to_create)
           db.session.commit()     
           flash(f'Data uploaded successfully', category='success')
           return redirect(url_for('addEvent_calendar'))
          # save the file
      else:
           flash(f'Please add file to proceed', category='danger')
           return redirect(url_for('addEvent_calendar'))

#Page to upload the bulk assessment events - ADMIN
@app.route("/uploadSuP", methods=['POST'])
def uploadSuP():
      # get the uploaded file
      uploaded_file = request.files['file']
      form = add_topic()
      upload_folder = app.config['UPLOAD_FOLDER']
     
      if uploaded_file.filename != '':
           file_path = os.path.join(upload_folder, uploaded_file.filename)
          # set the file path
           uploaded_file.save(file_path)
           col_names = ['topics_added','question_type', 'Function', 'Level' , 'Mandatory', 'TimeInMin', 'PassMark', 'NoOfGDC', 'SPSubTopic']
            # Use Pandas to parse the CSV file
           # csvData = pd.read_csv(file_path,names=col_names, header=0)
           csvData = pd.read_excel(uploaded_file)
           
           for i,row in csvData.iterrows():
                   
                   topic_exist = SUP_topic.query.filter_by(topics_added=row['topics_added']).first()  
                   if topic_exist:
                       print("pass")
                   else:
                       week_to_create = SUP_topic(
                                           topics_added = row['topics_added'],
                                         question_type=row['question_type'],
                                         Department=row['Function'],
                                         Level=row['Level'],
                                         Mandatory=row['Mandatory'],
                                         TimeInMin=row['TimeInMin'],
                                         PassMark=row['PassMark'],
                                         NoOfGDC=row['NoOfGDC'],
                                         SPSubTopic=row['SPSubTopic'],
                                         uploadby=current_user.email_add
                                         )
                       db.session.add(week_to_create)
           db.session.commit()     
           flash(f'Data uploaded successfully', category='success')
           return redirect(url_for('adm_nomination'))
          # save the file
      else:
           flash(f'Please add file to proceed', category='danger')
           return redirect(url_for('adm_nomination'))

#download the subjective file uploaded by admin - USER
@app.route('/download/<upload_id>',methods=['GET','POST'])
def download_file(upload_id):
    upload = Upload.query.filter_by(topic_added=upload_id).first()    
    return send_file(BytesIO(upload.data),download_name=upload.filename,as_attachment=True)

#download the objective file uploaded by admin - USER
@app.route('/downloadfile/<upload_id>',methods=['GET','POST'])
def download_file_addedbyadmin(upload_id):
    upload = UploadByAdmin.query.filter_by(id=upload_id).first()    
    return send_file(BytesIO(upload.data),download_name=upload.filename,as_attachment=True)

#download the subjective file uploaded by admin - USER
@app.route('/download_file_sub_assess/<getInfoTopic>',methods=['GET','POST'])
def download_file_sub_assess(getInfoTopic):
   
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    level_order = case(
        (Nominate_topic_by.Level == 'L1', 1),
        (Nominate_topic_by.Level == 'L2', 2),
        (Nominate_topic_by.Level == 'L3', 3),
    )
    completed_topics = set(
       answer.topics_added for answer in Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id,SUP_status="Completed")
    )
    get_takeSup = TakeSup()
    topics_nominatedBY = Nominate_topic_by.query.filter_by(year=year,period=period,nominatedBY=user_id, question_type="subj").order_by(level_order).all()
    upload = UploadByAdmin.query.filter_by(topic_added=getInfoTopic).first()    
    num_results = UploadByAdmin.query.filter_by(topic_added=getInfoTopic).count()  
   
    if num_results==0:
        flash(f'No files to download', category='danger')
        return redirect(url_for('sup_subjective'))
    else:
        return send_file(BytesIO(upload.data),download_name=upload.filename,as_attachment=True)

#Upload file for subjective assessment - USER
@app.route("/User upload",methods=['GET','POST'])
def USERupload_file():
        
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
     
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    getSubj = Nominate_topic_by.query.filter_by(year=year,period=period,question_type="subj",nominate_topic="Nominated",nominatedBY=user_id)
   
    new_var = Upload.query.filter_by(year=year,period=period,uploadby=user_id)    
    new_var2 = Upload.query.filter_by(year=year,period=period).all()
    
    deleteEvent = DeleteEvent()

    if deleteEvent.validate_on_submit:
    # Delete topic
     deleteThisUpload = request.form.get('deleteThisUpload')      
    print(f"Item to delete: {deleteThisUpload}")
    
    # Query the Upload model using db.session.query
    delete_object = db.session.query(Upload).filter_by(id=deleteThisUpload).first()
    
    print(f"Query result: {delete_object}")
    
    if delete_object:
        db.session.delete(delete_object)
        db.session.commit()
        flash('Deleted successfully.', category='success')
        return redirect(url_for('USERupload_file'))
    else:
        print(f"Query result: {delete_object}")
   
    if request.method == 'POST':
        
        file = request.files['file']       
        topicSubj = request.form['topic_subj'] 
        
        leveldata = Nominate_topic_by.query.filter_by(year=year,period=period,topics_added=topicSubj,nominatedBY=user_id).first()
        level = leveldata.Level
        file_check = db.session.query(Upload).filter_by(year=year,period=period,topic_added=topicSubj,level=level,uploadby=user_id).first()
        username1 = os.environ.get('USERNAME')
        if file_check:
            flash('You have already uploaded for this topic.', category='danger')
            return redirect(url_for('USERupload_file'))
        fname = topicSubj + "_" + username1 + "_" + str(datetime.now()) + "_" + file.filename 
        upload = Upload(year=leveldata.year,period=leveldata.period,level=level,filename=fname,data=file.read(),dept=user_dept,topic_added=topicSubj,uploadby=user_id)
        db.session.add(upload)
        flash('Uploaded successfully.', category='success')
        filterbynomiated = Nominate_topic_by.query.filter_by(year=year,period=period,Level=level,topics_added=upload.topic_added,nominatedBY=user_id,question_type="subj").first()
        filterbynomiated.SUP_status="Completed"
        # print("checking")
        
        db.session.commit()  

        
        const=win32com.client.constants
        pythoncom.CoInitialize()
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        mailItem = obj.CreateItem(olMailItem)
        mailItem.Subject = 'Subjective Assessment Uploaded - ' + str(topicSubj) 
        mailItem.BodyFormat = 2
        mailItem.HTMLBody = '<HTML><BODY>Dear Team, <br><br> Your team member has successfully uploaded a file for the subjective assessment.</BODY></HTML>'
        # mailItem.start = "02/01/2024 07:00:00 PM"  
        # mailItem.duration = 30
        # mailItem.importance = 2
        # mailItem.meetingstatus = 1
        # required = mailItem.Recipients.add("harish.satyan@kantar.com")
        # required.Type = 1
        # optional = mailItem.Recipients.add("harish.satyan@kantar.com")
        # optional.Type = 2
        # mailItem.respond(Response(True))
        
        mailItem.To = 'harish.satyan@kantar.com'
        mailItem.CC = 'harish.satyan@kantar.com'
        # mailItem.display()
        # mailItem.Send()
    
        return redirect(url_for('USERupload_file',const=const))
    
    return render_template('user_upload.html', user_type=user_type,new_var=new_var, getSubj=getSubj, new_var2=new_var2, deleteEvent=deleteEvent)


#View the subjective upload files by user - ADMIN
@app.route("/View upload",methods=['GET','POST'])
def viewUpload_file():
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    ViewUpload = Upload.query.filter_by(year=year,period=period).all()
    ViewUpload_count = Upload.query.filter_by(year=year,period=period).count()
    if ViewUpload_count==0:
        flash("There are no files to view for this period")
        return redirect(url_for('admSide'))
    
    subject_data = {}
    for item in ViewUpload:
        subject_name = item.topic_added
        if subject_name not in subject_data:
            subject_data[subject_name] = []
        subject_data[subject_name].append({
            'upload': item,
            'user_data': UserDataNew.query.filter_by(id=item.uploadby).first()
        })
    
    return render_template('admin_uplView.html', ViewUpload=ViewUpload, subject_data=subject_data,user_type=user_type)


#Download the admin uploaded files - USER
@app.route('/downloadAdm/<upload_id>',methods=['GET','POST'])
def download_fileAdm(upload_id):
    upload = Upload.query.filter_by(id=upload_id).first()    
    return send_file(BytesIO(upload.data),download_name=upload.filename,as_attachment=True)

#View the user feedback  - ADMIN
@app.route("/View feedback",methods=['GET','POST'])
def viewFeedback():
    feedbackView = Attendance.query.all()
    return render_template('admin_viewFeedback.html', feedbackView=feedbackView)


#Nominate training Only for SP - USER 
@app.route('/UserTrainingNominationSP/<SPSubTopic>', methods=['GET', 'POST']) 
def user_trainingSP(SPSubTopic):
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type    
    user_id = user_data_new.id
    GDCSelect = user_data_new.GDCSelect
  
    # Fetch today's date and relevant date information
    todays_date = date.today()
    date_val = date_table.query.first()
    nominate_val = date_val.start1
    nominate_Eval = date_val.end1

    # Parse nomination dates
    IDlistE = nominate_Eval.split('-')
    N_Eyear = int(IDlistE[0])
    N_Emonth = int(IDlistE[1])
    N_Eday = int(IDlistE[2])

    IDlist = nominate_val.split('-')
    N_year = int(IDlist[0])
    N_month = int(IDlist[1])
    N_day = int(IDlist[2])

    d1 = date(todays_date.year, todays_date.month, todays_date.day)
    d2 = date(N_year, N_month, N_day)
    d3 = date(N_Eyear, N_Emonth, N_Eday)
    print("d1 - ",d1)
    print("d2 - ",d2)
    # Handle nomination date checks
    if d1 < d2:
        flash('The nominations are yet to open, please wait.', category='danger')
        return redirect(url_for('admSide'))

    if d1 > d3:
        flash('The nominations closed, please try for next period.', category='danger')
        return redirect(url_for('admSide'))

    # Get period and year from date_val
    period = date_val.period
    year = date_val.year

    # Initialize formToUpdate and CurrentUserId
    formToUpdate = Update_add_question()
    CurrentUserId = user_id

    # Fetch nominated topics for the current user
    nominated_topics = training_topic_nomination.query.filter_by(year=year, period=period, NominatedBy=user_id).all()

    # Filter calendar_events based on dept
    training_topics = calendar_events.query.filter(calendar_events.SPSubTopic == SPSubTopic,func.instr(calendar_events.NoOfGDC, GDCSelect) > 0).order_by(calendar_events.level).distinct().all()
    
    # Get distinct departments and levels for dropdowns or other purposes
    depts = [str(item[0]) for item in calendar_events.query.with_entities(calendar_events.SPSubTopic).distinct().order_by(calendar_events.SPSubTopic).all()]
    levels = [str(item[0]) for item in calendar_events.query.with_entities(calendar_events.level).distinct().order_by(calendar_events.level).all()]

    if request.method == 'POST':
        # Handle form submission logic
        insert = request.form.getlist('mycheckbox')

        if len(insert) == 0:
            flash('Please select the options to proceed.', category='danger')
            return render_template('user_trainingNomination.html', CurrentUserId=CurrentUserId, training_topics=training_topics, formToUpdate=formToUpdate, nominated_topics=nominated_topics, SPSubTopic=SPSubTopic)

        for ele in insert:
            # Create and commit userTraining_to_create object
            userTraining_to_create = training_topic_nomination(
                topics_added=ele,
                nominate_topic=formToUpdate.nominated_topic.data,
                Dept="SP",
                NominatedBy=user_id,
                year=year,
                period=period,
                uploadby=current_user.email_add
            )

            db.session.add(userTraining_to_create)
            db.session.commit()

        # Update topic counts
        test_topics = training_topic_nomination.query.with_entities(training_topic_nomination.topics_added, training_topic_nomination.year, training_topic_nomination.period, training_topic_nomination.Dept).distinct().all()
        # for i in test_topics:
            
        for i in test_topics:
            nominated_count = training_topic_nomination.query.filter_by(year=year, period=period, topics_added=i.topics_added).count()
            already_exist = training_topic_count.query.filter_by(topic_name=i.topics_added, year=year, period=period).all()

            if already_exist:
                count_val = training_topic_count.query.filter_by(topic_name=i.topics_added, year=year, period=period).first()
                count_val.Count = nominated_count
            else:
                topiccount_to_add = training_topic_count(topic_name=i.topics_added, Count=nominated_count, Dept=i.Dept, year=year, period=period)
                db.session.add(topiccount_to_add)

            db.session.commit()

        flash('Successfully nominated.', category='success')
        return redirect(url_for('TrainingNomination_Before'))

    # Render template with formatted nomination date
    date_obj = datetime.strptime(nominate_Eval, "%Y-%m-%d")
    formatted_date = date_obj.strftime("%d-%B-%Y")

    return render_template('user_trainingNomination.html', nominate_Eval=formatted_date, CurrentUserId=CurrentUserId, training_topics=training_topics, formToUpdate=formToUpdate, nominated_topics=nominated_topics, depts=depts, selected_checkboxes=request.args.getlist('mycheckbox', type=str), SPSubTopic=SPSubTopic,user_type=user_type)

#Nominate training except SP department - USER 
@app.route("/UserTrainingNomination/<dept>", methods=['GET', 'POST']) 
def user_training(dept):
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type    
    user_id = user_data_new.id
    GDCSelect = user_data_new.GDCSelect
    
    
    # Fetch today's date and relevant date information
    todays_date = date.today()
    date_val = date_table.query.first()
    nominate_val = date_val.start1
    nominate_Eval = date_val.end1

    # Parse nomination dates
    IDlistE = nominate_Eval.split('-')
    N_Eyear = int(IDlistE[0])
    N_Emonth = int(IDlistE[1])
    N_Eday = int(IDlistE[2])

    IDlist = nominate_val.split('-')
    N_year = int(IDlist[0])
    N_month = int(IDlist[1])
    N_day = int(IDlist[2])

    d1 = date(todays_date.year, todays_date.month, todays_date.day)
    d2 = date(N_year, N_month, N_day)
    d3 = date(N_Eyear, N_Emonth, N_Eday)

    # Handle nomination date checks
    if d1 < d2:
        flash('The nominations are yet to open, please wait.', category='danger')
        return redirect(url_for('admSide'))

    if d1 > d3:
        flash('The nominations closed, please try for next period.', category='danger')
        return redirect(url_for('admSide'))

    # Get period and year from date_val
    period = date_val.period
    year = date_val.year

    # Initialize formToUpdate and CurrentUserId
    formToUpdate = Update_add_question()
    CurrentUserId = user_id

    # Fetch nominated topics for the current user
    nominated_topics = training_topic_nomination.query.filter_by(year=year, period=period, NominatedBy=user_id).all()

    # Filter calendar_events based on dept
    training_topics = calendar_events.query.filter(calendar_events.dept == dept,func.instr(calendar_events.NoOfGDC, GDCSelect) > 0).order_by(calendar_events.level).distinct().all()

    # Get distinct departments and levels for dropdowns or other purposes
    depts = [str(item[0]) for item in calendar_events.query.with_entities(calendar_events.dept).distinct().order_by(calendar_events.dept).all()]
    levels = [str(item[0]) for item in calendar_events.query.with_entities(calendar_events.level).distinct().order_by(calendar_events.level).all()]

    if request.method == 'POST':
        # Handle form submission logic
        insert = request.form.getlist('mycheckbox')

        if len(insert) == 0:
            flash('Please select the options to proceed.', category='danger')
            return render_template('user_trainingNomination.html', CurrentUserId=CurrentUserId, training_topics=training_topics, formToUpdate=formToUpdate, nominated_topics=nominated_topics, dept=dept)

        for ele in insert:
            # Create and commit userTraining_to_create object
            userTraining_to_create = training_topic_nomination(
                topics_added=ele,
                nominate_topic=formToUpdate.nominated_topic.data,
                Dept=dept,
                NominatedBy=user_id,
                year=year,
                period=period,
                uploadby=current_user.email_add
            )

            db.session.add(userTraining_to_create)
            db.session.commit()

        # Update topic counts
        test_topics = training_topic_nomination.query.with_entities(training_topic_nomination.topics_added, training_topic_nomination.year, training_topic_nomination.period, training_topic_nomination.Dept).distinct().all()

        for i in test_topics:
            nominated_count = training_topic_nomination.query.filter_by(year=year, period=period, topics_added=i.topics_added).count()
            already_exist = training_topic_count.query.filter_by(topic_name=i.topics_added, year=year, period=period).all()

            if already_exist:
                count_val = training_topic_count.query.filter_by(topic_name=i.topics_added, year=year, period=period).first()
                count_val.Count = nominated_count
            else:
                topiccount_to_add = training_topic_count(topic_name=i.topics_added, Count=nominated_count, Dept=i.Dept, year=year, period=period)
                db.session.add(topiccount_to_add)

            db.session.commit()

        flash('Successfully nominated.', category='success')
        return redirect(url_for('TrainingNomination_Before'))

    # Render template with formatted nomination date
    date_obj = datetime.strptime(nominate_Eval, "%Y-%m-%d")
    formatted_date = date_obj.strftime("%d-%B-%Y")

    return render_template('user_trainingNomination.html', nominate_Eval=formatted_date, CurrentUserId=CurrentUserId, training_topics=training_topics, formToUpdate=formToUpdate, nominated_topics=nominated_topics, depts=depts, selected_checkboxes=request.args.getlist('mycheckbox', type=str), dept=dept,user_type=user_type)

#Download calendar event report - ADMIN
@app.route('/download/report/excel14', methods=['GET', 'POST'])
def download_Event_report(**kwargs):
        
        new_var = calendar_events.query.all()
       
        df = pd.DataFrame(columns=['Topic Name','Department','Level','Start Time','End Time','Date','GDCs','Meeting URL','Sub Topic'])
        for i in new_var:
           df = df._append({'Topic Name': i.title,
                   'Department': i.dept,
                   'Level': i.level,
                   'Start Time': i.Time,
                   'End Time': i.Etime,
                   'Date': i.date,
                   'GDCs': i.NoOfGDC,
                   'Meeting URL': i.url,
                   'Sub Topic': i.SPSubTopic,
                   },ignore_index=True)
           
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
           df.to_excel(writer, index=False, sheet_name='Sheet1')
       # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=Calender_Event_Report.xlsx"})
          
#Download the attendance report - ADMIN
@app.route('/download/report/excel25', methods=['GET', 'POST'])
def download_report_attn(**kwargs):
       
        new_var = AttendanceUploadByAdmin.query.all()
        
        df = pd.DataFrame(columns=['Subject','Function','Email','First_Join','Last_Leave','Duration','Feedback','Period','Year'])
        for i in new_var:
            df = df._append({'Subject': i.Subject,
                    'Function': i.Dept,
                    'Email': i.Email,
                    'First_Join': i.First_Join,
                    'Last_Leave': i.Last_Leave,
                    'Duration': i.Duration,
                    'Feedback': i.Feedback,
                    'Period': i.Period,
                    'Year': i.Year,
                    'UploadBy': i.uploadby,
                    },ignore_index=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
           df.to_excel(writer, index=False, sheet_name='Sheet1')
       # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=AttendanceReport.xlsx"})
        
#Download the question bank report - ADMIN
@app.route('/download/report/excel/<getInfoTopic>', methods=['GET', 'POST'])
def download_report(**kwargs):
        print(f'Form Data: {request.form}')
        getInfoTopic = kwargs.get('getInfoTopic')
       
        new_var = topic_question.query.filter_by(q_topic=getInfoTopic).all()
       
        df = pd.DataFrame(columns=['Function','choices_no','Points','Question_Text','Option 1','Option 2','Option 3','Option 4','Correct Answer'])
        for i in new_var:
            df = df._append({'Function': i.dept_Func,
                    'choices_no': i.choices_no,
                    'Points': i.q_points,
                    'Question_Text': i.q_text,
                    'Option 1': i.q_choice1,
                    'Option 2': i.q_choice2,
                    'Option 3': i.q_choice3,
                    'Option 4': i.q_choice4,
                    'Correct Answer': i.q_ans,
                    'UploadBy':i.uploadby
                    },ignore_index=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
           df.to_excel(writer, index=False, sheet_name='Sheet1')
       # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=Questionreport.xlsx"})
    

#Download the Assessment topic report - ADMIN
@app.route('/download/report/excel13', methods=['GET', 'POST'])
def download_SUP_report(**kwargs):
        
        new_var = SUP_topic.query.all()
        
        df = pd.DataFrame(columns=['Topic Name','Question Type','Function','Level','Mandatory','Time in Min','Pass Mark','No of GDCs','Sub Topic'])
        for i in new_var:
            df = df._append({'Topic Name': i.topics_added,
                    'Question Type': i.question_type,
                    'Function': i.Department,
                    'Level': i.Level,
                    'Mandatory': i.Mandatory,
                    'Time in Min': i.TimeInMin,
                    'Pass Mark': i.PassMark,
                    'No of GDCs': i.NoOfGDC,
                    'Sub Topic': i.SPSubTopic,
                    'UploadBy':i.uploadby
                    },ignore_index=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
          df.to_excel(writer, index=False, sheet_name='Sheet1')
      # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=AssessmentTopicReport.xlsx"})
       
#Not used - USER
@app.route('/TalentMeter',methods=['GET','POST'])
def TalentMeter():
    iframe = "https://app.powerbi.com/reportEmbed?reportId=e3be424b-dff9-4638-9f29-efc319f3d42f&autoAuth=true&ctid=1e355c04-e0a4-42ed-8e2d-7351591f0ef1"
    return render_template('talentmeter.html', iframe=iframe)


#View OJT SME details - ADMIN
@app.route('/GiveAccess1')
def GiveAccess1():   
 
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
      
    formToUpdate = getAccess()
    formToUpdate1 = RmAccess()
    q= request.args.get('q')
    if q:
        page = request.args.get('page', 1, type=int)
        items=UserDataNew.query.filter(UserDataNew.email_add.contains(q)).paginate(page=page, per_page=10)
        
        # items = OJTNew.query.paginate(page=page, per_page=2)
        return render_template("access1.html",items=items,formToUpdate=formToUpdate,formToUpdate1=formToUpdate1,user_type=user_type) 
    else:
        # q='sme'
        page = request.args.get('page', 1, type=int)
        # items = UserData.query.filter_by(user_type="user").paginate(page=page, per_page=2)
        items = UserDataNew.query.filter(UserDataNew.user_type.contains("abc")).paginate(page=page, per_page=10)
        # print(items)
        return render_template("access1.html",items=items,formToUpdate=formToUpdate,formToUpdate1=formToUpdate1,user_type=user_type)  
    

#Provide OJT SME access - ADMIN
@app.route('/SaveAccess1',methods=['GET','POST'])
def SaveAccess1():
    formToUpdate = getAccess()
    if request.method == "POST" and formToUpdate.validate_on_submit():
        insert = request.form.get('GetEmail_ID') 
        quest = UserDataNew.query.filter_by(email_add=insert).first()
        quest.user_type=quest.user_type + ",sme"
        db.session.commit()
        flash(f'Access granted to the user', category='success')
        return redirect(url_for('GiveAccess1'))
   

#Decline OJT SME access - ADMIN
@app.route('/DeleteAccess1',methods=['GET','POST'])
def DeleteAccess1():
    formToUpdate1 = RmAccess()
    if request.method == "POST" and formToUpdate1.validate_on_submit():
        insert = request.form.get('GetEmail_ID') 
        quest = UserDataNew.query.filter_by(email_add=insert).first()
        str_val = quest.user_type
        str_val = str_val.replace(",sme","")
        quest.user_type = str_val
        db.session.commit()
        flash(f'Access declined to the user', category='success')
        return redirect(url_for('GiveAccess1'))
        
#Not used - ADMIN
@app.route('/report_page')
def report_page():
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    q= request.args.get('q')
    cnt_items = OJTNew.query.filter_by(TrainerEmail=current_user.email_add).count()
    if q:
        page = request.args.get('page', 1, type=int)
        odata = OJTNew.query.filter(OJTNew.TraineeEmail.contains(q)).paginate(page=page, per_page=2)
    else:
        page = request.args.get('page', 1, type=int)
        odata = OJTNew.query.filter_by(TrainerEmail=current_user.email_add).paginate(page=page, per_page=2)
    if cnt_items==0:
        flash('There are no OJT feedbacks in data.',category='danger')
        return redirect(url_for('admSide'))
    return render_template("report.html",odata=odata,user_type=user_type)
    
    
#View OJT report - ADMIN
@app.route('/report_page_ad')
def report_page_ad():
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type    
    
    cnt_items = OJTNew.query.count()
    if cnt_items==0:
        flash('There are no OJT feedbacks in data',category='danger')
        return redirect(url_for('admSide'))
    
    week_number = request.args.get('weekNumber', None)
    submitted_by = request.args.get('submittedBy', None)
    page = request.args.get('page', 1, type=int)
    
    filters = []
    
    if week_number:
        filters.append(OJTNew.WeekNumber == week_number)
    if submitted_by:
        filters.append(OJTNew.TrainerEmail == submitted_by)
    
    odata = OJTNew.query.filter(*filters).paginate(page=page, per_page=10)
    
    
    # Pass week_numbers and submitters to the template
    week_numbers = OJTNew.query.with_entities(OJTNew.WeekNumber).distinct().all()
    submitters = OJTNew.query.with_entities(OJTNew.TrainerEmail).distinct().all()
    
    # Get distinct week numbers and submitters
    week_numbers = [str(item[0]) for item in OJTNew.query.with_entities(OJTNew.WeekNumber).distinct().all()]
    submitters = [item[0] for item in OJTNew.query.with_entities(OJTNew.TrainerEmail).distinct().all()]
    
    return render_template("report_ad.html",odata=odata,week_numbers=week_numbers, submitters=submitters,user_type=user_type)   


#Update question details - ADMIN
@app.route('/UpdateQn/<q_ID>',methods=['GET','POST'])
def UpdateQn(q_ID):
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    add_q_form = add_question()
    qntype = "obj"
    topic_q = topic_question.query.all()
    event = topic_question.query.filter_by(q_ID=q_ID).first()
    topicname = event.q_topic
    topicINF = topic_question.query.filter_by(q_topic=topicname).all()
    dept_name = event.dept_Func
    level = event.level
    deleteEvent = DeleteFormQNR()
    if request.method == 'POST':
       if request.form['choices_no']=="2" and (request.form['q_choice3']!="" or request.form['q_choice4']!=""):
             flash(f'Please keep the option 3 and option 4 text boxes blank for dual choice question', category='danger')
             return render_template('admin_UpdateQ.html', user_type=user_type,dept_name=dept_name,topicname=topicname,qntype=qntype,event=event,add_q_form=add_q_form)
       if request.form['choices_no']=="2" and (request.form['q_ans']=="Option 3" or request.form['q_ans']=="Option 4"):
             flash(f'You cannot chose option 3 or option 4 text boxes for dual choice question', category='danger')
             return render_template('admin_UpdateQ.html', user_type=user_type,dept_name=dept_name,topicname=topicname,qntype=qntype,event=event,add_q_form=add_q_form) 
       if request.form['choices_no']=="2" and (request.form['q_choice1']==request.form['q_choice2']):
              flash(f'Please keep the option 1 and option 2 labels unique', category='danger')
              return render_template('admin_UpdateQ.html', user_type=user_type,dept_name=dept_name,topicname=topicname,qntype=qntype,event=event,add_q_form=add_q_form)
       if request.form['choices_no']=="4" and (request.form['q_choice1']==request.form['q_choice2'] or request.form['q_choice1']==request.form['q_choice3'] or request.form['q_choice1']==request.form['q_choice4'] or request.form['q_choice2']==request.form['q_choice3'] or request.form['q_choice2']==request.form['q_choice4'] or request.form['q_choice3']==request.form['q_choice4']):
              flash(f'Please provide unique value for the options', category='danger')
              return render_template('admin_UpdateQ.html', user_type=user_type,dept_name=dept_name,topicname=topicname,qntype=qntype,event=event,add_q_form=add_q_form) 
       if request.form['choices_no']=="4" and (request.form['q_choice3']=="" or request.form['q_choice4']==""):
             flash(f'Please provide value for option 3 and option 4 text boxes', category='danger')
             return render_template('admin_UpdateQ.html', user_type=user_type,dept_name=dept_name,topicname=topicname,qntype=qntype,event=event,add_q_form=add_q_form)
       
       if event:
           db.session.delete(event)
           db.session.commit()
           dept_Func = request.form['dept_Func']
           q_points = request.form['q_points']
           q_ans = request.form['q_ans']
           q_type = request.form['q_type']
           q_choice4 = request.form['q_choice4']
           q_choice3 = request.form['q_choice3']
           q_choice2 = request.form['q_choice2']
           q_choice1 = request.form['q_choice1']
           q_text = request.form['q_text']
           choices_no = request.form['choices_no']
           q_topic = request.form['q_topic']
           question_type = request.form['question_type']
           q_ID = request.form['q_ID']
           level = level
           event = topic_question(q_type=q_type,level=level,dept_Func=dept_Func, q_points=q_points, q_ans=q_ans, q_choice4=q_choice4, q_choice3 = q_choice3,q_choice2=q_choice2,q_choice1=q_choice1,q_text=q_text,choices_no=choices_no,q_topic=q_topic,question_type=question_type,q_ID=q_ID,uploadby=current_user.email_add)
           db.session.add(event)
           db.session.commit()
           flash('Data updated successfully',category='success')
           return render_template('admin_addQ.html', user_type=user_type,getLevel=level,dept_name=dept_name,add_q_form=add_q_form, topic_q=topic_q, getInfoTopic=topicname, topicINF=topicINF, getTypeQ=qntype, deleteEvent=deleteEvent) 
    return render_template('admin_UpdateQ.html',user_type=user_type,getLevel=level,topicname=topicname,qntype=qntype,event=event,add_q_form=add_q_form)

#Add OJT feedback - SME
@app.route('/ojt_page',methods=['GET','POST'])
def ojt_page():
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    form = OJTForm()
    if form.validate_on_submit():
         service = OJTNew.query.filter_by(WeekNumber=form.WeekNumber.data,TraineeEmail=form.TraineeEmail.data).first()
         if service is not None:
             flash(f'You have already provided feedback for this trainee for week {form.WeekNumber.data}',category='danger')
         else:
            week_to_create = OJTNew(TraineeName=form.TraineeName.data,
                                    TrainerEmail = current_user.email_add,
                                  TraineeEmail=form.TraineeEmail.data,
                                  WeekNumber=form.WeekNumber.data,
                                  rate_1=form.rate_1.data,
                                  rate_2=form.rate_2.data,
                                  rate_3=form.rate_3.data,
                                  Strengths=form.Strengths.data,
                                  Todo_better=form.Todo_better.data,
                                  FeedbackLogDate = date.today(),
                                  uploadby=current_user.id)
            db.session.add(week_to_create)
            db.session.commit()        

            flash(f'Feedbacks marked successfully!',category='success')
            return redirect(url_for('report_page'))
       
    # if form.errors != {}:
    #     for err_msg in form.errors.values():
    #         flash(f'There was an error while filling attendance :{err_msg}',category='danger')
    return render_template('ojtform.html',form=form,user_type=user_type)

#Enable the email generation button - ADMIN
@app.route('/revokemail',methods=['GET','POST'])
def revokemail():
    
    formUpdate = EnableMail()
    if formUpdate.validate_on_submit():
        title_value = request.form.get('GetTitle') 
        dep_value = request.form.get('GetDept') 
        quest = calendar_events.query.filter_by(title=title_value,dept=dep_value).first()
        quest.Mail="NotSent"
        quest.ReMail="Disable"
        db.session.commit()
        return redirect(url_for('sendmail')) 
    
#Not used - ADMIN
@app.route('/sentmail',methods=['GET','POST'])
def sentmail():
    formToUpdate = SendMail()
    formUpdate = EnableMail()
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_type = user_data_new.user_type
    cal_items = calendar_events.query.all()
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    topic_count = training_topic_count.query.filter_by(year=year,period=period).all()
    if formToUpdate.validate_on_submit():
        title_value = request.form.get('GetTitle') 
        dep_value = request.form.get('GetDept') 
        quest = calendar_events.query.filter_by(title=title_value,dept=dep_value).first()
        quest.Mail="sent"
        quest.ReMail="Enable"
        alluser = training_topic_nomination.query.filter_by(year=year,period=period,topics_added=title_value,Dept=dep_value).all()
        if len(alluser)==0:
            flash(f'No nominations for this topic', category='danger')
            return render_template('sendmail.html',topic_count=topic_count,cal_items=cal_items,formToUpdate=formToUpdate,formUpdate=formUpdate,user_type=user_type)

        str1 = ""
        i = 1
      
        # traverse in the string
        for ele in alluser:
            mailid = ele.NominatedBy
            mailad = UserDataNew.query.filter_by(id=mailid).first()
            mailaddress = mailad.email_add
            if i > 1:
                str1 = str1 + ';'
            str1 += mailaddress
            i+=1
       
        clist = training_topic_count.query.filter_by(year=year,period=period,topic_name=title_value,Dept=dep_value).first()
        clist.ContactList = str1
        
        
        const=win32com.client.constants
        pythoncom.CoInitialize()
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        # mailItem = obj.CreateItem(olMailItem)
        mailItem = obj.CreateItem(1)
        mailItem.Subject = 'Training Session - ' + str(title_value) 
        mailItem.BodyFormat = 2
        # mailItem.location = "test_location"
        mailItem.body = "Dear Team, \r\n\r\nA training for " + str(title_value) + " session is assigned on the following date. Please attend the session without fail.\r\n\r\nRegards\r\nCI Team"
        # input = quest.date
        # format = '%Y-%m-%d'
        # import datetime 
        # # convert from string format to datetime format
        # dt = datetime.datetime.strptime(input, format).date()
        concat_time = quest.date + " " + quest.Time + ":00"
        mailItem.start = concat_time
        start = datetime.strptime(quest.Time, "%H:%M") 
        end = datetime.strptime(quest.Etime, "%H:%M") 
          
        difference = end - start 
          
        seconds = difference.total_seconds() 
          
        minutes = seconds / 60
        
        mailItem.duration = minutes
        mailItem.importance = 2
        mailItem.meetingstatus = 1
        required = mailItem.Recipients.add(str1)
        required.Type = 1
        optional = mailItem.Recipients.add("harish.satyan@kantar.com")
        optional.Type = 2
        # mailItem.respond(Response(True))
        # mailItem.display()
        db.session.commit()
        # flash(f'Emails sent successfully', category='success')
        # return redirect(url_for('sendmail')) 
        return render_template('sendmail.html',str1=str1,topic_count=topic_count,cal_items=cal_items,formToUpdate=formToUpdate,formUpdate=formUpdate,user_type=user_type)

#Add bulk questions to the table - ADMIN
@app.route('/uploadFiles/<getInfoTopic>/<getTypeQ>/<getLevel>', methods=['POST', 'GET'])
def uploadFiles(**kwargs):
      # get the uploaded file
      uploaded_file = request.files['file']
      getInfoTopic = kwargs.get('getInfoTopic')
      getTypeQ = kwargs.get('getTypeQ')
      topicINF_dept = SUP_topic.query.filter_by(topics_added=getInfoTopic).first()
      dept_name = topicINF_dept.Department
      getLevel = kwargs.get('getLevel')
      form = add_question()
      value = 0
     
      user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()   
      user_type = user_data_new.user_type
      
      # Accessing app-level configuration
      upload_folder = app.config['UPLOAD_FOLDER']
      
      if uploaded_file.filename != '':
           file_path = os.path.join(upload_folder, uploaded_file.filename)
          # set the file path
           uploaded_file.save(file_path)
           col_names = ['question_type','q_topic', 'choices_no', 'q_text' , 'q_choice1', 'q_choice2', 'q_choice3', 'q_choice4','q_ans','q_points','Function','level']
           
           csvData = pd.read_excel(uploaded_file)
           
            # Loop through the Rows
           for i,row in csvData.iterrows():
               if (row['q_choice1']==row['q_choice2']) or (row['q_choice1']==row['q_choice3']) or (row['q_choice1']==row['q_choice4']) or (row['q_choice2']==row['q_choice3']) or (row['q_choice2']==row['q_choice4']) or (row['q_choice3']==row['q_choice4']):
                   value = 1
           if value==1:
               flash(f'There are some duplicate values present between the options, please recheck the data and upload the file again', category='success')
               return redirect(url_for('info_topic',user_type=user_type,dept_name=dept_name,getInfoTopic=getInfoTopic, getTypeQ=getTypeQ,getLevel=getLevel))
  
           for i,row in csvData.iterrows():
                   week_to_create = topic_question(
                                           question_type = row['question_type'],
                                         q_topic=row['q_topic'],
                                         choices_no=row['choices_no'],
                                         q_text=row['q_text'],
                                         q_choice1=row['q_choice1'],
                                         q_choice2=row['q_choice2'],
                                         q_choice3=row['q_choice3'],
                                         q_choice4=row['q_choice4'],
                                         q_ans=row['q_ans'],
                                         q_points=row['q_points'],
                                         dept_Func=row['Function'],
                                         level=row['level'],
                                         q_type = row['type'],
                                         uploadby=current_user.email_add)
                  
                   db.session.add(week_to_create)
                   db.session.commit()     
                   flash(f'Data uploaded successfully', category='success')
           return redirect(url_for('info_topic',user_type=user_type,dept_name=dept_name,getInfoTopic=getInfoTopic, getTypeQ=getTypeQ,getLevel=getLevel))
          # save the file
      else:
           flash(f'Please add file to proceed', category='danger')
           return redirect(url_for('info_topic',user_type=user_type,dept_name=dept_name,getInfoTopic=getInfoTopic, getTypeQ=getTypeQ,getLevel=getLevel))

#Show results for subjective assessment - USER
@app.route("/sub_test_results",methods=['GET','POST'])
def sub_test_results():
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    todays_date = date.today() 

    date_val = date_table.query.first()
    nominate_val = date_val.startR
    IDlist = nominate_val.split('-')
    N_year = int(IDlist[0])
    N_month = int(IDlist[1])
    N_day = int(IDlist[2])
    day = todays_date.day
    month = todays_date.month
    year1 = todays_date.year
    d1 = date(year1,month,day)
    d2 = date(N_year,N_month,N_day)
    
    if d1 < d2:
        flash(f'The results are yet to published, please wait.', category='success')
        return redirect(url_for('admSide'))  
    user_sub_data = SupResultStatusByAdmin.query.filter_by(year=year,period=period,mailid=current_user.email_add).all()
    user_sub_data_cnt = SupResultStatusByAdmin.query.filter_by(year=year,period=period,mailid=current_user.email_add).count()
    if user_sub_data_cnt==0:
        flash(f'No results to be published for this year.', category='success')
        return redirect(url_for('admSide'))  
    return render_template('user_sub_result.html', user_sub_data=user_sub_data,user_type=user_type)

#Subjective result upload module - ADMIN
@app.route("/ResultFileSub",methods=['GET','POST'])
def ResultFileSub():
   
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
   
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    
    # Accessing app-level configuration
    upload_folder = app.config['UPLOAD_FOLDER']
   
    if request.method == 'POST':
        file = request.files['file']      
     
        if file.filename != '':
           
           file_path = os.path.join(upload_folder, file.filename)
          # set the file path
           file.save(file_path)
           col_names = ['mailid','topic', 'level', 'function' , 'marks', 'status','year','period']
            # Use Pandas to parse the CSV file
           # csvData = pd.read_csv(file_path,names=col_names, header=0)
           csvData = pd.read_excel(file)
            # Loop through the Rows
           for i,row in csvData.iterrows():
                    result_to_create = SupResultStatusByAdmin(
                                  mailid = row['mailid'],
                                  topic=row['topic'],
                                  level=row['level'],
                                  department=row['function'],
                                  marks=row['marks'],
                                  status=row['status'],
                                  year=row['year'],
                                  period=row['period'],
                                  comment=row['Comments'],
                                  uploadby=current_user.email_add
                                 )
                    db.session.add(result_to_create)
                    db.session.commit()    
                    # else:
                    #     print("Pass")
                   
        flash(f'File uploaded successfully', category='success')
        return redirect(url_for('ResultFileSub'))
   
    q= request.args.get('q')
    if q:
        page = request.args.get('page', 1, type=int)
        odata = SupResultStatusByAdmin.query.filter(SupResultStatusByAdmin.Topicname.contains(q)).paginate(page=page, per_page=8)
    else:
        page = request.args.get('page', 1, type=int)
        odata = SupResultStatusByAdmin.query.filter_by(year=year,period=period).paginate(page=page, per_page=8)
       
    return render_template('result_upload.html',odata=odata,user_type=user_type)

#Result template for Subjective assessment - ADMIN
@app.route('/download_Restemplate', methods=['GET', 'POST'])
def download_Restemplate(**kwargs):
       
        df = pd.DataFrame({'mailid': ['abc.xyz@kantar.com', 'def.xyz@kantar.com'],
                       'topic': ['Conjoint_L1_SP', 'Matrix_L1_SP'],
                       'level': ['L1','L1'],
                       'function': ['SP','SP'],
                       'marks': [80,60],
                       'status': ['Pass', 'Fail'],
                       'year': ['2024', '2024'],
                       'period': ['H1', 'H1'],
                       'Comments': ['Pass', 'Fail, the script is incorrect']
                       })
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
    # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=ResultTemplate.xlsx"})
       
#Not used - ADMIN
@app.route('/sendmail')
def sendmail():
   
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
     
    cal_items = calendar_events.query.all()
    topic_count = training_topic_count.query.filter_by(year=year,period=period).all()
  
    formToUpdate = SendMail()
    formUpdate = EnableMail()
    return render_template('sendmail.html',topic_count=topic_count,cal_items=cal_items,formToUpdate=formToUpdate,formUpdate=formUpdate,user_type=user_type)

#Add adhoc training - USER
@app.route("/training_request", methods=['GET','POST'])
def trainingrequest():
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    user_location = user_data_new.GDCSelect
    user_cluster = user_data_new.Cluster
          
    form = TrainingRequestForm()
    if form.validate_on_submit():
            request_to_create = TrainingRequest(location = user_location,
                                    user_name = current_user.email_add,
                                    user_dept = user_dept,
                                    Topicname = form.Topicname.data,
                                    cluster = user_cluster,
                                    Department = form.Department.data,
                                    PeopleCount = form.PeopleCount.data,
                                    Description = form.Description.data,
                                    SMEContact = form.SMEContact.data
                                 )
            db.session.add(request_to_create)
            db.session.commit()        
            
            To = "harish.satyan@kantar.com"
            Cc = "deepak.sartape@kantar.com"
            Subject = 'ADHOC TRAINING REQUEST -'  + str(form.Topicname.data) 
            Description = "We have a training request in place for a new topic " + str(form.Topicname.data) +". Please review and take actions."
          
         # if form.validate_on_submit():
            const=win32com.client.constants
            pythoncom.CoInitialize()
            olMailItem = 0x0
            obj = win32com.client.Dispatch("Outlook.Application")
            mailItem = obj.CreateItem(olMailItem)
            mailItem.Subject = Subject
            mailItem.BodyFormat = 2
            mailItem.HTMLBody = Description
            # mailItem.Attachment = file
            # mailItem.respond(Response(True))
            # mailItem.display()
            mailItem.To = To
            mailItem.CC = Cc
            # mailItem.display()
            # mailItem.Send()
            flash('Request raised successfully!',category='success')
            return redirect(url_for('admSide'))
       
       
    return render_template('training_request.html',form=form,user_type=user_type)


#Download OJT feedback report - ADMIN
@app.route('/download/report/excel', methods=['GET', 'POST'])
def download_OJT_report(**kwargs):
        print(f'Form Data: {request.form}')
        new_var_1 = OJTNew.query.all()
        
        df = pd.DataFrame(columns=['ID','Trainee Email','Week Number','Scripting Skills','Learning Mindset','Team Management','Strength','To Do Better','Posted By','Date'])
        for i in new_var_1:
           df = df._append({'ID': i.ojtid,
                   'Trainee Email': i.TraineeEmail,
                   'Week Number': i.WeekNumber,
                   'Scripting Skills': i.rate_1,
                   'Learning Mindset': i.rate_2,
                   'Team Management': i.rate_3,
                   'Strength': i.Strengths,
                   'To Do Better': i.Todo_better,
                   'Posted By': i.TrainerEmail,
                   'Date': i.FeedbackLogDate,
                   },ignore_index=True)
         # Save DataFrame to a BytesIO object
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
          df.to_excel(writer, index=False, sheet_name='Sheet1')
      # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=OJT_Report.xlsx"})
        

#Download Attendance report - ADMIN
@app.route('/download/report/excel1', methods=['GET', 'POST'])
def download_Attendance_report(**kwargs):
        print(f'Form Data: {request.form}')
        new_var_2 = Attendance.query.all()
        
        df = pd.DataFrame(columns=['ID','Attendance confirmation','Location','Trainer Name','Function','Topic Name','The Objectives of the training were clearly defined','Participation and interaction were encouraged','Delivery of the learning session','Did the trainer demonstrate the topic with relevant examples','What did you like about the session ?','What can make the session even better ?','What did the SME do well in the session ?','What could the SME do to make the session even better ?','Any Suggestion','Date'])
        for i in new_var_2:
           df = df._append({'ID': i.id,
                   'Attendance confirmation': i.confirm_attendance,
                   'Location': i.location,
                   'Trainer Name': i.name,
                   'Function': i.user_dept,
                   'Topic Name': i.Topicname,
                   'The Objectives of the training were clearly defined': i.rate_1,
                   'Participation and interaction were encouraged': i.rate_2,
                   'Delivery of the learning session': i.rate_3,
                   'Did the trainer demonstrate the topic with relevant examples':i.rate_4,
                   'What did you like about the session ?':i.session_like,
                   'What can make the session even better ?':i.session_better,
                   'What did the SME do well in the session ?':i.session_do_well,
                   'What could the SME do to make the session even better ?':i.session_even_better,
                   'Any Suggestion':i.suggestion,
                   'Date':i.FeedbackDate
                   },ignore_index=True)
         # Save DataFrame to a BytesIO object
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
     # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=Attendance_Report.xlsx"})
        
#Download Incident report - ADMIN
@app.route('/download/report/excel2', methods=['GET', 'POST'])
def download_all_incident_report(**kwargs):
    print(f'Form Data: {request.form}')
    new_var_3 = IncidentTbl.query.all()
    
   
    df = pd.DataFrame(columns=['Incident ID','User Name','Location','Module Name','Department','Subject','Description','Raised Date','Closed Date','Status','AdminComments'])
    for i in new_var_3:
        df = df._append({'Incident ID': i.inc_id,
                'User Name': i.user_name,
                'Location': i.location,
                'Module Name': i.modulename,
                'Department': i.user_dept,
                'Subject': i.Subject,
                'Description': i.Description,
                'Raised Date': i.Raiseddate,
                'Closed Date': i.Closuredate,
                'Status':i.Status,
                'AdminComments':i.Admincomment,
                },ignore_index=True)
         # Save DataFrame to a BytesIO object
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
           df.to_excel(writer, index=False, sheet_name='Sheet1')
       # Seek to the beginning of the BytesIO object
    output.seek(0)
    return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=Incident_Report.xlsx"})

# NominatedVsCompleted Report - ADMIN
@app.route('/download/report/excel32', methods=['GET', 'POST'])
def download_NVSC_report(**kwargs):
        print(f'Form Data: {request.form}')
        new_var_3 = sup_nominated_count.query.all()
      
        df = pd.DataFrame(columns=['TopicID','Topic Name','Function','Completed Count','InComplete Count','Year','Period'])
        for i in new_var_3:
           df = df._append({'TopicID': i.topic_ID,
                   'Topic Name': i.topic_name,
                   'Function': i.Dept,
                   'Completed Count': i.Count,
                   'InComplete Count': i.IncompleteCount,
                   'Year': i.year,
                   'Period': i.period,
                   },ignore_index=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
         df.to_excel(writer, index=False, sheet_name='Sheet1')
     # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=NominatedVsCompleted_Report.xlsx"})   
           
#Download Nomination report - ADMIN
@app.route('/download/report/excel3', methods=['GET', 'POST'])
def download_Nomination_report(**kwargs):
        print(f'Form Data: {request.form}')
        new_var_3 = Nominate_topic_by.query.all()
      
        df = pd.DataFrame(columns=['ID','Topic Name','Type','Department','Nominated By','Status','Nomination Date','Year','Period'])
        for i in new_var_3:
           df = df._append({'ID': i.topic_ID,
                   'Topic Name': i.topics_added,
                   'Type': i.question_type,
                   'Department': i.Department,
                   'Nominated By': get_username_by_id(i.nominatedBY),
                   'Status': i.SUP_status,
                   'Nomination Date': i.NominationDate,
                   'Year': i.year,
                   'Period': i.period,
                   },ignore_index=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
         df.to_excel(writer, index=False, sheet_name='Sheet1')
     # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=Nomination_Report.xlsx"})   
           
       
# Download subjective result data - ADMIN
@app.route('/download/report/excel7', methods=['GET', 'POST'])
def download_subj_result(**kwargs):
        print(f'Form Data: {request.form}')
        new_var_7 = SupResultStatusByAdmin.query.all()
        df = pd.DataFrame(columns=['ID','Mail ID','Topic','Level','Function','Marks','Status','Year','Period'])
        for i in new_var_7:
            df = df._append({'ID': i.id,
                    'Mail ID': i.mailid,
                    'Topic': i.topic,
                    'Level': i.level,
                    'Function': i.department,
                    'Marks': i.marks,
                    'Status': i.status,
                    'Year': i.year,
                    'Period': i.period,
                    'UploadBy': i.uploadby
                    },ignore_index=True)
      
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
          df.to_excel(writer, index=False, sheet_name='Sheet1')
      # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=Subjective_Result_File.xlsx"})
        
# Download SME report - ADMIN
@app.route('/download/report/excel34/<dept>', methods=['GET', 'POST'])
def download_SME_report(**kwargs):
        print(f'Form Data: {request.form}')
        gDept = kwargs.get('dept')
      
        if gDept == "none":
            new_var_4 = calendar_events.query.all()
        else:
            new_var_4 = calendar_events.query.filter_by(dept=gDept).all()
        
        df = pd.DataFrame(columns=['Topic Name','SME'])
        for i in new_var_4:
          df = df._append({'Topic Name': i.title,
                  'SME': i.sme,
                  },ignore_index=True)
      
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
         df.to_excel(writer, index=False, sheet_name='Sheet1')
     # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=SME_Report.xlsx"})
        
# Download Training Nomination Count report - ADMIN    
@app.route('/download/report/excel4/<dept>', methods=['GET', 'POST'])
def download_NominationCount_report(**kwargs):
        print(f'Form Data: {request.form}')
        gDept = kwargs.get('dept')
       
        if gDept == "none":
            new_var_4 = training_topic_count.query.all()
        else:
            new_var_4 = training_topic_count.query.filter_by(Dept=gDept).all()
       
        df = pd.DataFrame(columns=['ID','Topic Name','Function','Nominated Count','Year','Period'])
        for i in new_var_4:
          df = df._append({'ID': i.topic_ID,
                  'Topic Name': i.topic_name,
                  'Function': i.Dept,
                  'Nominated Count': i.Count,
                  'Year': i.year,
                  'Period': i.period,
                  },ignore_index=True)
        
    
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
         df.to_excel(writer, index=False, sheet_name='Sheet1')
     # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=NominationCount_Report.xlsx"})
        
# Download_QuestionBank_report - ADMIN
@app.route('/download/report/excel5', methods=['GET', 'POST'])
def download_QuestionBank_report(**kwargs):
        print(f'Form Data: {request.form}')
        new_var_5 = topic_question.query.all()
        
        df = pd.DataFrame(columns=['ID','Function','Topic Name','Choices','Points','Question Text','Option 1','Option 2','Option 3','Option 4','Correct Answer'])
        for i in new_var_5:
            df = df._append({'ID': i.q_ID,
                    'Function': i.dept_Func,
                    'Topic Name': i.q_topic,
                    'Choices': i.choices_no,
                    'Points': i.q_points,
                    'Question Text': i.q_text,
                    'Option 1': i.q_choice1,
                    'Option 2': i.q_choice2,
                    'Option 3': i.q_choice3,
                    'Option 4': i.q_choice4,
                    'Correct Answer':i.q_ans
                    },ignore_index=True)
       
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
           df.to_excel(writer, index=False, sheet_name='Sheet1')
       # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=QuestionBank_Report.xlsx"})
        
# Download_UserAnswerReport - ADMIN
@app.route('/download/report/excel6', methods=['GET', 'POST'])
def download_UserAnswerReport_report(**kwargs):
        print(f'Form Data: {request.form}')
        new_var_6 = answerPerQnr.query.all()
        
        df = pd.DataFrame(columns=['Level','User Name','User Email','Topic Name','Question Id','Question Text','User Answer','Correct Answer','Points'])
        for i in new_var_6:
            df = df._append({'Level': i.level,
                'User Name': get_username_by_id(i.useraccount),
                'User Email': get_useremail_by_id(i.useraccount),
                'Topic Name': i.session_id,
                'Question Id': i.question_id,
                'Question Text': get_qnname_by_id(i.question_id),
                'User Answer': i.useranswer,
                'Correct Answer': i.correctanswer,
                'Points': i.q_points,
                },ignore_index=True)
       
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
           df.to_excel(writer, index=False, sheet_name='Sheet1')
       # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=Answer_Per_Question_Report.xlsx"})
         

#Download user files for a specific topic - ADMIN
@app.route('/download_topic_files', methods=['GET', 'POST'])
def download_topic_files():
    subject = request.form.get('subject')
    uploads = Upload.query.filter_by(topic_added=subject).all()
    username = current_user.username
    username1 = os.environ.get('USERNAME')
    zip_filename = f"C:/Users/{username1}/Downloads/{subject}_files.zip"

    
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for upload in uploads:
            data = upload.data
            file_name = upload.filename
            zipf.writestr(file_name, data)
           
    return send_file(zip_filename, as_attachment=True)
    
def get_username_by_id(user_id):
    # Logic to fetch username from user_data table based on user id
    # Return the username    
    user_data = UserDataNew.query.filter_by(id=user_id).first()
   
    if user_data:
        return user_data.username # Assuming the column name for username is 'username'
    else:
        return None  # Return None if user with given id is not found
    
def get_useremail_by_id(user_id):
    # Logic to fetch username from user_data table based on user id
    # Return the username    
    user_data = UserDataNew.query.filter_by(id=user_id).first()
 
    if user_data:
        return user_data.email_add # Assuming the column name for username is 'username'
    else:
        return None  # Return None if user with given id is not found    

def get_qnname_by_id(qn_id):
    # Logic to fetch username from user_data table based on user id
    # Return the username    
    user_data = topic_question.query.filter_by(q_ID=qn_id).first()
  
    if user_data:
        return user_data.q_text # Assuming the column name for username is 'username'
    else:
        return None  # Return None if user with given id is not found
    
    
# ------------------------- USER MODULE -----------------------------------
#Upload bulk OJT data - SME
@app.route("/uploadOJT", methods=['POST'])
def uploadOJT():
      # get the uploaded file
      print("UploadOJT route triggered")

      if 'file' not in request.files:
        flash('No file part', category='danger')
        return redirect(url_for('uploadOJTPage'))

      uploaded_file = request.files['file']
      upload_folder = app.config['UPLOAD_FOLDER']
      
      if uploaded_file.filename != '':
           file_path = os.path.join(upload_folder, uploaded_file.filename)
          # set the file path
           uploaded_file.save(file_path)
           col_names = ['email_add']
            # Use Pandas to parse the CSV file
           csvData = pd.read_excel(uploaded_file)
           
           duplicate_rows = csvData[csvData.duplicated(keep=False)]
           
           if not duplicate_rows.empty:
            flash(f'There are some duplicate values present in upload file, please recheck the data and upload the file again', category='danger')
            return redirect(url_for('uploadOJTPage'))
           
           
           for i,row in csvData.iterrows():
                   
                   email_add = row['email_add']
                   email_add_lower = email_add.lower()
        
                    # Case-insensitive email matching
                   existing_user = UserDataNew.query.filter(func.lower(UserDataNew.email_add) == email_add_lower).first()

                   if existing_user:
                    # Update the existing record
                    existing_user.user_type = "user,sme"
                    db.session.commit() 
                   else:
                    flash(f'Unable to find the email id : {email_add_lower}', category='danger')    
                       
           flash(f'Process done', category='success')
           return redirect(url_for('uploadOJTPage'))
          # save the file
      else:
           flash(f'Please add file to proceed', category='danger')
           return redirect(url_for('uploadOJTPage'))

# Download_OJT_template - SME
@app.route('/download_OJT_template', methods=['GET', 'POST'])
def download_OJT_template(**kwargs):
       
        df = pd.DataFrame({'email_add': ['test@kantar.com']
                  })
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
         df.to_excel(writer, index=False, sheet_name='Sheet1')
     # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=OJTAccessTemplate.xlsx"})

# Download_ojt_access_report - ADMIN
@app.route('/download/report/excel26', methods=['GET', 'POST'])
def download_ojt_report(**kwargs):
       
        # Initialize query
        query = UserDataNew.query

        # Apply filter for user_type if it includes 'sme'
        #if 'sme' in user_type.split(','):
        new_var = query.filter(UserDataNew.user_type.like('%sme%'))

        df = pd.DataFrame(columns=['email_add','user_type'])
        for i in new_var:
            df = df._append({'email_add': i.email_add,
                    'user_type': i.user_type,
                    },ignore_index=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
           df.to_excel(writer, index=False, sheet_name='Sheet1')
       # Seek to the beginning of the BytesIO object
        output.seek(0)
        return Response(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Content-Disposition": "attachment;filename=OJTAccessReport.xlsx"})
        #output in bytes

#Revoke the SME access - ADMIN
@app.route('/revoke_access', methods=['POST'])
def revoke_access():
    # Get email address from the form submission
    email_add = request.form.get('email_add')
    
    # Convert to lower case for case-insensitive comparison
    email_add_lower = email_add.lower()
    
    # Find the user with the given email address
    existing_user = UserDataNew.query.filter(func.lower(UserDataNew.email_add) == email_add_lower).first()
    
    if existing_user:
        # Update user_type to "user"
        existing_user.user_type = "user"
        db.session.commit()
        flash(f'Access revoked successfully for {email_add}')
    else:
        flash(f'No user found with email {email_add}')
    
    # Redirect back to the page with the list of users
    return redirect(url_for('uploadOJTPage'))

# View the Upload details for OJTPage - ADMIN
@app.route("/uploadOJTPage", methods=['GET', 'POST'])
def uploadOJTPage():
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
  
    date_val = date_table.query.first()
    period = date_val.period
    year = date_val.year
    q= request.args.get('q')

    # Initialize query
    query = UserDataNew.query

    query = query.filter(
        UserDataNew.user_type.like('%,sme,%') | 
        UserDataNew.user_type.like('sme,%') | 
        UserDataNew.user_type.like('%,sme') | 
        UserDataNew.user_type.like('sme')
    )

    # Get the page number from the request, default to 1 if not provided
    page = request.args.get('page', 1, type=int)

    # Paginate the query results
    odata = query.paginate(page=page, per_page=10)

    # Render the template with the paginated data and user_type
    return render_template('admin_ojt_upload.html', odata=odata, user_type=user_type)


#------------------------------- USER MODULE CODE ------------------------------
#View the list of departments - USER
@app.route("/TrainingNomination_Before", methods=['GET', 'POST'])
def TrainingNomination_Before():
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    date_val = date_table.query.first()
    distinct_depts = calendar_events.query.with_entities(calendar_events.dept).distinct().order_by(calendar_events.dept).all()
    return render_template('user_trainingNomination_Before.html',training_topics=distinct_depts,user_type=user_type)    

#View the list of SP departments - USER
@app.route("/TrainingNominationSP_Before", methods=['GET', 'POST'])
def TrainingNominationSP_Before():
    
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    date_val = date_table.query.first()
    distinct_depts = calendar_events.query.with_entities(calendar_events.SPSubTopic).distinct().all()
    return render_template('user_trainingNominationSP_Before.html',training_topics=distinct_depts,user_type=user_type)

#View the list of departments - USER
@app.route("/UserNomination_Before", methods=['GET', 'POST'])
def UserNomination_Before():
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    date_val = date_table.query.first()
    distinct_depts = SUP_topic.query.with_entities(SUP_topic.Department).distinct().order_by(SUP_topic.Department).all()
    return render_template('user_nomination_Before.html',training_topics=distinct_depts,user_type=user_type)

#View the list of SP departments - USER
@app.route("/UserNominationSP_Before", methods=['GET', 'POST'])
def UserNominationSP_Before():
     
    user_data_new = UserDataNew.query.filter_by(email_add=current_user.email_add).first()
    user_id = user_data_new.id
    user_dept = user_data_new.user_dept
    user_type = user_data_new.user_type
    date_val = date_table.query.first()
    distinct_depts = SUP_topic.query.with_entities(SUP_topic.SPSubTopic).distinct().all()
    return render_template('user_nominationSP_Before.html',training_topics=distinct_depts,user_type=user_type)