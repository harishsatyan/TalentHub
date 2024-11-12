from MyApp import db, login_manager
from MyApp import bycrpt, app
from flask_login import UserMixin
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text

@login_manager.user_loader
def load_user(user_id):
     return UserDataNew.query.get(int(user_id))


class GDCList(db.Model, UserMixin):
      __tablename__ = 'GDCList'
      id = db.Column(db.Integer(), primary_key=True )
      GDCSelect = db.Column(db.String(length=30), unique=True, nullable=False)
      Cluster = db.Column(db.String(length=50), unique=True, nullable=False)
      
class date_table(db.Model):
   __tablename__ = 'date_table'
   Date_ID = db.Column(db.Integer(), primary_key=True )
   start1 = db.Column(db.String(length=30), nullable=True )
   end1 = db.Column(db.String(length=30), nullable=True )
   start2= db.Column(db.String(length=30), nullable=True )
   end2 = db.Column(db.String(length=30), nullable=True )
   start3 = db.Column(db.String(length=50),nullable=True)
   end3 = db.Column(db.String(length=30), nullable=True )
   startR = db.Column(db.String(length=50),nullable=True)
   year = db.Column(db.String(length=50),nullable=True)
   period = db.Column(db.String(length=50),nullable=True)

class UserDataNew(db.Model, UserMixin):    
      __tablename__ = 'user_data_new'
      id = db.Column(db.Integer(), primary_key=True )
      username = db.Column(db.String(length=30), unique=True, nullable=False)
      email_add = db.Column(db.String(length=50), unique=True, nullable=False)
      user_type = db.Column(db.String(length=30), nullable=True)
      user_dept = db.Column(db.String(length=30), nullable=True)
      GDCSelect = db.Column(db.String(length=30), nullable=True)
      Cluster = db.Column(db.String(length=30), nullable=True)
      MngID = db.Column(db.String(length=30), nullable=True)

class UserDataEmail(db.Model, UserMixin):    
      __tablename__ = 'UserDataEmail'
      id = db.Column(db.Integer(), primary_key=True )
      username = db.Column(db.String(length=30), unique=True, nullable=False)
      email_add = db.Column(db.String(length=50), unique=True, nullable=False)
    

class TMStatusByAdmin(db.Model):
    id = db.Column(db.Integer(),primary_key=True)
    department = db.Column(db.String(length=30))
    level = db.Column(db.String(length=30))
    topic = db.Column(db.String(length=30))
    mailid = db.Column(db.String(length=30))
    score = db.Column(db.Integer(),nullable=False)
    remarks = db.Column(db.Integer())
    year = db.Column(db.Integer(),nullable=False)
    period = db.Column(db.String(length=50),nullable=False) 
    uploadby = db.Column(db.String(), db.ForeignKey('user_data_new.email_add'))
    

class UploadByAdmin_obj(db.Model):
    __tablename__ = 'UploadByAdmin_obj'
    id = db.Column(db.Integer(),primary_key=True)
    filename = db.Column(db.String(length=30))
    data = db.Column(db.LargeBinary)
    dept = db.Column(db.String(length=30),nullable=True)
    topic_added = db.Column(db.String(length=30), nullable=False )
    level = db.Column(db.String(length=30), nullable=False )
    year = db.Column(db.Integer(),nullable=False)
    period = db.Column(db.String(length=50),nullable=False) 
    uploadby = db.Column(db.Integer(), db.ForeignKey('user_data_new.id'))
    def __repr__(self):
        return f'<UploadByAdmin_obj {self.id}>'


class sup_topic_count(db.Model):
   __tablename__ = 'sup_topic_count'
   topic_ID = db.Column(db.Integer(), primary_key=True )
   topic_name = db.Column(db.String(length=30), nullable=False )
   Count = db.Column(db.Integer(), nullable=True )
   Dept = db.Column(db.String(length=30), nullable=True )
   year = db.Column(db.Integer(),nullable=False)
   period = db.Column(db.String(length=50),nullable=False)
   def __repr__(self):
        return f'<sup_topic_count {self.topic_ID}>'

class sup_nominated_count(db.Model):
   __tablename__ = 'sup_nominated_count'
   topic_ID = db.Column(db.Integer(), primary_key=True )
   topic_name = db.Column(db.String(length=30), nullable=False )
   Count = db.Column(db.Integer(), nullable=True,default='0' )
   IncompleteCount= db.Column(db.Integer(), nullable=True ,default='0')
   Dept = db.Column(db.String(length=30), nullable=True )
   year = db.Column(db.Integer(),nullable=False)
   period = db.Column(db.String(length=50),nullable=False)
   def __repr__(self):
        return f'<sup_nominated_count {self.topic_ID}>'

class AttendanceUploadByAdmin(db.Model):
    __tablename__ = 'AttendanceUploadByAdmin'
    id = db.Column(db.Integer(),primary_key=True)
    Subject = db.Column(db.String(), nullable=False)
    Dept = db.Column(db.String(), nullable=False)
    Email = db.Column(db.String(length=30),nullable=False)
    First_Join = db.Column(db.String(length=30),nullable=False)
    Last_Leave = db.Column(db.String(length=30),nullable=False)
    Duration = db.Column(db.String(length=30),nullable=False)
    Feedback = db.Column(db.String(length=30),nullable=True,default='Not Done')
    Period = db.Column(db.String(length=30),nullable=False)
    Year = db.Column(db.String(length=30),nullable=False)
    uploadby = db.Column(db.String(), db.ForeignKey('user_data_new.email_add'))
    def __repr__(self):
        return f'<AttendanceUploadByAdmin {self.id}>'
    

class SupResultStatusByAdmin(db.Model):
    __tablename__ = 'SupResultStatusByAdmin'
    id = db.Column(db.Integer(),primary_key=True)
    mailid = db.Column(db.String(length=30))
    topic = db.Column(db.String(length=30))
    level = db.Column(db.String(length=30))
    department = db.Column(db.String(length=30))
    marks = db.Column(db.Integer())
    status = db.Column(db.String(length=30))
    comment = db.Column(db.String(length=200)) 
    year = db.Column(db.Integer(),nullable=False)
    period = db.Column(db.String(length=50),nullable=False) 
    uploadby = db.Column(db.String(), db.ForeignKey('user_data_new.email_add'))


class SUP_topic(db.Model):
   __tablename__ = 'SUP_topic'
   topic_ID = db.Column(db.Integer(), primary_key=True )
   topics_added = db.Column(db.String(length=30), nullable=False )
   question_type = db.Column(db.String(length=30), nullable=False)
   Department = db.Column(db.String(length=30), nullable=False)
   Level = db.Column(db.String(length=30), nullable=False)
   Mandatory = db.Column(db.String(length=30), nullable=False)
   TimeInMin = db.Column(db.Integer(), nullable=False)
   PassMark = db.Column(db.Integer(), nullable=False)
   NoOfGDC = db.Column(db.String(length=100), nullable=False)
   SPSubTopic = db.Column(db.String(length=100), nullable=True)
   IsDeleted = db.Column(db.Text(), default="No")
   uploadby = db.Column(db.String(), db.ForeignKey('user_data_new.email_add'))
   def __repr__(self):
        return f'<SUP_topic {self.topic_ID}>'


class Nominate_topic_by(db.Model):
   topic_ID = db.Column(db.Integer(), primary_key=True )
   topics_added = db.Column(db.String(length=30), nullable=False )
   question_type = db.Column(db.String(length=30), nullable=False)
   Department = db.Column(db.String(length=30), nullable=False)
   nominate_topic = db.Column(db.Text())
   nominatedBY = db.Column(db.Integer())
   SUP_status = db.Column(db.Text(), default="Incomplete")
   TimeInMin = db.Column(db.Integer(), nullable=False)
   NominationDate = db.Column(db.String, nullable=False)
   Level = db.Column(db.String, nullable=False)
   year = db.Column(db.Integer(),nullable=False)
   AssessmentTime = db.Column(db.Integer(),nullable=False)
   period = db.Column(db.String(length=50),nullable=False) 
   StartTime = db.Column(db.String(),nullable=False) 
   StopTime = db.Column(db.String(),nullable=False) 
   uploadby = db.Column(db.String(), db.ForeignKey('user_data_new.email_add'))

class topic_question(db.Model):
   __tablename__ = 'topic_question' 
   q_ID = db.Column(db.Integer(), primary_key=True )
   question_type = db.Column(db.String(length=30), nullable=False)
   q_topic = db.Column(db.String(length=50), nullable=False)
   q_type = db.Column(db.String(length=50), nullable=False)
   choices_no = db.Column(db.Integer(), nullable=False)
   q_text = db.Column(db.String(length=1000), nullable=False)
   q_choice1 = db.Column(db.String(length=100), nullable=True)
   q_choice2 = db.Column(db.String(length=100), nullable=True)
   q_choice3 = db.Column(db.String(length=100), nullable=True)
   q_choice4 = db.Column(db.String(length=100), nullable=True)
   q_ans = db.Column(db.String(length=100), nullable=False)
   q_points = db.Column(db.Integer(), nullable=False)
   dept_Func = db.Column(db.String(length=100), nullable=False)
   level = db.Column(db.String(length=100), nullable=False)
   IsDeleted= db.Column(db.String(length=100), nullable=False, default="No")
   uploadby = db.Column(db.String(), db.ForeignKey('user_data_new.email_add'))
   def __repr__(self):
        return f'<topic_question {self.q_ID}>' 
   
class calendar_events(db.Model,UserMixin):
    eventid = db.Column(db.Integer, primary_key=True)
    title =db.Column(db.String, nullable=False)
    date =db.Column(db.String, nullable=False)
    dept =db.Column(db.String, nullable=False)
    Time = db.Column(db.String, nullable=False)
    Etime = db.Column(db.String, nullable=False)
    url = db.Column(db.String, nullable=True)
    level = db.Column(db.String, nullable=False)
    sme = db.Column(db.String, nullable=True)
    Mail = db.Column(db.String, nullable=True,default="NotSent")
    ReMail = db.Column(db.String, nullable=True,default="Disable")
    NoOfGDC = db.Column(db.String(length=100), nullable=False)
    SPSubTopic = db.Column(db.String(length=100), nullable=True)
    IsDeleted = db.Column(db.String, nullable=True,default="No")
    uploadby = db.Column(db.String(), db.ForeignKey('user_data_new.email_add'))
    
    def to_json(self):
        return {
            'eventid': self.eventid,
            'title': self.title,
            'date': self.date,
            # 'end': self.end,
            'url': self.url
        }
    
    def __repr__(self):
        return f'Item {self.eventid}'
    

class Attendance(db.Model):
    id = db.Column(db.Integer(),primary_key=True)
    user_name = db.Column(db.String(), db.ForeignKey('user_data_new.email_add'))
    user_dept = db.Column(db.String(), db.ForeignKey('user_data_new.user_dept'))
    location = db.Column(db.String(length=30),nullable=False)
    Topicname = db.Column(db.String(length=30),nullable=False)
    name = db.Column(db.String(length=30),nullable=False)
    confirm_attendance = db.Column(db.String(length=30),nullable=True)
    rate_1 = db.Column(db.String(length=30),nullable=True)
    rate_2 = db.Column(db.String(length=30),nullable=True)
    rate_3 = db.Column(db.String(length=30),nullable=True)
    rate_4 = db.Column(db.String(length=30),nullable=True)
    session_like = db.Column(db.String(length=30),nullable=True)
    session_better = db.Column(db.String(length=30),nullable=True)
    session_do_well = db.Column(db.String(length=30),nullable=True)
    session_even_better = db.Column(db.String(length=30),nullable=True)
    suggestion = db.Column(db.String(length=30),nullable=True)
    level = db.Column(db.String(length=30),nullable=False)
    FeedbackDate = db.Column(db.String, nullable=False)
    year = db.Column(db.Integer(),nullable=False)
    period = db.Column(db.String(length=50),nullable=False)

class IncidentTbl(db.Model):
    __tablename__ = 'IncidentTbl'
    id = db.Column(db.Integer(),primary_key=True)
    inc_id = db.Column(db.String(length=30),nullable=False)
    user_name = db.Column(db.String(), db.ForeignKey('user_data_new.email_add'))
    user_dept = db.Column(db.String(), db.ForeignKey('user_data_new.user_dept'))
    location = db.Column(db.String(length=30),nullable=False)
    issuetype = db.Column(db.String(length=30),nullable=False)
    modulename = db.Column(db.String(length=30),nullable=False)
    Subject = db.Column(db.String(length=30),nullable=False)
    Description = db.Column(db.String(length=30),nullable=False)
    filename = db.Column(db.String(length=30))
    Raiseddate =db.Column(db.String, nullable=False)
    Closuredate=db.Column(db.String, nullable=True)
    Status=db.Column(db.String, nullable=False,default='Open')
    Admincomment = db.Column(db.String, nullable=False)
    data = db.Column(db.LargeBinary)

class Upload(db.Model):
    id = db.Column(db.Integer(),primary_key=True)
    filename = db.Column(db.String(length=30))
    data = db.Column(db.LargeBinary)
    dept = db.Column(db.String(length=30),nullable=True)
    level = db.Column(db.String(length=30),nullable=False)
    topic_added = db.Column(db.String(length=30), nullable=False )
    year = db.Column(db.Integer(),nullable=False)
    period = db.Column(db.String(length=50),nullable=False) 
    uploadby = db.Column(db.Integer(), db.ForeignKey('user_data_new.id'))

class answerPerQnr(db.Model):
    id = db.Column(db.Integer(),primary_key=True)
    useraccount = db.Column(db.Integer(), db.ForeignKey('user_data_new.id'))
    useranswer = db.Column(db.String(length=50),nullable=True)
    correctanswer = db.Column(db.String(length=50),nullable=True)
    question_id = db.Column(db.String(length=50),nullable=True)
    session_id = db.Column(db.String(length=50),nullable=True)
    q_points = db.Column(db.String(length=50),nullable=True)
    isitdone = db.Column(db.String(length=50),nullable=True)
    level = db.Column(db.String(length=50),nullable=False)
    year = db.Column(db.Integer(),nullable=False)
    period = db.Column(db.String(length=50),nullable=False)
    qtype = db.Column(db.String(length=50),nullable=False)

class result_test(db.Model):
    __tablename__ = 'result_test'
    id = db.Column(db.Integer(),primary_key=True)
    session_id = db.Column(db.String(length=50),nullable=True)
    overall_score = db.Column(db.String(length=50),nullable=True)
    remarks = db.Column(db.String(length=50),nullable=True)
    NominatedBy  = db.Column(db.String(length=30), nullable=False) 
    level = db.Column(db.String(length=50),nullable=False)
    year = db.Column(db.Integer(),nullable=False)
    period = db.Column(db.String(length=50),nullable=False)
    PassMark = db.Column(db.Integer(),nullable=False)
    uploadby = db.Column(db.String(), db.ForeignKey('user_data_new.email_add'))


class training_topic_nomination(db.Model):
   topic_ID = db.Column(db.Integer(), primary_key=True )
   topics_added = db.Column(db.String(length=30), nullable=False )
   Dept = db.Column(db.String(length=30), nullable=True )
   nominate_topic = db.Column(db.Text())
   NominatedBy  = db.Column(db.String(length=30), nullable=False)  
   year = db.Column(db.Integer(),nullable=False)
   period = db.Column(db.String(length=50),nullable=False) 
   uploadby = db.Column(db.String(), db.ForeignKey('user_data_new.email_add'))


class training_topic_count(db.Model):
   topic_ID = db.Column(db.Integer(), primary_key=True )
   topic_name = db.Column(db.String(length=30), nullable=False )
   Count = db.Column(db.Integer(), nullable=True )
   Dept = db.Column(db.String(length=30), nullable=True )
   year = db.Column(db.Integer(),nullable=False)
   period = db.Column(db.String(length=50),nullable=False)
   ContactList = db.Column(db.String(),nullable=False)


class OJTNew(db.Model):
    ojtid = db.Column(db.Integer(),primary_key=True)
    TrainerEmail = db.Column(db.String(length=30),nullable=False)
    TraineeName = db.Column(db.String(length=30),nullable=False)
    TraineeEmail = db.Column(db.String(length=30),nullable=False)
    WeekNumber = db.Column(db.String(length=30),nullable=False)
    rate_1 = db.Column(db.String(length=30),nullable=True)
    rate_2 = db.Column(db.String(length=30),nullable=True)
    rate_3 = db.Column(db.String(length=30),nullable=True)
    Strengths = db.Column(db.String(length=30),nullable=True)
    Todo_better = db.Column(db.String(length=30),nullable=True)
    FeedbackLogDate = db.Column(db.String, nullable=False)
    uploadby = db.Column(db.Integer(), db.ForeignKey('user_data_new.id'))
    

class UploadByAdmin(db.Model):
    id = db.Column(db.Integer(),primary_key=True)
    filename = db.Column(db.String(length=30))
    data = db.Column(db.LargeBinary)
    dept = db.Column(db.String(length=30),nullable=True)
    topic_added = db.Column(db.String(length=30), nullable=False )
    level = db.Column(db.String(length=30), nullable=False )
    year = db.Column(db.Integer(),nullable=False)
    period = db.Column(db.String(length=50),nullable=False) 
    uploadby = db.Column(db.Integer(), db.ForeignKey('user_data_new.id'))


class TrainingRequest(db.Model):
    __tablename__ = 'TrainingRequest'
    id = db.Column(db.Integer(),primary_key=True)
    user_name = db.Column(db.String(), db.ForeignKey('user_data_new.email_add'))
    user_dept = db.Column(db.String(), db.ForeignKey('user_data_new.user_dept'))
    location = db.Column(db.String(length=30),nullable=False)
    Topicname = db.Column(db.String(length=30),nullable=False)
    Department = db.Column(db.String(length=30),nullable=False)
    PeopleCount = db.Column(db.Integer(),nullable=False)
    cluster = db.Column(db.String(length=30), nullable=False )
    Description = db.Column(db.String(length=30), nullable=True )
    SMEContact = db.Column(db.String(), nullable=True)
    def __repr__(self):
        return f'<TrainingRequest {self.id}>'






with app.app_context():
 db.create_all()
'''
#SQL DROP TABLE statement
table_name = "trainer" 
# Create an executable SQL statement
drop_table_statement = text(f"DROP TABLE IF EXISTS `{table_name}`")

# Execute the SQL DROP TABLE statement
with app.app_context():
    connection = db.session.connection()
    connection.execute(drop_table_statement)

# this will create your database table 
#with app.app_context():
 #db.create_all()
'''
    