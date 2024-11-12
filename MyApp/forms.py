from flask_wtf import FlaskForm
from wtforms import SelectMultipleField,StringField, PasswordField, SubmitField,SelectField, RadioField,IntegerField
from wtforms.validators import Length, EqualTo, Email, DataRequired, ValidationError,NumberRange


class RegisterForm(FlaskForm):
        
    username = StringField(label='User Name:', validators=[Length(min=2, max=30), DataRequired()])
    email_add = StringField(label='Email Address:', validators=[Email(), DataRequired()])
    GDCSelect = SelectField(u"Select GDC Location.", choices=[('', 'Please select GDC'),('GDC-India', 'GDC-India'), ('GDC-Philippines', 'GDC-Philippines'), ('GDC-Colombia', 'GDC-Colombia'), ('GDC-Egypt', 'GDC-Egypt'), ('GDC-Bratislava', 'GDC-Bratislava'), ('DRC-Poland', 'DRC-Poland'), ('GRC-Czech', 'GRC-Czech')])
    Cluster = StringField(label='Please enter the custer name:', validators=[Length(min=2, max=50), DataRequired()])
    Dept = SelectField(u"Choose your Function", choices=[('', 'Please select'),('SP', 'SP'), ('DP', 'DP'),('PM', 'PM'),('CH', 'CH'),('CO', 'CO')])
    MngID = StringField(label='Enter Manager Email Address:', validators=[Email(), DataRequired()])
    submit = SubmitField(label='Create Account')

class UploadFile(FlaskForm):
    submit = SubmitField(label='Upload File')
    
class LoginForm(FlaskForm):
    username = StringField(label='User Name:', validators=[DataRequired()])
    password = PasswordField(label='Password:', validators=[DataRequired()])
    submit = SubmitField(label='Log in')

class Feedback_form(FlaskForm):
    name = StringField(label='Name of the trainer:', validators=[DataRequired()])
    email = StringField(label='Your email address:', validators=[Email(), DataRequired()])
    message = StringField(label='Feedback message:', validators=[DataRequired()])
    submit = SubmitField(label='Submit feedback')

class Training_Initiate(FlaskForm):
    To = StringField(label='To',validators=[DataRequired()])
    Cc = StringField(label='CC',validators=[DataRequired()])
    Subject = StringField(label='Subject',validators=[DataRequired()])
    Body = StringField(label='Description of Email',validators=[DataRequired()])
    submit = SubmitField(label='Send')

class PurchaseItemForm(FlaskForm):
    submit = SubmitField(label='Purchase Item')

class SellItemForm(FlaskForm):
    submit = SubmitField(label='Sell Item')

class DeleteForm(FlaskForm):
    submit = SubmitField(label='Delete Item')

class DeleteFormQNR(FlaskForm):
    submit = SubmitField(label='Delete Item')

class DeleteEvent(FlaskForm):
    submit = SubmitField(label='Delete')

class SubmitEvent(FlaskForm):
    submit = SubmitField(label='Submit')

#harish
class add_topic(FlaskForm):
    topic_added = StringField(label='Enter topic name', validators=[DataRequired()])
    question_types = SelectField(u"Choose test type", choices=[('', 'Please select'),('obj', 'Objective'), ('subj', 'Subjective')], validators=[DataRequired()])
    Department = SelectField(u"Choose Function", choices=[('', 'Please select'),('SP', 'SP'), ('DP', 'DP'), ('CO', 'CO'), ('CH', 'CH'), ('PM', 'PM')], validators=[DataRequired()])
    Level = SelectField(u"Choose Level", choices=[('', 'Please select'),('L1', 'Level1'), ('L2', 'Level2'), ('L3', 'Level3')], validators=[DataRequired()])
    Mandatory = SelectField(u"Is it mandatory", choices=[('', 'Please select'),('Yes', 'Yes'), ('No', 'No')], validators=[DataRequired()])
    TimeInMin = IntegerField(label='Enter time in minutes.', validators=[DataRequired(), NumberRange(min=1, max=1000)])
    PassMark = IntegerField(label='Enter Passing Marks', validators=[DataRequired(), NumberRange(min=1, max=100)])
    NoOfGDC = SelectMultipleField(u"Select GDCs", choices=[('All', 'All'), ('GDC-India', 'GDC-India'), ('GDC-Philippines', 'GDC-Philippines'), ('GDC-Colombia', 'GDC-Colombia'), ('GDC-Egypt', 'GDC-Egypt'), ('GDC-Bratislava', 'GDC-Bratislava'), ('DRC-Poland', 'DRC-Poland'), ('GRC-Czech', 'GRC-Czech')], validators=[DataRequired()])
    SPSubTopic = SelectField(u"Platform", choices=[('', 'Select One'),('IOM', 'IOM'), ('NIPO', 'NIPO')])
    submit = SubmitField(label='Submit')

#harish
class add_question(FlaskForm):
    question_types = StringField(label='Question type')
    q_topics = StringField(label='Topic')
    choices_nos = SelectField(u"Choose an option", choices=[('', 'Please select choice'),('2', '2 choices'), ('4', '4 choices')], validators=[DataRequired()])
    q_type = SelectField(u"Choose an option", choices=[('', 'Please select type'),('Single', 'Single'), ('Multi', 'Multi')], validators=[DataRequired()])
    q_texts = StringField(label='Question label', validators=[DataRequired()])
    q_choice1s = StringField(label='Option 1', validators=[DataRequired()])
    q_choice2s = StringField(label='Option 2', validators=[DataRequired()])
    q_choice3s = StringField(label='Option 3')
    q_choice4s = StringField(label='Option 4')
    level = StringField(label='Enter level', validators=[DataRequired()])
    q_pointss = IntegerField(label='Points', validators=[DataRequired()])
    dept_Funcs = StringField(label='Department')
    submit = SubmitField(label='Submit')


class getInfo(FlaskForm):
    submit = SubmitField(label='View')

class getAccess(FlaskForm):
    submit = SubmitField(label='Grant')

class CloseInc(FlaskForm):
    submit = SubmitField(label='Mark as Closed')

class SendMail(FlaskForm):
    submit = SubmitField(label='Generate List')

class EnableMail(FlaskForm):
    submit = SubmitField(label='Regenerate')

class RmAccess(FlaskForm):
    submit = SubmitField(label='Decline')

class TakeSup(FlaskForm):
    submit = SubmitField(label='Take Assessment')

class Update_add_question(FlaskForm):
    topic_added = StringField(label='Enter topic', validators=[DataRequired()])
    nominated_topic = StringField(label='Enter topic', validators=[DataRequired()])
    question_types = StringField(label='Enter topic', validators=[DataRequired()])
    submit = SubmitField(label='Nominate')

class EventForm_add(FlaskForm):
    Ctitle = StringField(label='Title',validators=[DataRequired()])
    CDate = StringField(label='Date',validators=[DataRequired()])
    CDept = StringField(label='Dept',validators=[DataRequired()])
    Curl = StringField(label='Url')
    CNoOfGDC = StringField(label='NoOfGDC',validators=[DataRequired()])
    submit = SubmitField(label='Add')
    


class AttendanceForm(FlaskForm):
    location = RadioField('Please choose your center/Location', choices=[
        ('Phillippines', 'Phillippines'),
        ('India', 'India'),
        ('Colombia', 'Colombia'),
        ('Egypt', 'Egypt'),
        ('Bratislava', 'Bratislava'),
        ('GRC', 'GRC'),
    ], validators=[DataRequired()])
    name = StringField(label="Enter the trainer's name:",validators=[DataRequired()])
    Topicname = StringField(label="Enter the topic name:",validators=[DataRequired()])
    level = StringField(label="Enter the level:",validators=[DataRequired()])
    confirm_attendance = RadioField('Attendance confirmation', choices=[
         ('Online', 'Online'),
        ('Offline', 'Offline')
    ], validators=[DataRequired()])
    rate_1 = RadioField('The Objectives of the training were clearly defined', choices=[
        ('5', '5 - Highest Rate'),
        ('4', '4'),
        ('3', '3'),
        ('2', '2'),
        ('1', '1 - Lowest Rating')
    ], validators=[DataRequired()])
    rate_2 = RadioField('Participation and interaction were encouraged', choices=[
        ('5', '5 - Highest Rate'),
        ('4', '4'),
        ('3', '3'),
        ('2', '2'),
        ('1', '1 - Lowest Rating')
    ], validators=[DataRequired()])
    rate_3 = RadioField('Delivery of the learning session', choices=[
        ('5', '5 - Highest Rate'),
        ('4', '4'),
        ('3', '3'),
        ('2', '2'),
        ('1', '1 - Lowest Rating')
    ], validators=[DataRequired()])
    rate_4 = RadioField('Did the trainer demonstrate the topic with relevant examples', choices=[
        ('5', '5 - Highest Rate'),
        ('4', '4'),
        ('3', '3'),
        ('2', '2'),
        ('1', '1 - Lowest Rating')
    ], validators=[DataRequired()])
    session_like = StringField(label="What did you like about the session ?",validators=[DataRequired()])
    session_better = StringField(label="What can make the session even better ?",validators=[DataRequired()])
    session_do_well = StringField(label="What did the SME do well in the session ?",validators=[DataRequired()])
    session_even_better = StringField(label="What could the SME do to make the session even better ?",validators=[DataRequired()])
    suggestion = StringField(label="Any Suggestion")
    
    submit = SubmitField(label='Submit')

class OJTForm(FlaskForm):
    TraineeName = StringField(label="Enter the trainee name:",validators=[DataRequired()])
    TraineeEmail = StringField(label="Enter the trainee email:",validators=[DataRequired()])
    WeekNumber = SelectField(u'Select Week Number',choices=[
        ('','Please select one'),
        ('1','Week 1'),
        ('2','Week 2'),
        ('3','Week 3'),
        ('4','Week 4')
    ],validators=[DataRequired()])
    rate_1 = RadioField('Scripting Skills', choices=[
        ('VeryGood', 'VeryGood'),
        ('Good', 'Good'),
        ('Average', 'Average'),
        ('Bad', 'Bad'),
        ('VeryBad', 'VeryBad'),
    ], validators=[DataRequired()])
    rate_2 = RadioField('Learning Mindset', choices=[
        ('VeryGood', 'VeryGood'),
        ('Good', 'Good'),
        ('Average', 'Average'),
        ('Bad', 'Bad'),
        ('VeryBad', 'VeryBad'),
    ], validators=[DataRequired()])
    rate_3 = RadioField('Team Management', choices=[
        ('VeryGood', 'VeryGood'),
        ('Good', 'Good'),
        ('Average', 'Average'),
        ('Bad', 'Bad'),
        ('VeryBad', 'VeryBad'),
    ], validators=[DataRequired()])

    
    Strengths = StringField(label="What are his/her strengths",validators=[DataRequired()])
    Todo_better = StringField(label="What can he/she do better ?",validators=[DataRequired()])
    submit = SubmitField(label='Submit')
    
class TrainingRequestForm(FlaskForm):
  
    Topicname = StringField(label="Enter the topic name:",validators=[DataRequired()])
    Department = RadioField('Please choose your function', choices=[
        ('SP', 'SP'),
        ('DP', 'DP'),
        ('PM', 'PM'),
        ('CH', 'CH'),
        ('CO', 'CO'),
    ], validators=[DataRequired()])
    PeopleCount = IntegerField(label="No of attendees for this topic:",validators=[DataRequired()])
    Description = StringField(label="Please elaborate the impact on this training:",validators=[DataRequired()])
    SMEContact = StringField(label="Do you know anyone who is aware on this topic:")
    submit = SubmitField(label='Submit')