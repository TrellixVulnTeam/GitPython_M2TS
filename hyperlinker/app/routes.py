from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_file
from app import app, db
from flask_login import current_user, login_user, logout_user, login_required
from app.models import User, Post
from app.forms import LoginForm, RegistrationForm, EditProfileForm, PostForm, ResetPasswordForm
from werkzeug.urls import url_parse
from datetime import datetime
from app.forms import ResetPasswordRequestForm
from app.email import send_password_reset_email
import pandas as pd

@app.route('/', methods=['GET', 'POST'])
@app.route('/index', methods=['GET', 'POST'])
def index():
    return '''
    <!doctype html>
    <title>Python Tools</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>LSNYC Python Tools</h1>
    
    <h3>Tools for manipulating reports from LegalServer:</h3>
    
    <b>General Purpose:</b>
    <ul type="disc">
    <li><a href="/Hyperlinker">Hyperlinker</a></li>
    <li><a href="/BoroughSplitter">Borough Splitter</a></li>

    <li><a href="/AccessLineSplitter">Access Line Splitter</a></li>
    
    
    <li><a href="/RadSearcher">Identify Clients in RAD Buildings</a></li>
    <li><a href="/AnnualComparer">Compare Spreadsheets Between Years</a></li>
    
    </ul>
    
    <b>Housing Tools:</b>
    <ul type="disc">
    
    <li><a href="/AllHousing">All Housing Report</a></li>
    <li><a href="/DHCIPrep">DHCI Document Extractor Prep</a></li>
    <li><a href="/WaiverMaker">Prepare Waiver Request Templates</a></li>
    </br>
    <li><a href="/UAHPLPCovidClean">[Revised for Covid] UAHPLP Cleaner</a></li>     
    <li><a href="/UAHPLPExternalPrepCovid">[Revised for Covid] Prepare UA/HPLP Reports for Submission to HRA</a></li>
    <li><a href="/UAHPLPConsolidatedBoroughCleaner">UAHPLP Consolidated Borough Cleaner</a></li>
    </br>
    <li><a href="/TRCCovidClean">[Revised for Covid] TRC Cleaner</a></li>
    <li><a href="/TRCExternalPrepCovid">[Revised for Covid] Prepare TRC Reports for Submission to HRA</a></li>
    <li><a href="/TRCtallyCovid">[Revised for Covid] TRC Tally</a></li>
    <li><a href="/TRCDuplicateIdentifier">TRC Identify Duplicates</a></li>
    </br>
    <li><a href="/MLSIntakeConsolidatedHousingCleaner">MLS Consolidated Intake Cleaner</a></li>
    </br>
    <li><a href="/AllHousingSimpler">All Housing Report Tester</a></li>
    
    
    </ul>

    
    <b>IOI:</b>
    <ul type="disc">
    
    <li><a href="/IOIOnePreparer"> IOI 1 Preparer</a></li>
    <br>
    
    <li><a href="/IOIimmQuarterly">Quarterly IOI Immigration</a></li>
    <li><a href="/IOIimmTally"> IOI Immigration Tally</a></li>
    <br>
    
    <li><a href="/IOIempQuarterly">Quarterly IOI Employment</a></li>
    <li><a href="/IOIempTally">IOI Employment Tally</a></li>
    <br>
    
    <li><a href="/IOIWaiverMaker">Prepare Waiver Request Templates</a></li>
    </ul>
    
    <b>Pascale Tools:</b>
    <ul type="disc">
    <li><a href="/DAPException">DAP Exception & Report Prep Tool</a></li>
    <li><a href="/LSCCovidPrep">LSC Covid Report Prep</a></li>
    </ul>
    
    <b>IOLA Tools:</b>
    <ul type="disc">
    <li><a href="/IOLADollarCleaner">IOLA $ Cleaner</a></li>
    </ul>
    <b>Foreclosure Tools:</b>
    <ul type="disc">
    <li><a href="/CNYCNCleaner">CNYCN Cleaner</a></li>
    </ul>
    
    <b>Compliance Tools:</b>
    <ul type="disc">
    <li><a href="/ComplianceConsolidater">Compliance Consolidater</a></li>
    <li><a href="/ComplianceConsolidaterYearEnd">Year End Compliance Consolidater</a></li>
    <li><a href="/CSRDeterminer">Determine CSR Eligibility</a></li>
    
    
    </ul>
    
    <br>
    <b>Development Sandbox:</b>
    <ul type="disc">
    <li><a href="/BoxTester">Checkbox Tester</a></li>
    <li><a href="/PTOEstimator">PTO Estimator</a></li>
    <li><a href="/Sandbox1">Sandbox1</a></li>
    
    </ul>
    <br>
    <p>If you have any questions or concerns about these tools (or ideas for new ones!) please reach out to <a href="mailto:jackley@lsnyc.org?cc=rstump@lsnyc.org&Subject=Hi Jay! (Idea for Python tool!)" target="_top">Jay Ackley</a></p>
    
    '''
    #<li><a href="/TRCtally">TRC Tally</a></li>
    #<li><a href="/BrownsvilleTRCclean">Brownsville TRC Cleaner</a></li>
    #<li><a href="/TRCExternalPrep">Prepare TRC Reports for Submission to HRA</a></li>
    #<li><a href="/AdvocateSplitter">Advocate Splitter</a></li>
    #<li><a href="/IOIempMonthly">Monthly IOI Employment</a></li>
    #<li><a href="/IOIimmMonthly">Monthly IOI Immigration</a></li>
    #<li><a href="/MLSIntakeConsolidatedHousingCleaner">[MLS] Consolidated Housing Cleaner</a></li>
    #<li><a href="/BkLSIntakeConsolidatedHousingCleaner">[BkLS] Consolidated Housing Cleaner</a></li>
    #<li><a href="/TRCtally">TRC Tally</a></li>
    #<li><a href="/AccessLineSplitter">Identify Access Line Cases</a></li>