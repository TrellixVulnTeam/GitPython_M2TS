from flask import Flask, request, jsonify
import pandas as pd

app=Flask(__name__)

@app.route("/upload", methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        data_xls = pd.read_excel(f)
        return data_xls.to_html()
    return '''
    <!doctype html>
    <title>Upload an excel file</title>
    <h1>Excel file upload (csv, tsv, csvz, tsvz only)</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Upload>
    </form>
    <h3>Click “Browse” to find the LegalServer report on your computer that you want to add case hyperlinks to. Note, the report should be an excel file and ‘fresh’ out of LegalServer, not edited, it should have the 2 extraneous rows at the top of the sheet before case data begins. Once you have identified this file, click ‘Hyperlink’ and you should shortly be given a prompt to either open the file directly or save the file to your computer. When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</h3>
    '''

@app.route("/export", methods=['GET'])
def export_records():
    return 

if __name__ == "__main__":
    app.run()