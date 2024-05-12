from flask import Flask, render_template,request
from spacy.lang.en.stop_words import STOP_WORDS
from string import punctuation
from collections import Counter
from heapq import nlargest
import json
import spacy
import pandas as pd
from pypdf import PdfReader
import textract
import docx2txt
import os
from werkzeug.utils import secure_filename
from extract import *


model = spacy.load('en_core_web_sm')

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/extract', methods=['POST', 'GET'])
def extract():
    filename = summary = keywords = file_type = ''  # Initialize variables including file_type
    if request.method == 'POST':
        if 'file' in request.files and request.files['file'].filename != '':
            file = request.files['file']
            filename = secure_filename(file.filename)
            file_path = os.path.join('temp_save', filename)
            print(f'uploaded file: {filename}')
            if 'pdf' in filename:
                file_type = 'PDF'
            elif 'docx' in filename:
                file_type = 'DOCX'
            elif 'pptx' in filename:
                file_type = 'POWERPOINT'
            else:
                file_type = 'file processed'
            file.save(file_path)
            filename, summary, keywords, json_result = summarize(model, file_path)
            # print(summary)
            # print(keywords)
        else:
            print("No file part")
    return render_template('extract.html', title=filename, description=summary, keywords=keywords, file_type=file_type)

if __name__ == '__main__':
    app.run(debug=True)