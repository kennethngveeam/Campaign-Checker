#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from flask import Flask, request, render_template
import pandas as pd

app = Flask(__name__)

def find_non_matching_emails(listupload_path, sfdc_path):
    try:
        listupload = pd.read_excel(listupload_path)
        sfdc = pd.read_excel(sfdc_path, skiprows=10)
        sfdc = sfdc.dropna(subset=['Member Email'])
        emails_in_listupload = listupload['Email Address'].str.lower().dropna().unique()
        emails_in_sfdc = sfdc['Member Email'].str.lower().dropna().unique()
        
        non_matching_details = listupload[~listupload['Email Address'].str.lower().isin(emails_in_sfdc)]
        non_matching_details = non_matching_details[['First Name', 'Last Name', 'Email Address', 'Company Name']]
        # Remove rows where all specified columns are NaN
        non_matching_details = non_matching_details.dropna(how='all', subset=['Email Address'])
        return non_matching_details.to_dict(orient='records')
    except Exception as e:
        print(f"Error processing files: {e}")
        return []

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        try:
            listupload = request.files['listupload']
            sfdc = request.files['sfdc']
            listupload.save('listupload.xlsx')
            sfdc.save('sfdc.xlsx')
            non_matching_details = find_non_matching_emails('listupload.xlsx', 'sfdc.xlsx')
            count_non_matching = len(non_matching_details)
            return render_template('result.html', details=non_matching_details, count=count_non_matching)
        except Exception as e:
            print(f"Error in upload route: {e}")
            return "An error occurred during file upload and processing."
    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=False)


# In[ ]:





# In[ ]:





# In[ ]:




