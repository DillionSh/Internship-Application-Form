from flask import Flask, render_template, request, flash, send_from_directory, url_for, redirect,session
from dotenv import load_dotenv
import os
from simple_salesforce import Salesforce, SalesforceMalformedRequest
from flask_mail import Mail, Message
import datetime
from utils.salesforceHandling import SalesforceHandler

from flask_bootstrap import Bootstrap5
import base64
from difflib import SequenceMatcher
import re
import pandas as pd



app = Flask(__name__)
bootstrap = Bootstrap5(app)
app.secret_key = "random_secret_key"

app.config['MAIL_SERVER'] = 'smtp.office365.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USE_SSL'] = False
app.config['MAIL_USERNAME'] = os.getenv('MAIL_USERNAME')
app.config['MAIL_PASSWORD'] = os.getenv('MAIL_PASSWORD')
MAX_FILE_SIZE = 10 * 1024 * 1024  # 16 MB in bytes

mail = Mail(app)

load_dotenv()
sf_username = os.getenv('SF_USERNAME')
sf_password = os.getenv('SF_PASSWORD')
sf_security_token = os.getenv('SF_SECURITY_TOKEN')

# Salesforce login
sf = Salesforce(
    username=os.getenv('SF_USERNAME'),
    password=os.getenv('SF_PASSWORD'),
    security_token=os.getenv('SF_SECURITY_TOKEN'),
    domain='test'
)


sf = SalesforceHandler().sf
gender_values = SalesforceHandler().getPicklistValues('Contact', 'Gender__c')
citizenship_values = SalesforceHandler().getPicklistValues('Contact', 'Citizenship__c')
employment_values = SalesforceHandler().getPicklistValues('Contact', 'Employment_Status__c')
languages_values = SalesforceHandler().getPicklistValues('Contact', 'Language_Spoken__c')


def normalize_name(name):
    # Convert to lowercase
    name = name.lower()
    # Remove extra spaces
    name = re.sub(r'\s+', ' ', name).strip()
    return name


# Function to check if names are similar
def are_names_similar(name1, name2, threshold=1):
    if pd.isna(name1) or pd.isna(name2):
        return True
    name1_normalized = normalize_name(name1)
    name2_normalized = normalize_name(name2)

    # Check if one name contains the other
    if name1_normalized in name2_normalized or name2_normalized in name1_normalized:
        return True

    # Use SequenceMatcher to get a similarity ratio
    similarity_ratio = SequenceMatcher(None, name1_normalized, name2_normalized).ratio()

    return similarity_ratio >= threshold



    if conditions:
        query += " OR ".join(conditions)

    result = sf.query(query)

    # If the query returns any records, a record with the same first name, last name, or email exists
    return len(result['records']) > 0



internship_object_api_name = 'Contact'
gender_field_api_name = 'Gender__c'



def render_template_with_context(template_name, **context):
    """Render a template with common context data."""

    context.setdefault('gender_values', gender_values)
    context.setdefault('citizenship_values', citizenship_values)
    context.setdefault('employment_values', employment_values)
    context.setdefault('language_values', languages_values)






    return render_template(template_name, **context)



@app.route("/")
def index():
    form_data = session.pop('form_data', {})

    return render_template_with_context("index.html", form_data=form_data)


@app.route('/submit-form', methods=['POST'])
def submit_form():
    form_data = request.form.to_dict()

    # Collect form data
    name = request.form.get('name')
    nric = request.form.get('nric')

    dob = request.form.get('dob')
    gender = request.form.get('gender')
    citizenship = request.form.get('citizenship')

    # Convert language output to string
    language = request.form.getlist('language')
    language_str = ', '.join(language)

    employment = request.form.get('employment')
    name_of_school = request.form.get('name_of_school')
    home_no = request.form.get('home_no')
    mobile_no = request.form.get('mobile_no')
    email = request.form.get('email')
    block_no = request.form.get('block_no')
    street_name = request.form.get('street_name')
    unit_no = request.form.get('unit_no')
    postal_code = request.form.get('postal_code')
    father_name = request.form.get('father_name')
    father_occupation = request.form.get('father_occupation')
    father_employer = request.form.get('father_employer')
    mother_name = request.form.get('mother_name')
    mother_occupation = request.form.get('mother_occupation')
    mother_employer = request.form.get('mother_employer')
    start_date = request.form.get('start_date')
    end_date = request.form.get('end_date')
    relations = request.form.get('relations')
    criminal = request.form.get('criminal')
    dismiss = request.form.get('dismiss')
    illness = request.form.get('illness')
    others = request.form.get('others')
    file = request.files.get('fileUpload')

    try:
        if language_str == '':
            flash("Please choose at least one language spoken")
            session['form_data'] = form_data
            return redirect(url_for('index'))


        existing_contact_email = sf.query(f"SELECT Id FROM Contact WHERE Email = '{email}'")
        existing_contact_nric = sf.query(f"SELECT Id FROM Contact WHERE Client_NRIC__c = '{nric}'")

        #

        # if existing_contact_email['totalSize'] > 0:
        #     flash("A contact with the same email already exists. Please use a different email.")
        #     session['form_data'] = form_data
        #     return redirect(url_for('index'))











        if file:

            file_content = file.read()
            file_size = len(file_content)
            file.seek(0)
            if file_size > MAX_FILE_SIZE:
                flash("File size exceeds 10 MB limit. Please upload a smaller file.")
                return redirect(url_for('index'))

            elif existing_contact_nric['totalSize'] > 0:
                query_name = "SELECT Id, Name, Name__c FROM Internship__c"
                query_name_2 = "SELECT Id, Name, Client_NRIC__c FROM Contact"

                query_name_result = sf.query(query_name)
                query_name_result_2 = sf.query(query_name_2)

                best_match_contact_id = None
                best_match_internship_id = None

                if 'records' in query_name_result and 'records' in query_name_result_2:
                    internship_records = query_name_result['records']
                    contact_records = query_name_result_2['records']

                    for contact in contact_records:
                        contact_id = contact.get("Id")
                        contact_name = contact.get("Name")
                        contact_nric = contact.get("Client_NRIC__c")



                        print(f"Contact ID: {contact_id}, Contact NRIC: {contact_nric}")



                        if nric == contact_nric:
                            best_match_contact_id = contact_id
                            best_match_contact_name = contact_name
                            # Stop after finding the best match (if required, adjust this logic)
                            print(best_match_contact_name)
                            print(best_match_contact_id)
                    if best_match_contact_id:
                        for internship in internship_records:
                            internship_id = internship.get('Id')
                            internship_name_c = internship.get('Name__c')
                            print(internship_id)
                            print(internship_name_c)

                            if internship_name_c == best_match_contact_id:
                                best_match_internship_id = internship_id
                if best_match_internship_id:

                        update_data_contact = {

                            'DOB__c': dob,

                            'Gender__c': gender,
                            'Client_NRIC__c': nric,
                            'Mobile_Number__c': mobile_no,
                            'Citizenship__c': citizenship,
                            'Email': email,
                            'Language_Spoken__c': language_str,
                            'Employment_Status__c': employment,
                            'Block_Number__c': block_no,
                            'Street_Name__c': street_name,
                            'Unit_Number__c': unit_no,
                            'Postal_Code__c': postal_code,

                        }

                        result_contact = sf.Contact.update(best_match_contact_id, update_data_contact)

                        update_data_intern = {

                            'School_Name__c': name_of_school,

                            'Preferred_date_to_start_internshp__c': start_date,
                            'Preferred_date_to_end_internship__c': end_date,
                            'Date_of_Application__c': datetime.date.today().isoformat(),
                            'Criminal_Record__c': criminal,
                            'Relatives__c': relations,
                            'Dismissed__c': dismiss,
                            'Illness__c': illness,
                            'Details__c': others


                        }

                        # Update the record in Salesforce using the Internship record ID
                        result_intern = sf.Internship__c.update(best_match_internship_id, update_data_intern)

                        file_base64 = base64.b64encode(file_content).decode('utf-8')
                        content_version = sf.ContentVersion.create({
                            'Title': file.filename,
                            'PathOnClient': file.filename,
                            'VersionData': file_base64
                        })
                        content_version_id = content_version['id']
                        content_document_id = \
                            sf.query(f"SELECT ContentDocumentId FROM ContentVersion WHERE Id = '{content_version_id}'")[
                                'records'][0]['ContentDocumentId']
                        sf.ContentDocumentLink.create({
                            'ContentDocumentId': content_document_id,
                            'LinkedEntityId': best_match_internship_id,
                            'ShareType': 'I'
                        })
                        ##############################################################################

                        dob_datetime = datetime.datetime.strptime(dob, '%Y-%m-%d')

                        dob_formatted = dob_datetime.strftime('%d/%m/%Y')
                        start_date_formatted = datetime.datetime.strptime(start_date, '%Y-%m-%d').strftime('%d/%m/%Y')
                        end_date_formatted = datetime.datetime.strptime(end_date, '%Y-%m-%d').strftime('%d/%m/%Y')

                        # Send email
                        msg1 = Message('New Internship Form Submission', sender='dillion_intern@brahmcentre.com',
                                       recipients=['g59643072@gmail.com'])
                        msg1.body = f"""
                                            Name: {name}
                                            Date of Birth: {dob_formatted}
                                            Gender: {gender}
                                            Email: {email}
                                            Mobile Number: {mobile_no}
                                            Preferred Start Date: {start_date_formatted}
                                            Preferred End Date: {end_date_formatted}

                                            Salesforce record: "https://brahmcentre.lightning.force.com/lightning/r/Contact/{best_match_internship_id}/view"

                                            """
                        msg2 = Message('Successful Internship Form Submission', sender='dillion_intern@brahmcentre.com',
                                       recipients=[email])
                        msg2.body = f"""
                                                  Thank you for applying! You will get a email or phone call soon.
                                                   """

                        file_content = file.read()
                        msg1.attach(file.filename, file.content_type, file_content)

                        mail.send(msg1)
                        mail.send(msg2)

                        return render_template("success.html")

                elif not best_match_internship_id:

                    update_data_contact = {

                        'DOB__c': dob,

                        'Gender__c': gender,
                        'Client_NRIC__c': nric,
                        'Mobile_Number__c': mobile_no,
                        'Citizenship__c': citizenship,
                        'Email': email,
                        'Language_Spoken__c': language_str,
                        'Employment_Status__c': employment,
                        'Block_Number__c': block_no,
                        'Street_Name__c': street_name,
                        'Unit_Number__c': unit_no,
                        'Postal_Code__c': postal_code,

                    }

                    result_contact = sf.Contact.update(best_match_contact_id, update_data_contact)

                    new_member = best_match_contact_id

                    record_2 = sf.Internship__c.create({
                        'Name__c': new_member,

                        'School_Name__c': name_of_school,

                        'Preferred_date_to_start_internshp__c': start_date,
                        'Preferred_date_to_end_internship__c': end_date,
                        'Date_of_Application__c': datetime.date.today().isoformat(),
                        'Criminal_Record__c': criminal,
                        'Relatives__c': relations,
                        'Dismissed__c': dismiss,
                        'Illness__c': illness,
                        'Details__c': others

                    })
                    new_member_internship = record_2['id']

                    file_base64 = base64.b64encode(file_content).decode('utf-8')
                    content_version = sf.ContentVersion.create({
                        'Title': file.filename,
                        'PathOnClient': file.filename,
                        'VersionData': file_base64
                    })
                    content_version_id = content_version['id']
                    content_document_id = \
                        sf.query(f"SELECT ContentDocumentId FROM ContentVersion WHERE Id = '{content_version_id}'")[
                            'records'][0]['ContentDocumentId']
                    sf.ContentDocumentLink.create({
                        'ContentDocumentId': content_document_id,
                        'LinkedEntityId': new_member_internship,
                        'ShareType': 'I'
                    })
                    ##############################################################################

                    dob_datetime = datetime.datetime.strptime(dob, '%Y-%m-%d')

                    dob_formatted = dob_datetime.strftime('%d/%m/%Y')
                    start_date_formatted = datetime.datetime.strptime(start_date, '%Y-%m-%d').strftime('%d/%m/%Y')
                    end_date_formatted = datetime.datetime.strptime(end_date, '%Y-%m-%d').strftime('%d/%m/%Y')

                    # Send email
                    msg1 = Message('New Internship Form Submission', sender='dillion_intern@brahmcentre.com',
                                   recipients=['g59643072@gmail.com'])
                    msg1.body = f"""
                                Name: {name}
                                Date of Birth: {dob_formatted}
                                Gender: {gender}
                                Email: {email}
                                Mobile Number: {mobile_no}
                                Preferred Start Date: {start_date_formatted}
                                Preferred End Date: {end_date_formatted}

                                Salesforce record: "https://brahmcentre.lightning.force.com/lightning/r/Contact/{best_match_internship_id}/view"

                                """
                    msg2 = Message('Successful Internship Form Submission', sender='dillion_intern@brahmcentre.com',
                                   recipients=[email])
                    msg2.body = f"""
                                  Thank you for applying! You will get a email or phone call soon.
                                   """

                    file_content = file.read()
                    msg1.attach(file.filename, file.content_type, file_content)

                    mail.send(msg1)
                    mail.send(msg2)

                    return render_template("success.html")


            else:

                record = sf.Contact.create({
                    'LastName': name,
                    'DOB__c': dob,

                    'Gender__c': gender,
                    'Client_NRIC__c': nric,
                    'Mobile_Number__c': mobile_no,
                    'Citizenship__c': citizenship,
                    'Email': email,
                    'Language_Spoken__c': language_str,
                    'Employment_Status__c': employment,
                    'Block_Number__c': block_no,
                    'Street_Name__c': street_name,
                    'Unit_Number__c': unit_no,
                    'Postal_Code__c': postal_code,

                })
                new_member = record['id']
                record_2 = sf.Internship__c.create({
                    'Name__c': new_member,

                    'School_Name__c': name_of_school,

                    'Preferred_date_to_start_internshp__c': start_date,
                    'Preferred_date_to_end_internship__c': end_date,
                    'Date_of_Application__c': datetime.date.today().isoformat(),
                    'Criminal_Record__c': criminal,
                    'Relatives__c': relations,
                    'Dismissed__c': dismiss,
                    'Illness__c': illness,
                    'Details__c': others

                })
                new_member_internship = record_2['id']



                ########### Code to upload files to Salesforce internship object ############

                file_base64 = base64.b64encode(file_content).decode('utf-8')
                content_version = sf.ContentVersion.create({
                    'Title': file.filename,
                    'PathOnClient': file.filename,
                    'VersionData': file_base64
                })
                content_version_id = content_version['id']
                content_document_id = \
                    sf.query(f"SELECT ContentDocumentId FROM ContentVersion WHERE Id = '{content_version_id}'")[
                        'records'][0]['ContentDocumentId']
                sf.ContentDocumentLink.create({
                    'ContentDocumentId': content_document_id,
                    'LinkedEntityId': new_member_internship,
                    'ShareType': 'I'
                })
                ##############################################################################




                dob_datetime = datetime.datetime.strptime(dob, '%Y-%m-%d')

                dob_formatted = dob_datetime.strftime('%d/%m/%Y')
                start_date_formatted = datetime.datetime.strptime(start_date, '%Y-%m-%d').strftime('%d/%m/%Y')
                end_date_formatted = datetime.datetime.strptime(end_date, '%Y-%m-%d').strftime('%d/%m/%Y')

                # Send email
                msg1 = Message('New Internship Form Submission', sender='dillion_intern@brahmcentre.com',
                               recipients=['g59643072@gmail.com'])
                msg1.body = f"""
                     Name: {name}
                     Date of Birth: {dob_formatted}
                     Gender: {gender}
                     Email: {email}
                     Mobile Number: {mobile_no}
                     Preferred Start Date: {start_date_formatted}
                     Preferred End Date: {end_date_formatted}

                     Salesforce record: "https://brahmcentre.lightning.force.com/lightning/r/Contact/{record['id']}/view"

                     """
                msg2 = Message('Successful Internship Form Submission', sender='dillion_intern@brahmcentre.com',
                               recipients=[email])
                msg2.body = f"""
                           Thank you for applying! You will get a email or phone call soon.
                            """

                file_content = file.read()
                msg1.attach(file.filename, file.content_type, file_content)

                mail.send(msg1)
                mail.send(msg2)

                return render_template("success.html")

    except SalesforceMalformedRequest as e:

        error_content = e.content
        for error_dict in error_content:
            error_message = error_dict.get('message')
            if error_message:
                flash(f"{error_message}")
                session['form_data'] = form_data

                return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)
