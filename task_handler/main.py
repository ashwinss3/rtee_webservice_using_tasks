from flask import Flask,request,jsonify
from flask_restful import Resource, Api
import xlsxwriter
from datetime import datetime
import os
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (Mail,Attachment,FileContent,FileName,FileType)
import base64


def json_to_excel(input_json,filename):
	"""
	Function expects a json and filename as input and creates an xlsx file from the json.
	"""
	workbook = xlsxwriter.Workbook(filename) #creating workbook
	worksheet = workbook.add_worksheet() #addding a sheet to the workbook, use differnt sheets if needed
	bold = workbook.add_format({'bold': True}) #adding bold format for headers
	try:
		#writing header after taking keys, from initial dictionary. All dictionaries in the list should have same number keys.
		for index,header in enumerate(input_json[0].keys()):
			worksheet.write(0,index,header,bold) #writing to the 0th row

		#writing data from json
		for row,data in enumerate(input_json):	
			for col,value in enumerate(data.values()):
				worksheet.write(row+1,col,value)		#row+1 because header is already written in the initial row.

		workbook.close()

	except Exception as ex:
		print("Error while making xlsx file: {}".format(ex))


def get_attachment(file_path):
	"""
	Generates an attachment for sendgrid from the file in the path
	"""
	with open(file_path, 'rb') as f:
		data = f.read()

	encoded = base64.b64encode(data).decode()
	attachment = Attachment()
	attachment.file_content = FileContent(encoded)
	attachment.file_type = FileType('application/xlsx') #can be changed to take type from extension of file, or maybe not mandatory :p
	attachment.file_name = FileName(file_path.split("/")[-1])
	return attachment


def send_mail_with_attachment(email_from,email_to,email_subject,attachment_path,content="PFA"):
	"""
	Function to send an email, with given attachment
	"""
	message = Mail(
	from_email = email_from,
	to_emails = email_to,
	subject = email_subject,
	html_content = content)
	message.attachment = get_attachment(attachment_path)
	try:
		sg = SendGridAPIClient(os.environ.get('SENDGRID_API_KEY'))
		response = sg.send(message)
		print(response.status_code)
		print(response.body)
		print(response.headers)
	except Exception as e:
		print("Error while sending mail : {}".format(e))


class JsonToXlsx(Resource):
	def post(self):
		"""
		post request is served
		"""
		email_from = os.environ.get('EMAIL_SENDER')
		email_subject = "Open For Excel File"
		filename = "/tmp/request_{}.xlsx".format(datetime.now().strftime('%Y-%m-%dT%H:%M'))
		input_json = request.get_json()
		email_to = input_json['email']
		print("creating excel file")
		json_to_excel(input_json['body'],filename)
		print("Sending Email")
		send_mail_with_attachment(email_from,email_to,email_subject,filename)
		print("mail sent")
		resp = jsonify(success=True)
		return resp


app = Flask(__name__)
api = Api(app)
api.add_resource(JsonToXlsx, '/xlsx-mail')


if __name__ == '__main__':
    app.run(debug=True)
