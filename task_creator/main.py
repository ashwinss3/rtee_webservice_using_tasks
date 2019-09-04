from flask import Flask,request,jsonify
from flask_restful import Resource, Api
from datetime import datetime
import os
import base64
import json

from google.cloud import tasks_v2

def create_task_xlsx_mail(payload):
    
	client = tasks_v2.CloudTasksClient()
	project = os.environ.get('PROJECT_ID')
	location = os.environ.get('LOCATION')
	queue = os.environ.get('QUEUE_NAME')

	parent = client.queue_path(project, location, queue)
	#task_name = "xlsx-mail-task"

	#creating a task
	task = {
		#'name': client.task_path(project, location, queue, task_name), #cannot create 2 tasks with same name, within a certain period
		'app_engine_http_request': {
			'http_method': 'POST',
			'relative_uri': '/xlsx-mail',
			'app_engine_routing': {
				'service': 'xlsx-mail'
			},
			'headers': {
				'Content-Type': 'application/json'
			},
			'body': json.dumps(payload).encode()
		}
	}

	response = client.create_task(parent, task)
	
	print('Task Added')
	print(response)

class JsonToXlsxTask(Resource):
	def post(self):
		"""
		post request is served and task is created
		"""
		input_json = request.get_json()
		print("creating task")
		create_task_xlsx_mail(input_json)
		resp = jsonify(success=True)
		return resp


app = Flask(__name__)
api = Api(app)
api.add_resource(JsonToXlsxTask, '/')


if __name__ == '__main__':
    app.run(debug=True)
