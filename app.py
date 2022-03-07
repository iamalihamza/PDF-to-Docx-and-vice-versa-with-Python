from flask import Flask, request
from flask_restful import Api, Resource
import os
from pathlib import Path
from pdf2docx import parse
import sys
import subprocess
import re
import logging
logging.basicConfig(level=logging.INFO)

app = Flask("__name__")
api = Api(app)

# Create a directory in a known location to save files to.
uploads_dir = os.path.join(app.instance_path, 'pdf_uploads')
converted_dir = os.path.join(app.instance_path, 'converted')
os.makedirs(uploads_dir, exist_ok=True)
os.makedirs(converted_dir, exist_ok=True)


def convert_to(folder, source, timeout=None):  # To convert docx to PDF
	args = [libreoffice_exec(), '--headless', '--convert-to', 'pdf', '--outdir', folder, source]
	process = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
	filename = re.search('-> (.*?) using filter', process.stdout.decode())

	if filename is None:
		raise LibreOfficeError(process.stdout.decode())
	else:
		return filename.group(1)


def libreoffice_exec():
	# TODO: Provide support for more platforms
	if sys.platform == 'darwin':
		return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
	return 'libreoffice'


class LibreOfficeError(Exception):
	def __init__(self, output):
		self.output = output


class Converter(Resource):
	def post(self):
		try:
			file_to_be_converted = request.files['file']
			file_extension = Path(file_to_be_converted.filename).suffix
			converted_path = None

			if file_extension == ".pdf":
				file_to_be_converted.save(os.path.join(uploads_dir, file_to_be_converted.filename))
				saved_file_name = file_to_be_converted.filename.split(".pdf")[0]
				word_file = f"{converted_dir}/{saved_file_name}.docx"
				converted_path = word_file
				parse(f"/{uploads_dir}/{file_to_be_converted.filename}", word_file, start=0, end=None)
				# deleting the user uploaded pdf file after converting
				os.remove(os.path.join(uploads_dir, file_to_be_converted.filename))

			elif file_extension == ".docx":
				file_to_be_converted.save(os.path.join('', file_to_be_converted.filename))
				file_path = f"{file_to_be_converted.filename}"
				result_path = f"{converted_dir}/"
				if os.path.isfile(file_path):
					converted_path = convert_to(result_path, file_path)
					# deleting the user uploaded docx file after converting
					os.remove(file_to_be_converted.filename)

			else:
				logging.info('Error: Invalid file extension')
				return {"data": "Invalid file extension"}

			return {"data": "File Converted successfully", 'file_path': converted_path}

		except Exception as e:
			logging.info(f'Error: {e}')
			return {"data": "Something went wrong, Please try again"}


api.add_resource(Converter, "/convert")

if __name__ == "__main__":
	app.run(debug=True)  # set debug to False before moving to production
