from docx import *
from docx.dml.color import ColorFormat
from docx.text.run import Font, Run

def parse_resume(loc):
	document = Document(loc)

	response = {
		'name': 'name',
		'email': 'email',
		'experience': [],
		'education': [],
		'courses': [],
		'projects': [],
		'certifications': [],
		'skills': [],
		'patents': [],
		'languages': [],
		'honors': [],
		'organizations': [],
		'other': 'other'
	}

	i = 0
	val = -1

	temp_exp = {
		'post': '',
		'company': '',
		'description': '',
		'start_date': '',
		'end_date': '',
		'time_period': ''
	}

	temp_pro = {
		'name': '',
		'start_date': '',
		'end_date': '',
		'description': ''
	}

	temp_pat = {
		'name': '',
		'id': '',
		'description': ''
	}


	temp_hon = {
		'title': '',
		'issuer': '',
		'description': ''
	}

	for para in document.paragraphs:
		for run in para.runs:
			if run.font.size.pt == 16.0:
				if val == 1:
					response['experience'].append(temp_exp)
				elif val == 7:
					response['projects'].append(temp_pro)
				elif val == 8:
					response['honors'].append(temp_hon)
				elif val == 10:
					response['patents'].append(temp_pat)

				if run.text == 'Experience':
					val = 1
					temp_exp = {
						'post': '',
						'company': '',
						'description': '',
						'start_date': '',
						'end_date': '',
						'time_period': ''
					}

					i = 1
					continue
				elif run.text == 'Certifications':
					val = 2
					continue
				elif run.text == 'Education':
					val = 3
					continue
				elif run.text == 'Skills & Expertise':
					val = 4
					continue
				elif run.text == 'Courses':
					val = 5
					continue
				elif run.text == 'Languages':
					val = 6
					continue
				elif run.text == 'Projects':
					temp_pro = {
						'name': '',
						'start_date': '',
						'end_date': '',
						'description': ''
					}

					val = 7
					i = 1
					continue
				elif run.text == 'Honors and Awards':
					temp_hon = {
						'title': '',
						'issuer': '',
						'description': ''
					}

					val = 8
					i = 1
					continue
				elif run.text == 'Organizations':
					val = 9
					continue
				elif run.text == 'Patents':
					temp_pat = {
						'name': '',
						'id': '',
						'description': ''
					}

					val = 10
					i = 1
					continue	
				elif run.text == 'Other':
					val = -1
					continue
			
			if val == 1:
				if run.bold:
					if i == 5:
						response['experience'].append(temp_exp)

						temp_exp = {
							'post': '',
							'company': '',
							'description': '',
							'start_date': '',
							'end_date': '',
							'time_period': ''
						}
						i = 1

					if i == 1:
						temp_exp['post'] = run.text[:-3]
						i += 1
					elif i == 2 and (not run.text == ''):
						temp_exp['company'] = run.text
						i += 1
				elif i == 3:
					temp_exp['start_date'] = run.text.split(' - ')[0]
					temp_exp['end_date'] = run.text.split(' - ')[1]
					i += 1
				elif i == 4:
					temp_exp['time_period'] = run.text[1:-1]
					i += 1
				elif i == 5:
					temp_exp['description'] += run.text
			
			elif val == 2:
				if run.bold:
					response['certifications'].append(run.text) # will read only 1 value because rest is in a table
			
			elif val == 3:
				if run.bold:
					response['education'].append(run.text)

			elif val == 4:
				if run.bold:
					response['skills'].append(run.text)
			
			elif val == 5:
				if run.font.size.pt > 10.0:
					response['courses'].append(run.text)

			elif val == 6:
				if run.bold:
					response['languages'].append(run.text)

			elif val == 7:
				if run.bold:
					if i == 4:
						response['projects'].append(temp_pro)

						temp_pro = {
							'name': '',
							'start_date': '',
							'end_date': '',
							'description': ''
						}
						i = 1

					if i == 1:
						temp_pro['name'] = run.text
						i += 1
				elif i == 2:
					#read newline as space. difference of ms word and libreoffice
					temp_pro['start_date'] = run.text.split(' to ')[0]
					temp_pro['end_date'] = run.text.split(' to ')[1]
					i += 1
				elif i == 3:
					i += 1
				elif i == 4 and run.font.size.pt > 10.0:
					temp_pro['description'] += run.text
			
			elif val == 8:
				if run.bold:
					if i == 3:
						response['honors'].append(temp_hon)

						temp_hon = {
							'title': '',
							'issuer': '',
							'description': ''
						}
						i = 1

					if i == 1:
						temp_hon['title'] = run.text
						i += 1
				elif i == 2:
					temp_hon['issuer'] = run.text
					i += 1
				elif i == 3:
					temp_hon['description'] += run.text

			elif val == 9:
				if run.bold:
					response['organizations'].append(run.text)

			elif val == 10:
				if run.bold:
					if i == 3:
						response['patents'].append(temp_pat)
						temp_pat = {
							'name': '',
							'id': '',
							'description': ''
						}

						i = 1
					if i == 1:
						temp_pat['name'] = run.text
						i += 1
				elif i == 2:
					temp_pat['id'] = run.text
					i += 1
				elif i == 3:
					temp_pat['description'] += run.text

			elif run.bold:
				if run.font.size.pt == 20.0:
					response['name'] = run.text

			elif run.font.size.pt == 13.0:
				response['email'] = run.text


	response['courses'] = response['courses'][2:]

	return response


print parse_resume('/home/samkit/Downloads/SamkitJainProfile (copy).docx')