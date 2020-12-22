import pickle
import os.path
import tkinter as tk
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from random import sample, choice, randint, choices
import urllib.request
import re

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
ID = open('ID').readline()

sheet = ""
n = 0
words = []
cur = -1
weights = []
dif = []
start = 1

def colorCell(xb, yb, xe, ye, r, g, b):
	global sheet
	requests = []
	requests.append({
		'repeatCell' : {
			'range' : {
				'sheetId' : 0,
				'startRowIndex' : xb,
				'endRowIndex' : xe,
				'startColumnIndex' : yb,
				'endColumnIndex' : ye
			},
			'cell' : {
			'userEnteredFormat' : {
					'backgroundColor' : {
						'red' : r,
						'green' : g,
						'blue' : b
					}
				}
			},
			'fields': 'userEnteredFormat(backgroundColor)'

		}
	})
	body = {
    	'requests': requests
	}
	response = sheet.batchUpdate(spreadsheetId=ID, body=body).execute()

def showWord(textArea, text):
	textArea['state'] = 'normal'
	textArea.delete(1.0, 'end')
	textArea.insert(1.0, text)
	textArea['state'] = 'disabled'


def run():
	global sheet, words, n, weights, dif
	creds = None
	# The file token.pickle stores the user's access and refresh tokens, and is
	# created automatically when the authorization flow completes for the first
	# time.
	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			creds = pickle.load(token)
	# If there are no (valid) credentials available, let the user log in.
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				'credentials.json', SCOPES)
			creds = flow.run_local_server(port=0)
		# Save the credentials for the next run
		with open('token.pickle', 'wb') as token:
			pickle.dump(creds, token)
	service = build('sheets', 'v4', credentials=creds)
	sheet = service.spreadsheets()
	result = sheet.values().get(spreadsheetId=ID,
								range='B1:B1').execute()
	n = int(result.get('values', [])[0][0])
	result = sheet.values().get(spreadsheetId=ID,
								range=f'A2:D{1 + n}').execute()
	words = result.get('values', [])
	weights = [start] * n	
	dif = [0] * n	


def main(w, v):
	global sheet, n, words, cur, weights, dif

	if v == "":
		c = n
	else:
		c = int(v)
	if cur != -1:
		if w == -1:
			weights[cur] //= 2
		if w == 1:
			weights[cur] *= 2
#		dif[cur] += w
	print(len(weights) - weights.count(0))
	cur = choices(range(0, c), weights[:c])[0]
	showWord(wordText, words[cur][0])
	showWord(translateText, '')
	showWord(sourceText, '')
	showWord(frequencyText, words[cur][2])
	showWord(restText, c - weights.count(0))
	


	#	print(dir(sheet))
	# The file token.pickle stores the user's access and refresh tokens, and is
	# created automatically when the authorization flow completes for the first
	# time.
	
	"""pool = set((i[1] for i in elements))
	numpool = {}
	for i in elements:
		numpool[i[1]] = i[0]
	s = sample(elements, 10)
	fp, sp = s[:5], s[5:]
	fp = { 'values' : fp }
	sp = { 'values' : sp }
	result = sheet.values().update(spreadsheetId=ID,
								range='Игрок 1!A1', body=fp, valueInputOption='RAW').execute()
	result = sheet.values().update(spreadsheetId=ID,
								range='Игрок 2!A1', body=sp, valueInputOption='RAW').execute()
	result = sheet.values().update(spreadsheetId=ID,
								range='Аргументы!C2', body={'values' : [['']] * 45}, valueInputOption='RAW').execute()
	colorCell(1, 1, 50, 50, 1, 1, 1)
	"""

def add():
	global sheet, words
	addTk = tk.Toplevel(root)
	textFieldDict = {'height' : 1, 'width' : 25, 'pady' : 5, 'state' : 'normal'}

	wordText = tk.Text(addTk, **textFieldDict)
	wordText.grid(row=3, column=0)

	translateText = tk.Text(addTk, **textFieldDict)
	translateText.grid(row=3, column=1)

	sourceText = tk.Text(addTk, **textFieldDict)
	sourceText.grid(row=3, column=2)


	addButton = tk.Button(addTk, text='add', command=lambda : addWord([wordText.get(1.0, 'end'),  translateText.get(1.0, 'end'), frequency(wordText.get(1.0, 'end'))]))
	addButton.grid(row=4, column=1)
	 
def delete():
	global sheet, words, n
	words.pop(cur)
	dif.pop(cur)
	weights.pop(cur)
	result = sheet.values().update(spreadsheetId=ID,
								range='A2', body={'values' : words}, valueInputOption='RAW').execute()
	result = sheet.values().update(spreadsheetId=ID,
								range='B1', body={'values' : [[n - 1]]}, valueInputOption='RAW').execute()
	n -= 1

def frequency(word):
    word = word.strip()
    webUrl = urllib.request.urlopen(f'https://books.google.com/ngrams/graph?year_end=2019&year_start=1950&content={word}_INF')
    text = webUrl.read().decode('utf-8')
    pattern = re.compile('ngrams.data = \[.*?\[(.*)\];')
    p = pattern.findall(text)[0]
    frRow = list(map(float, (e[0] for e in re.findall(r'(\d+\.?\d+(e-)?\d+)', p))))
    average = sum(frRow) / len(frRow) * 10000000
    return average

def addWord(data):
	global sheet, words, n
	words.insert(0, data)
	weights.insert(0, 1)
	result = sheet.values().update(spreadsheetId=ID,
								range='A2', body={'values' : words}, valueInputOption='RAW').execute()
	result = sheet.values().update(spreadsheetId=ID,
								range='B1', body={'values' : [[n + 1]]}, valueInputOption='RAW').execute()
	n += 1

def translate():
	global sheet, n, words, cur
	showWord(translateText, words[cur][1])
	showWord(sourceText, words[cur][2])

def clear():
	global weights, dif
	weights = [start] * n
	dif = [0] * n


if __name__ == '__main__':
	root = tk.Tk()
	runButton = tk.Button(root, text='run', command=run)
	runButton.grid(row=0, column=1)

	showWordRightButton = tk.Button(root, text='right', command=lambda : main(-1, strNumberWordsEntry.get()), width=10)
	showWordRightButton.grid(row=1, column=0)

	showWordButton = tk.Button(root, text='show word', command=lambda : main(0, strNumberWordsEntry.get()), width=10)
	showWordButton.grid(row=1, column=1)


	showWordWrongButton = tk.Button(root, text='wrong', command=lambda : main(1, strNumberWordsEntry.get()), width=10)
	showWordWrongButton.grid(row=1, column=2)

	clearButton = tk.Button(root, text='clear', command=clear)
	clearButton.grid(row=1, column=3)


	translateWordButton = tk.Button(root, text='transalte word', command=translate)
	translateWordButton.grid(row=2, column=1)
	
	textFieldDict = {'height' : 1, 'width' : 25, 'pady' : 5, 'state' : 'disabled'}

	wordText = tk.Text(root, **textFieldDict)
	wordText.grid(row=3, column=0)

	translateText = tk.Text(root, **textFieldDict)
	translateText.grid(row=3, column=1)

	sourceText = tk.Text(root, **textFieldDict)
	sourceText.grid(row=3, column=2)

	frequencyText = tk.Text(root, **textFieldDict)
	frequencyText.grid(row=5, column=0)

	restText = tk.Text(root, **textFieldDict)
	restText.grid(row=5, column=2)


	addButton = tk.Button(root, text='add', command=add)
	addButton.grid(row=4, column=1)
	
	delButton = tk.Button(root, text='delete', command=delete)
	delButton.grid(row=5, column=1)



	strNumberWordsEntry = tk.StringVar()
	numberWordsEntry = tk.Entry(root, width=10, textvariable=strNumberWordsEntry)
	numberWordsEntry.grid(row=5, column=3)

	root.mainloop()