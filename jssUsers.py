import requests, json, openpyxl, datetime
from os.path import expanduser

home = expanduser("~")

cpwb = openpyxl.load_workbook(home + '\Desktop\jssUsers.xlsx')
sh = cpwb.active

sh.title = 'Data Set'
sh.cell(row = 1, column = 3).value = 'User'
sh.cell(row = 1, column = 4).value = 'Computer'

# get initial computer ID list
cResp = requests.get(<jssURL>, auth = (<username>, <password>), headers={'Accept': 'application/json'})
comp = cResp.json()['computers']

compID = []
for c in comp:
	compID.append(json.dumps(c['id']))

r, c = 2, 1

for i in compID:
	iResp = requests.get(<jssURL> + i, auth = (<username>, <password>), headers={'Accept': 'application/json'})
	infoFull = iResp.json()['computer']
	sh.cell(row = r, column = c).value = infoFull['location']['username']
	sh.cell(row = r, column = c + 1).value = infoFull['general']['name']
	r += 1

cpwb.save(home + '\Desktop\jssUsers.xlsx')
