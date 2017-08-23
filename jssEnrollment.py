import requests, json, datetime
from os.path import expanduser

home = expanduser("~")

lWeek = datetime.datetime.today() - datetime.timedelta(days = 7)
last = lWeek.strftime('%m/%d/%Y')

cResp = requests.get(<jssURL>, auth = (<username>, <password>), headers={'Accept': 'application/json'})
comp = cResp.json()['computers']

compID = []
for c in comp:
    compID.append(json.dumps(c['id']))

compCheck = []
for c in compID:
    iResp = requests.get(<jssURL> + '{0}'.format(c), auth = (<username>, <password>), headers = {'Accept': 'application/json'})
    info = iResp.json()['computer']
    date = datetime.datetime.strptime(info['general']['initial_entry_date'], '%Y-%m-%d')
    if date >= lWeek:
        compCheck.append('System: ' + json.dumps(info['general']['name']))
        compCheck.append('UserID: ' + json.dumps(info['location']['username']))
        compCheck.append('User Name: ' + json.dumps(info['location']['real_name']))
        compCheck.append('Enrolled: ' + json.dumps(info['general']['initial_entry_date']))

MSG = """ Here is a list of computers that have been enrolled in the last week:

%s
"""" % ('\n'.join(compCheck), last)

f = open(home + "\Desktop\enrollment.txt", "w")
f.write(MSG)
f.close() 
