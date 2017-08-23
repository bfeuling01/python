import requests, json, openpyxl, datetime
from os.path import expanduser

home = expanduser("~")

cpwb = openpyxl.load_workbook(home + '\Desktop\CrashPlanData.xlsx')
sh = cpwb.active

sh.title = 'Data Set'
sh.cell(row = 1, column = 1).value = 'Computer ID'
sh.cell(row = 1, column = 2).value = 'Device GUID'
sh.cell(row = 1, column = 3).value = 'User'
sh.cell(row = 1, column = 4).value = 'Computer'
sh.cell(row = 1, column = 5).value = 'Status'
sh.cell(row = 1, column = 6).value = 'OS'
sh.cell(row = 1, column = 7).value = 'Cold Storage'
sh.cell(row = 1, column = 8).value = 'Last Connection'
sh.cell(row = 1, column = 9).value = 'Last Completed Backup'
sh.cell(row = 1, column = 10).value = 'Last Activity'
sh.cell(row = 1, column = 11).value = 'Days Since Last Connection'
sh.cell(row = 1, column = 12).value = 'Days Since Last Completed Backup'
sh.cell(row = 1, column = 13).value = 'Days Since Last Activity'
sh.cell(row = 1, column = 14).value = 'Delinquent'

# get initial computer ID list
usrResp = requests.get(<crashplanurl>, auth=(<username>, <password>), headers={'Authorization': 'access_token <accesstoken>'})
usrFull = usrResp.json()['data']

devID = []
for u in usrFull:
    devID.append(json.dumps(u['deviceUid']))

r, c = 2, 1

for i in devID:
    devInfo = requests.get(<crashplanurl> + '%s' % str(i), auth=(<username>, <password>), headers={'Authorization': 'access_token <accesstoken>'})
    infoFull = devInfo.json()['data']
    compurl = <crashplanAPIURL> + "{0}?idType=guid".format(i)
    cUrl = compurl.replace('"', '')
    cInfo = requests.get(cUrl, auth=(<username>, <password>), headers={'Authorization': 'access_token <accesstoken>'})
    for i in infoFull:
        sep = 'T'
        lcd = str(i['lastConnectedDate']).split(sep, 1)[0]
        lcbd = str(i['lastCompletedBackupDate']).split(sep, 1)[0]
        la = str(i['lastActivity']).split(sep, 1)[0]
        sh.cell(row = r, column = c).value = cInfo.json()['data']['computerId']
        sh.cell(row = r, column = c + 1).value = i['deviceUid']
        sh.cell(row = r, column = c + 2).value = i['username']
        sh.cell(row = r, column = c + 3).value = i['deviceName']
        sh.cell(row = r, column = c + 4).value = i['status']
        sh.cell(row = r, column = c + 5).value = i['os']
        sh.cell(row = r, column = c + 6).value = i['coldStorage']
        sh.cell(row = r, column = c + 7).value = lcd
        sh.cell(row = r, column = c + 8).value = lcbd
        sh.cell(row = r, column = c + 9).value = la
        sh.cell(row = r, column = c + 10).value = """=IF(COUNTIF(H{0},"*none*"), "", DATEDIF(H{0},TODAY(),"d"))""".format(r)
        sh.cell(row = r, column = c + 11).value = """=IF(COUNTIF(I{0},"*none*"), "", DATEDIF(I{0},TODAY(),"d"))""".format(r)
        sh.cell(row = r, column = c + 12).value = """=IF(COUNTIF(J{0},"*none*"), "", DATEDIF(J{0},TODAY(),"d"))""".format(r)
        sh.cell(row = r, column = c + 13).value = """=IF(H{0} <> "", IF(M{0} > 30, "YES", "NO"), "NONE")""".format(r)
    r += 1

cpwb.save(home + '\Desktop\CrashPlanData.xlsx')
 
