import base64
import sys
import re
import pip
try:
    import requests
except ImportError:
    print 'requests module not installed.'
    print 'Installing requests...'
    pip.main(['install', 'requests'])
    import requests

try:
    import gspread
except ImportError:
    print 'gspread module not installed.'
    print 'Installing gspread...'
    pip.main(['install', 'gspread'])
    import gspread

import json
import time
from oauth2client.client import SignedJwtAssertionCredentials
from datetime import date, timedelta

def write_to_sheet(values, estimatedTime, noProjectTime, noPlannedTime):
    # Note: This is Andreas client email & private google API key. Plz no share
    client_email = 'account-1@toggltospreadsheet.iam.gserviceaccount.com'
    private_key = 'nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCV1KZiUni31qeq\nSAnwnN8PfAEiyQ0X9V9Qwye8HwRJQQ9J5FTfghWfFuAaGoGY86h/bTFpsEum3baG\n8peDU2u1h3SFHsJGr0crVjE66M5yAbPyzTza11vvrdnk/sKeDk2rwvjjmsSULI6q\nTJPeykcnKasPT+p2F/FPNkNsqupSJxENWWWak8Qi3Fx6kkFvhRJhjJeG/4wgjT02\nBeWyi1acjG4bh3uAUpnOPphO3Tbqn+6H6SUjssGDM8bx69UYs+mpgCUtGvV2Vh9U\n3FKuaFW5VwIEj6oetm9HpWdZPhMdTZ2ji8C1hFXD1PbetNLC9pWw2tIK39UN+0DO\nJIX4Yim5AgMBAAECggEADbG/+2CaVAM0k+W9stZDtQCagflqTvRIX/NWYGVdWOyv\nXwuVjlAvdbj7hqdBjR9IGU17pLNOVtQQ6MinfINVZuQYjEpBa4B9usJIAaBhrh3l\nup1zoHTqnD+BmOKIGzYndzl80iDFMt3XtE0HfudyFWxDjES7o0uQ7LKC+x2sCRIQ\n2Vt4PIUYCK10/eiHqV2MSdtKY3+6GvTelNSs7d8N0563B8CNdej/HU1/oEUzwXEo\n5mXZxUnnr9Jgt4NRhacaJR0E2NuxbQK7IYagLqufW3vK43FJUTRiBv8Xtl2miUWl\nGiktQDI8aVsFNqNUHhd1qrm3EZ6lYmGeyKVtzlp+AQKBgQDiZpxBe5WQRyE90dO7\nbraxJGeX1JqiE/RFYhjyWOebIsjnR4Suv/2dZR38rrWpWjRck5RqhsALRAuOCcTA\n6p++bNph/f3RlMICD+pqh/9rwuweqqETvaVZw/zsIfiftthQ7r9j/q8ebnSGVpHr\nybBovB5+TyHiJhmIgrdKMUlw2QKBgQCpa1HNBYzApY+gcEPiUkxizYf4JChULlGd\nlAV+DdLBxaSkYgEJ2MIuJQ/8ZOJm5swOxJp4UhXM/rSi0ObXnyL3AIIe56IK+kY9\nZbFMsxx/U7U8cyPObihH4tSOeJoeZdolI6YCON0Nz0obJ/PxmPYCd3RdgYuPifyC\nbvm8g3jz4QKBgGjpRaUugHsQCv5bmjLzteLWTM7VrSZH+tyf/ZFn00NXViOeR4S2\n4O4rqj6qMvIcI8F2fcLzWFCgIn6aVjtTPdz/Eh9wlEqnFVPhTi45gQnNlJ8NUIEW\nU2YKZMyDXXOdRhYS3EuY/EsswgByY0IQ/xc5fSPoxXnHT/OrJwZRWofZAoGAceTd\n9zCV8STcK4WNfWbKR1nY4K6eFgmVgJP0JUvxtabDCmeAPzhjQlZUKt8/fOIHqJ3v\nIpg8Y7WPhi1eIvKutNK4p0IdI7gg5EGrMd7vd4G0w1C8b5iKp9kMAEN/iJP8VR9k\nCPZlVVVXgm4XhwHH0Nyxc/MU+YhQIvesGFliRMECgYAhACR6D7rJ8VvBV8GKWWJr\nxholsBaE101nYbUQrZLRjBkOdw72PAhieHZivbkqf/ZjWyH1HmRdJ5X8P5a8THcp\n32ndHguve8MgAKOyYfkCfHhqnXijLSf4RfkaoCR8ki8Eq1diIi36DItVo6FvRzPs\nqmfEdiKscq6kTFzRBtMPpA=='
    
    # Change the path, otherwise the the json_key will not be found and writing to the Sheet will fail.
    json_key = json.load(open('TogglToSpreadSheet-97a0ed27dce5.json'))
    scope = ['https://spreadsheets.google.com/feeds']

    credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'].encode(), scope)
    gc = gspread.authorize(credentials)
    sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/1P8HFfktwSAF0rgi9rxpNvdMdR0GYuo845qAGikeJUh8/edit#gid=0')
    worksheet = sh.get_worksheet(1)

    currentDate = time.strftime("%Y-%m-%d")
    results = [currentDate]
    worksheet.insert_row(results, 2)

    worksheet.update_acell("C2", values)
    worksheet.update_acell("E2", noPlannedTime)
    worksheet.update_acell("G2", estimatedTime)
    return;

def get_totaltime_data():
    api_token = 'a000cbf7ec40079d37e3f4c795bec97c'
    _workspace_id = 1190663
    print 'Sending Request...'

    # Change/Rework next year/sprint
    sprintDay = date.today()
    sprintDay = sprintDay.replace(day=7, month=12,year=2015)

    #Project
    r = requests.get('https://toggl.com/reports/api/v2/summary', auth=(api_token, 'api_token'), params={'workspace_id': _workspace_id, 'since' : sprintDay, 'user_agent': 'api_test'})
    
    # No project
    r2 = requests.get('https://toggl.com/reports/api/v2/summary', auth=(api_token, 'api_token'), params={'workspace_id': _workspace_id, 'since' : sprintDay, 'project_ids' : '0', 'user_agent': 'api_test'})
    if r.status_code != 200:
        print 'Request Failed. Check your API Token'
        return;
    index = []

    rWorkspace = requests.get('https://www.toggl.com/api/v8/workspaces/1190663/projects', auth=(api_token, 'api_token'))
    workspaceText = rWorkspace.text
    taskMap = map(int, re.findall(r'\d+', workspaceText))
    estimatedTime = 0
    counter = 0
    index = [x-1 for x, i in enumerate(taskMap) if i == 1190663]
    

    for x in index:
        timeMap = []
        rProject = requests.get('https://www.toggl.com/reports/api/v2/project', auth=(api_token, 'api_token'), params={'user_agent': 'api_test','workspace_id': _workspace_id, 'project_id': taskMap[x]})
        if r.status_code != 200:
            print 'rProject request Failed. Check your API Token'
            return;
        projectText = rProject.text
        text = projectText.replace('null','0')
        finalText = text.replace('0.0', '0')
        timeMap = map(int, re.findall(r'\d+', finalText))
        if timeMap != []:
			estimatedTime = estimatedTime + timeMap[6]
        counter+= 1

    wholeText = r.text
    wholeText2 = r2.text
    allNumbers = map(int, re.findall('\d+', wholeText))
    allNumbers2 = map(int, re.findall('\d+', wholeText2))

    try:
        totalTime = allNumbers[0]
    except:
        totalTime = 0
        
    if totalTime < 0:
        totalTime = 1

    timeleft = totalTime
    hours = totalTime / 3600000
    estimatedTimeInHours = estimatedTime / 3600
    try:
        noProjectTime = allNumbers2[0] / 3600000
    except:
        noProjectTime = 0
    
    str = 'Total Time: ' + repr(hours)
    print str
    result = hours;

    # Not planned Time
    totalTime = allNumbers[1]
    timeleft = totalTime
    hours = totalTime / 3600000
    noPlannedTime = hours
    #result = []
    #result.append(hours);
    #result.append(min);
    #result.append(sec);
    result = result - noPlannedTime
    if result < 0:
        result = 0
    print 'Writing to Sheet...'
    write_to_sheet(result, estimatedTimeInHours, noProjectTime, noPlannedTime)
    return;

get_totaltime_data()
print 'Done!'
