import requests
import json
import sys
import pandas
import openpyxl
from bs4 import BeautifulSoup
from jira import JIRA
from datetime import datetime
from collections import defaultdict
from StyleFrame import StyleFrame, Styler, utils

#jira credentials
CONST_USERNAME =
CONST_PASSWORD =

#observation period
startdate =
enddate =
start_date = datetime.strptime(startdate, '%Y-%m-%dT%H:%M:%S.%f+0530')
end_date = datetime.strptime(enddate, '%Y-%m-%dT%H:%M:%S.%f+0530')

#total issues
total_issue =

#local jira link
post_url = "http:"
req_url = "http:"

payload = {
    'os_username': CONST_USERNAME,
    'os_password': CONST_PASSWORD
}

headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'en-US,en;q=0.9,bn;q=0.8',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
    'Cookie': 'doc-sidebar=300px; JIRASESSIONID=907FE168EA0F5C3AD44C209E666AA8D4; atlassian.xsrf.token=BDBU-3VFZ-JU4K-K1FN|d25f179131bea0b1ba40fcc2402571aba3ac71b8|lin',
    'Host': '107.109.112.116:8080',
    'Upgraide-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36'
}

with requests.Session() as session:
    post = session.post(post_url, data=payload)
    req = session.get(req_url)
    if(req.status_code != 200):
        sys.exit("Oops, JIRA not reachable:(")
    else:
        print("jira server working!!")

options = {
    'server': 'http://107.109.112.116:8080',
    'verify': False
}

jira = JIRA(options, basic_auth=(CONST_USERNAME, CONST_PASSWORD))

new_dic = {}
new_dic['username1'] = []
new_dic['username2'] = []
new_dic['username3'] = []
new_dic['username4'] = []
new_dic['username5'] = []
new_dic['username6'] = []
new_dic['username7'] = []
new_dic['username8'] = []

for i in new_dic:
    for j in range(0, 3):
        new_dic[i].append(0)

for i in range (1, total_issue + 1):
    if i == 6:
        continue
    print('\n' + 'Analyzing S19I2P-' + str(i) + '...')
    issue = jira.issue('S19I2P-' + str(i))
    if issue:
        valid_issue = False
        print ('Issue created ' + issue.raw['fields']['created'])
        print ('Issue updated last ' + issue.raw['fields']['updated'])
        create_date = datetime.strptime(issue.raw['fields']['created'], '%Y-%m-%dT%H:%M:%S.%f+0530')
        update_date = datetime.strptime(issue.raw['fields']['updated'], '%Y-%m-%dT%H:%M:%S.%f+0530')
        if 'assignee' in issue.raw['fields'] and not (issue.raw['fields']['assignee'] is None):
            if 'name' in issue.raw['fields']['assignee'] and not (issue.raw['fields']['assignee']['name'] is None):
                print ('Assignee=' + issue.raw['fields']['assignee']['name'])
                assignee = issue.raw['fields']['assignee']['name']
                valid_issue = True
        
        if valid_issue == False:
            print ('Issue not assigned to anyone!')
            continue

        if 'status' in issue.raw['fields'] and not (issue.raw['fields']['status'] is None):
            if 'statusCategory' in issue.raw['fields']['status'] and not (issue.raw['fields']['status']['statusCategory'] is None):
                if 'name' in issue.raw['fields']['status']['statusCategory'] and not (issue.raw['fields']['status']['statusCategory']['name'] is None):
                    print ('Current status=' + issue.raw['fields']['status']['statusCategory']['name'])
                    state = issue.raw['fields']['status']['statusCategory']['name']
                    if state != 'Done':
                        new_dic[assignee][2] += 1
                    if (create_date > start_date and create_date < end_date):
                        print ('This issue created in this time')
                        new_dic[assignee][0] += 1
                    if (update_date > start_date) and (update_date < end_date) and (state == 'Done'):
                        print ('This issue completed in this time')
                        new_dic[assignee][1] += 1

df = pandas.DataFrame(new_dic)
df.index = ['Tasks Created', 'Tasks Completed', 'Tasks pending']
print(df.T)

#export output to excel sheet
writer = StyleFrame.ExcelWriter("taskreport.xlsx")
sf=StyleFrame(df.T)
sf.apply_column_style(cols_to_style=df.T.columns, styler_obj=Styler(bg_color=utils.colors.white, bold=True, font=utils.fonts.arial,font_size=8),style_header=True)
sf.apply_headers_style(styler_obj=Styler(bg_color=utils.colors.blue, bold=True, font_size=8, font_color=utils.colors.white,number_format=utils.number_formats.general, protection=False))
sf.to_excel(writer, sheet_name='Sheet1', index=True)
writer.save()
