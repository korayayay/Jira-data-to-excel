from jira.client import JIRA
import pandas as pd
import xlsxwriter

email = "your_jira_mail"
api_token = "project_api_key"
server = 'project_url'
jql = "project = 'project_task_key' ORDER BY created DESC"

jira = JIRA(options={'server': server}, basic_auth=(email, api_token))
jira_issues = jira.search_issues(jql, maxResults=0)
data = []


for issue in jira_issues:
    story_estimate_point = getattr(issue.fields, '#yourpointfieldid', 'N/A')
    
    comments = []
    for comment in issue.fields.comment.comments:
        comments.append(comment.body)
    
    d = {
        'key': issue.key,
        'assignee': str(issue.fields.assignee),
        'summary': str(issue.fields.summary),
        'description': str(issue.fields.description),
        'status': str(issue.fields.status.name),
        'status_description': str(issue.fields.status.description),
        'comments': "\n".join(comments),  
        'point': str(story_estimate_point)
    }
    data.append(d)

issues_df = pd.DataFrame(data)

workbook = xlsxwriter.Workbook('jira-excel.xlsx')
worksheet = workbook.add_worksheet('Issues')

header_format = workbook.add_format({'bold': True, 'bg_color': '#D8E4BC', 'align': 'center'})
for col_num, value in enumerate(issues_df.columns.values):
    worksheet.write(0, col_num, value, header_format)

for row_num, (index, row) in enumerate(issues_df.iterrows(), 1):
    for col_num, value in enumerate(row):
        worksheet.write(row_num, col_num, value)

workbook.close()

print("Excel dosyası başarıyla oluşturuldu.")