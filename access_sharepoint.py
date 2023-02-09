import json

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

f = open(r'.\keys.json')
crd = json.load(f)
f.close

url = 'https://lanecouncilofgovernments.sharepoint.com/:x:/r/sites/is/Shared%20Documents/Backup_Report_13-Jul-2020.xlsx?d=w8a3e7250f0554be69ffa745bbd1137e9&csf=1&web=1&e=Man5PN'
username, password = crd['username']['email'], crd['password']['pwd']
relative_url = '/sites/documentsite/Documents/filename.xlsx'

ctx_auth = AuthenticationContext(url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(url, ctx_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print("Web title: {0}".format(web.properties['Title']))

else:
    print(ctx_auth.get_last_error())
