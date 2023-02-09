from datetime import date, timedelta
import win32com.client
from win32com.client import Dispatch, constants
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
import json

f = open(r'.\keys.json')
crd = json.load(f)
f.close

def get_sharepoint_context_using_user():
 
    # Get sharepoint credentials
    sharepoint_url = 'https://lanecouncilofgovernments.sharepoint.com'

    # Initialize the client credentials
    user_credentials = UserCredential(crd['username']['email'], crd['password']['pwd'])

    # create client context object
    ctx = ClientContext(sharepoint_url).with_credentials(user_credentials)

    return ctx

def get_files(file_url):
    try:

        site_url = "https://lanecouncilofgovernments.sharepoint.com"
        credentials = ClientCredential(crd['username']['email'], crd['password']['pwd'])

        ctx = ClientContext(site_url).with_credentials(credentials)
        
        # file_url is the sharepoint url from which you need the list of files
        list_source = ctx.web.get_folder_by_server_relative_url(file_url)
        files = list_source.files
        ctx.load(files)
        ctx.execute_query()

        return files

    except Exception as e:
        print(e)

#get_files('https://lanecouncilofgovernments.sharepoint.com/:x:/r/sites/LCOGDiversityEquityBelonging/_layouts/15/Doc.aspx?sourcedoc=%7B95B76D95-363F-452E-A99D-77524AA7A68E%7D&file=Facilitator%20and%20Notetaker%20Rotation.xlsx&action=default&mobileredirect=true')

date2 = date.today() + timedelta(days=7)

const=win32com.client.constants
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = f'Reminder: DEIB meeting next Wednesday {str(date2)}'
newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
newMail.HTMLBody =  """
<HTML><BODY><p>Hello all,</p> 

<p>Our meeting facilitator and notetaker next week are Denise and Kate, and next month are Ellen and Heidi. Following the links, please find the folders for the <a href="https://lanecouncilofgovernments.sharepoint.com/:f:/r/sites/LCOGDiversityEquityBelonging/Shared%20Documents/Committee%20Meetings/Agendas/2023?csf=1&web=1&e=wBl76p">meeting agenda</a>, and <a href="https://lanecouncilofgovernments.sharepoint.com/:f:/r/sites/LCOGDiversityEquityBelonging/Shared%20Documents/Committee%20Meetings/Meeting%20Notes/2023?csf=1&web=1&e=03DAbf">meeting notes</a>.

Meeting facilitators will need to fill out the agenda details. Please note that we will have a hybrid meeting this time, and the in-person meeting room is McKenzie Room on PPB 4th floor.</p> 

<p>Have a great day!</p>

<p>Dongmei on behalf of DEIB Strategic Planning Subcommittee</p></BODY></HTML>
"""
# newMail.To = "LCOGDiversityEquityInclusionBelonging@lcog.org"
newMail.To = "dgw@mac.com"
# newMail.From = "LCOGDiversityEquityInclusionBelonging@lcog.org"

# newMail.Send()