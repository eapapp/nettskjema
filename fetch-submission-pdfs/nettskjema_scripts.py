import requests as rq
import os
import ipywidgets as widgets
from IPython.display import display
import nettskjema_auth_client


apiURL = 'https://nettskjema.no/api/v2'
headers = {}
forms = {
    "EBRAINS curation request form": 386195,
    "OLD EBRAINS curation request form": 277393,
    "EBRAINS Ethics Compliance Survey": 224765,
    "Request for version addition": 271758
}
prefixes = {
    386195: 'CR',
    277393: 'CR',
    224765: 'EC',
    271758: 'VA'
}
checkboxes = []
rootdir = ""


def getLatestSubmission(formdir):
    fpath = os.path.join(rootdir, formdir)
    files = os.listdir(fpath)

    ref = None
    if files:
        ref = 0
        for f in files:
            if not f[0] == '.':
                ref = f[3:11] if int(f[3:11]) > int(ref) else ref

    return(ref)


def getContactPerson(subID):
    url = '/'.join([apiURL, 'submissions', str(subID)])
    resp = rq.get(url=url, headers=headers).json()["answers"]
    contact = "_"
    for r in resp:
        if r["questionId"] in [6632775, 5936128, 4609360, 4509048]:
            contact = r["textAnswer"]
            if " " in contact: contact = contact.split(' ')[-1].strip()
            return(contact)

    # Search separately because of overlapping field labels for first name and full name
    for r in resp:
        if r["questionId"] in [6632774, 4769485]:
            contact = r["textAnswer"]
            if " " in contact: contact = contact.split(' ')[-1].strip()
            return(contact)

    return(contact)


def getSubmission(subID, formID):

    url = '/'.join([apiURL, 'submissions', str(subID), 'pdf'])
    resp = rq.get(url=url, headers=headers)

    contact = getContactPerson(subID)
    fname = '-'.join([prefixes[formID], str(subID), contact]) + '.pdf'    
    formdir = list(forms.keys())[list(forms.values()).index(formID)]        

    fpath = os.path.join(rootdir, formdir, fname)
    with open(fpath,'wb') as pf:
        pf.write(resp.content)

    return(formdir + "/" + fname)


def getSubmissions(formID, latest):

    url = '/'.join([apiURL, 'forms', str(formID), 'submissions?fields=submissionId'])
    if latest: url = url + '&fromSubmissionId=' + str(latest)

    resp = rq.get(url=url, headers=headers).json()
    subIDs = []
    for r in resp:
        subIDs.append(r["submissionId"])

    return(subIDs)


def selection_changed(sel):
    global checkboxes
    if sel['new']==True:
        for c in checkboxes:
            c.value=True
    else:
        for c in checkboxes:
            c.value=False


def init():   
    global rootdir
    rootdir = os.getcwd()
    if not rootdir.endswith('Nettskjema'):
        rootdir = rootdir[:rootdir.find('Nettskjema') + len('Nettskjema')]
    
    global headers
    token = nettskjema_auth_client.get_oauth2_token()
    headers = {"Authorization": "Bearer " + token}

    display(widgets.HTML(value='<h3>Please select the Nettskjema forms you would like to fetch the latest submissions from:</h3>'))

    selectall = widgets.Checkbox(
        value=False,
        description="Select all / clear selection",
        disabled=False,
        indent=False)
    
    selectall.observe(selection_changed, names='value')

    global checkboxes
    for f in forms.keys():
        w = widgets.Checkbox(
            value=False,
            description=f,
            disabled=False,
            indent=False)
        checkboxes.append(w)
        display(w)

    display(selectall)


def fetch():
    for c in checkboxes:
        if c.value == True:
            formID = forms[c.description]
            latest = getLatestSubmission(c.description)
            submissionIDs = getSubmissions(formID, latest)
            if submissionIDs:
                display(widgets.HTML(value='<h3>Downloading latest form submissions - ' + c.description + ':<h3>'))
                for subID in submissionIDs:
                    fname = getSubmission(subID, formID)
                    display(widgets.HTML(value=' - ' + fname))           
                display(widgets.HTML(value='<hr>'))

    display(widgets.HTML(value='Done.'))
