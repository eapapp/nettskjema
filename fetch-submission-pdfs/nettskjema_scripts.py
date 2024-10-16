import requests as rq
import os
import ipywidgets as widgets
from IPython.display import display
import nettskjema_auth_client
import ndjson


apiURL = 'https://api.nettskjema.no/v3/form'
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

    if type(ref)==str: ref = int(ref)
    return(ref)


def getContactPerson(subID):
    url = '/'.join([apiURL, 'submission', str(subID)])
    headers.update({"accept": "application/json"})
    resp = rq.get(url=url, headers=headers).json()["answers"]
    contact = "_"
    for r in resp:
        if r["elementId"] in [5990407, 5990420, 5403627, 4170089, 4254977, 4254983]:
            contact = r["textAnswer"]
            if " " in contact: contact = contact.split(' ')[-1].strip()
            return(contact)

    # Search separately because of overlapping field labels for first name and full name
    for r in resp:
        if r["elementId"] in [4397996, 5990406]:
            contact = r["textAnswer"]
            if " " in contact: contact = contact.split(' ')[-1].strip()
            return(contact)

    return(contact)


def getSubmission(subID, formID):

    url = '/'.join([apiURL, 'submission', str(subID), 'pdf'])
    headers.update({"accept": "application/pdf"})
    resp = rq.get(url=url, headers=headers)

    contact = getContactPerson(subID)
    fname = '-'.join([prefixes[formID], str(subID), contact]) + '.pdf'    
    formdir = list(forms.keys())[list(forms.values()).index(formID)]        

    fpath = os.path.join(rootdir, formdir, fname)
    with open(fpath,'wb') as pf:
        pf.write(resp.content)

    return(formdir + "/" + fname)


def getSubmissions(formID, latest):

    url = '/'.join([apiURL, str(formID), 'submission-metadata'])
    headers.update({"accept": "application/x-ndjson"})
    resp = rq.get(url=url, headers=headers)
    if not resp.status_code == 200: print("Error:", resp.status_code, resp.reason)
    data = ndjson.loads(resp.text)
    subIDs = []
    for d in data:
        if latest:
            if d["submissionId"] > latest: subIDs.append(d["submissionId"])
        else:
            subIDs.append(d["submissionId"])

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
