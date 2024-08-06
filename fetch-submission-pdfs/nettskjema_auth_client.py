import os
from dotenv import load_dotenv
from oauthlib.oauth2 import BackendApplicationClient
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth

token_url = 'https://authorization.nettskjema.no/oauth2/token'


def get_oauth2_token():

    load_dotenv()
    client_id = os.getenv('client_id')
    client_secret = os.getenv('client_secret')
    auth = HTTPBasicAuth(client_id, client_secret)
    session = OAuth2Session(client = BackendApplicationClient(client_id = client_id)) 

    try:
        token = session.fetch_token(token_url=token_url, auth=auth)

    except:
        raise Exception('Failed to fetch access token')

    return token['access_token']


if __name__=='__main__':
    token = get_oauth2_token()
    print('Token: Bearer', token)