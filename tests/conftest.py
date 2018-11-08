from pytest import fixture
from pathlib import Path
import google_sheets_lib
import os
import json


@fixture(scope='function')
def gs_client(pytestconfig) -> 'google_sheets_lib.GoogleSheets':
    client_secret_path = Path.cwd() / 'client_secret.json'
    with open(client_secret_path, 'w') as secret_file:
        secret_file.write(json.dumps({'installed': {
            'client_id': os.environ['OAUTH_TESTING_CLIENT_ID'],
            'client_secret': os.environ['OAUTH_TESTING_CLIENT_SECRET'],
            'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
            'token_uri': 'https://accounts.google.com/o/oauth2/token',
            'redirect_uris': ['http://localhost', 'https://localhost', 'urn:ietf:wg:oauth:2.0:oob']
        }}))
    if pytestconfig:
        capmanager = pytestconfig.pluginmanager.getplugin('capturemanager')
        capmanager.suspend_global_capture(in_=True)
    gs = google_sheets_lib.GoogleSheets(os.environ['GSHEETS_TESTING_FOLDER_ID'])
    client_secret_path.unlink()
    if pytestconfig:
        capmanager.resume_global_capture()
    return gs
