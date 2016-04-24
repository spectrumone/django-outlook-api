from urllib.parse import quote, urlencode
import requests
import base64
import json

# Client ID and secret
client_id = '5415a067-951a-47a1-8940-19af628afb1f'
client_secret = 'TfMP4Eorejvpvt0jfTKMgjk'

# Constant strings for OAuth2 flow
# The OAuth authority
authority = 'https://login.microsoftonline.com'

# The authorize URL that initiates the OAuth2 client credential flow for admin consent
authorize_url = '{0}{1}'.format(authority, '/common/oauth2/v2.0/authorize?{0}')

# The token issuing endpoint
token_url = '{0}{1}'.format(authority, '/common/oauth2/v2.0/token')

# The scopes required by the app
scopes = [ 'openid', 
           'https://outlook.office.com/mail.read',
           'https://outlook.office.com/calendars.read',
           'https://outlook.office.com/contacts.read', ]


def get_signin_url(redirect_uri):
    # Build the query parameters for the signin url
    params = { 'client_id': client_id,
               'redirect_uri': redirect_uri,
               'response_type': 'code',
               'scope': ' '.join(str(i) for i in scopes)}

    signin_url = authorize_url.format(urlencode(params))

    return signin_url


def get_token_from_code(auth_code, redirect_uri):
    # Build the post form for the token request
    post_data = {'grant_type': 'authorization_code',
                 'code': auth_code,
                 'redirect_uri': redirect_uri,
                 'scope': ' '.join(str(i) for i in scopes),
                 'client_id': client_id,
                 'client_secret': client_secret}

    r = requests.post(token_url, data=post_data)
    try:
        return r.json()
    except:
        return 'Error retrieving token: {0} - {1}'.format(r.status_code, r.text)


def get_user_email_from_id_token(id_token):
    # JWT is in three parts, header, token, and signature
    # separated by '.'

    token_parts = id_token.split('.')
    encoded_token = token_parts[1]
  
    # base64 strings should have a length divisible by 4
    # If this one doesn't, add the '=' padding to fix it
    leftovers = len(encoded_token) % 4
    if leftovers == 2:
        encoded_token += '=='
    elif leftovers == 3:
        encoded_token += '='
  
    # URL-safe base64 decode the token parts
    decoded = base64.urlsafe_b64decode(encoded_token.encode('utf-8')).decode('utf-8')
  
    # Load decoded token into a JSON object
    jwt = json.loads(decoded)
    return jwt['preferred_username']
