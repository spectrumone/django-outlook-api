import requests
import uuid
import json

outlook_api_endpoint = 'https://outlook.office.com/api/v2.0{0}'

# Generic API Sending
def make_api_call(method, url, token, payload = None, parameters = None):
    # Send these headers with all API calls
    headers = { 'User-Agent' : 'django-tutorial/1.0',
                'Authorization' : 'Bearer {0}'.format(token),
                'Accept' : 'application/json'}

    # Use these headers to instrument calls. Makes it easier
    # to correlate requests and responses in case of problems
    # and is a recommended best practice.
    request_id = str(uuid.uuid4())
    instrumentation = { 'client-request-id' : request_id,
                        'return-client-request-id' : 'true' }

    headers.update(instrumentation)

    response = None 

    payload = {
              "Subject": "Discuss the Calendar REST API",
              "Body": {
                "ContentType": "HTML",
                "Content": "I think it will meet our requirements!"
              },
              "Start": {
                  "DateTime": "2014-04-04T18:00:00",
                  "TimeZone": "Pacific Standard Time"
              },
              "End": {
                  "DateTime": "2014-04-04T19:00:00",
                  "TimeZone": "Pacific Standard Time"
              },
              "Attendees": [
                {
                  "EmailAddress": {
                    "Address": "janets@a830edad9050849NDA1.onmicrosoft.com",
                    "Name": "Janet Schorr"
                  },
                  "Type": "Required"
                }
              ]
            }

    if (method.upper() == 'GET'):
        response = requests.get(url, headers = headers, params = parameters)
    elif (method.upper() == 'DELETE'):
        response = requests.delete(url, headers = headers, params = parameters)
    elif (method.upper() == 'PATCH'):
        headers.update({ 'Content-Type' : 'application/json' })
        response = requests.patch(url, headers = headers, data = json.dumps(payload), params = parameters)
    elif (method.upper() == 'POST'):
        headers.update({ 'Content-Type' : 'application/json' })
        response = requests.post(url, headers = headers, data = json.dumps(payload), params = parameters)

    return response


def get_my_messages(access_token):
    get_messages_url = outlook_api_endpoint.format('/Me/Messages')

    # Use OData query parameters to control the results
    #  - Only first 10 results returned
    #  - Only return the ReceivedDateTime, Subject, and From fields
    #  - Sort the results by the ReceivedDateTime field in descending order
    query_parameters = {'$top': '10',
                        '$select': 'ReceivedDateTime,Subject,From',
                        '$orderby': 'ReceivedDateTime DESC'}

    r = make_api_call('GET', get_messages_url, access_token, parameters = query_parameters)

    if (r.status_code == requests.codes.ok):
        return r.json()
    else:
        return "{0}: {1}".format(r.status_code, r.text)


def get_my_events(access_token):
    get_events_url = outlook_api_endpoint.format('/Me/Events')
  
    # Use OData query parameters to control the results
    #  - Only first 10 results returned
    #  - Only return the Subject, Start, and End fields
    #  - Sort the results by the Start field in ascending order
    query_parameters = {'$top': '10',
                        '$select': 'Subject,Start,End',
                        '$orderby': 'Start/DateTime ASC'}
                      
    r = make_api_call('GET', get_events_url, access_token, parameters = query_parameters)
  
    if (r.status_code == requests.codes.ok):
        return r.json()
    else:
        return "{0}: {1}".format(r.status_code, r.text)

def post_my_events(access_token):
    post_events_url = outlook_api_endpoint.format('/Me/Events')
    r = make_api_call('POST', post_events_url, access_token)
    if (r.status_code == requests.codes.ok):
        return r.json()
    else:
        return "{0}: {1}".format(r.status_code, r.text)



def get_my_contacts(access_token):
    get_contacts_url = outlook_api_endpoint.format('/Me/Contacts')
  
    # Use OData query parameters to control the results
    #  - Only first 10 results returned
    #  - Only return the GivenName, Surname, and EmailAddresses fields
    #  - Sort the results by the GivenName field in ascending order
    query_parameters = {'$top': '10',
                        '$select': 'GivenName,Surname,EmailAddresses',
                        '$orderby': 'GivenName ASC'}
                      
    r = make_api_call('GET', get_contacts_url, access_token, parameters = query_parameters)
  
    if (r.status_code == requests.codes.ok):
        return r.json()
    else:
        return "{0}: {1}".format(r.status_code, r.text)
