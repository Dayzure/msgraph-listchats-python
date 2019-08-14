import json
import logging
import os
import sys
import adal
import urllib.request

from html.parser import HTMLParser

class MLStripper(HTMLParser):
   def __init__(self):
      super().__init__()
      self.reset()
      self.strict = False
      self.convert_charrefs= True
      self.fed = []
   def handle_data(self, d):
      self.fed.append(d)
   def get_data(self):
      return ''.join(self.fed)

def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return s.get_data()

def turn_on_logging():
    logging.basicConfig(level=logging.DEBUG)
    #or,
    #handler = logging.StreamHandler()
    #adal.set_logging_options({
    #    'level': 'DEBUG',
    #    'handler': handler
    #})
    #handler.setFormatter(logging.Formatter(logging.BASIC_FORMAT))


sample_parameters = {
   "resource": "https://graph.microsoft.com/",
   "tenant" : "your-tenant.onmicrosoft.com",
   "authorityHostUrl" : "https://login.microsoftonline.com",
   "clientId" : "your-app-registration-client-id(app id)"
}

authority_host_url = sample_parameters['authorityHostUrl']
authority_url = authority_host_url + '/' + sample_parameters['tenant']
clientid = sample_parameters['clientId']
GRAPH_RESOURCE = 'https://graph.microsoft.com/'
RESOURCE = sample_parameters.get('resource', GRAPH_RESOURCE)

#uncomment for verbose logging
#turn_on_logging()

### Initialize instance of ADAL
context = adal.AuthenticationContext(authority_url)
###

### Device Code Flow (https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code)
### Uncomment bellow to initially log-in the enduser
code = context.acquire_user_code(RESOURCE, clientid)
print(code['message'])
token = context.acquire_token_with_device_code(RESOURCE, code, clientid)
### Device Code Flow ends

### Info, uncomment to see details about the signed-in user, the access token and the refresh token
#print('Here is the token from "{}":'.format(authority_url))
#print(json.dumps(token, indent=2))


#### TODO
# Save the following infos in KeyVault
# token['accessToken']  -  this is the access token with live span of 1 hour
# token['refreshToken'] -  this is the refresh token, used to require new access token. 
#                          Refresh tokens have lifespan of "untill revoked"! 
#                          Password change is one of the events that can cause refresh token invalidation
# token['userId']       -  This is the UPN of the signed-in user
####

### Use refresh token from a back-office app (i.e. WebJob running every hour)
# refresh_token = "LOAD-A-REFRESH-TOKEN-FROM-PREVIOUS-SESSION"
# token = context.acquire_token_with_refresh_token(
#     refresh_token,
#     sample_parameters['clientId'],
#     RESOURCE
#     )
### END refresh token

# Construct the bearer authorization header for Microsoft Graph
bearerAuth = 'Bearer ' + token['accessToken']


print('Chats')
print('-----------------------------------------------')
# Call the MS Graph List Chats operation (https://docs.microsoft.com/en-us/graph/api/chat-list?view=graph-rest-beta&tabs=http)
chats = urllib.request.urlopen(urllib.request.Request(
    "https://graph.microsoft.com/beta/me/chats",
    headers= {
       "Accept" : 'application/json',
       "Authorization": bearerAuth
       }
)).read().decode('utf-8')
jsonChats = json.loads(chats)

print('Messages from the first 4 chats')
print('-----------------------------------------------')
# loop through the first 4 results
for i in range(4):
   # grab the messages for each of those chats (https://docs.microsoft.com/en-us/graph/api/chatmessage-list?view=graph-rest-beta&tabs=http)
   msg = urllib.request.urlopen(urllib.request.Request(
      "https://graph.microsoft.com/beta/me/chats/"+jsonChats["value"][i]["id"]+"/messages",
      headers= {
         "Accept" : 'application/json',
         "Authorization": bearerAuth
         }
   )).read().decode('utf-8')
   jsonMsg = json.loads(msg)

   print ("Chat messages in " + jsonChats["value"][i]["id"])
   print('-----------------------------------------------')
   for j in range(len(jsonMsg["value"])):
      print("From: " + jsonMsg["value"][j]["from"]["user"]["displayName"])
      print("Content: " + strip_tags(jsonMsg["value"][j]["body"]["content"]))
      print('-----------------------------------------------') 
   #print(json.dumps(jsonMsg, indent=2))
