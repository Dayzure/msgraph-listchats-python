# Sample python script to list chats for the signed-in user

This sample is best executed in the cloud shell ([https://shell.azure.com/]([https://shell.azure.com/])),
because all dependencies are already pre-installed.

You need to update the sample_parameters dictionary with the values of your AAD tenant:

```
sample_parameters = {
   "resource": "https://graph.microsoft.com/",
   "tenant" : "your-tenant.onmicrosoft.com",
   "authorityHostUrl" : "https://login.microsoftonline.com",
   "clientId" : "your-app-registration-client-id(app id)"
}
```

You have to (register application with Azure AD Tenant)[https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app] and grant Chat.Read API permissions. 

The sample is based on ADAL for Python and (Device Code Flow on Azure AD)[https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code].

It also include snippet for (acquring access token based on previously saved refresh token)[https://docs.microsoft.com/en-us/azure/active-directory/develop/v1-protocols-oauth-code#refreshing-the-access-tokens]:

```
### Use refresh token from a back-office app (i.e. WebJob running every hour)
refresh_token = "LOAD-A-REFRESH-TOKEN-FROM-PREVIOUS-SESSION"
token = context.acquire_token_with_refresh_token(
    refresh_token,
    sample_parameters['clientId'],
    RESOURCE
    )
### END refresh token
```
