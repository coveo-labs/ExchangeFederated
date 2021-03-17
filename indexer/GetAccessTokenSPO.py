# This information is obtained upon registration of a new Outlook Application
client_id = 'a6099661-70b9-471d-a638-13293159ef27'
client_secret = 'p7z?/Tv1t]x4wn]tjJMg3-RWcHNp@4.D'
tenant_id = '3961c1fa-1f64-4c73-ba2c-6923b1a375c9'
tenant_id = 'common'


# OAuth endpoints given in Outlook API documentation
authorization_base_url = 'https://login.microsoftonline.com/'+tenant_id+'/oauth2/v2.0/authorize'
token_url = 'https://login.microsoftonline.com/'+tenant_id+'/oauth2/v2.0/token'
#scope = ['offline_access','Directory.Read.All','Group.Read.All','User.Read','Sites.FullControl.All','AllSites.FullControl','User.Read.All','openid','profile','email']
scope = ['offline_access','Directory.Read.All','Group.Read.All','User.Read','Sites.FullControl.All','https://microsoft.sharepoint-df.com/AllSites.FullControl','User.Read.All','openid','profile','email']

redirect_uri = 'https://localhost/'     # Should match Site URL

from requests_oauthlib import OAuth2Session
outlook = OAuth2Session(client_id,scope=scope,redirect_uri=redirect_uri)

# Redirect  the user owner to the OAuth provider (i.e. Outlook) using an URL with a few key OAuth parameters.
authorization_url, state = outlook.authorization_url(authorization_base_url)
print ('Please go here and authorize,', authorization_url)

# Get the authorization verifier code from the callback url
redirect_response = input('Paste the full redirect URL here:')

# Fetch the access token
token = outlook.fetch_token(token_url,client_secret=client_secret,authorization_response=redirect_response)

print (token)
# Fetch a protected resource, i.e. calendar information
#o = outlook.get('https://outlook.office.com/api/v1.0/me/calendars')
#print (o.content)
