import pickle

with open('token_gmail_v1.pickle', 'rb') as file:
    cred = pickle.load(file)

print(cred)
# Print relevant details
print(f"Access token: {cred.token}")
print(f"Refresh token: {cred.refresh_token}")
print(f"Expiry: {cred.expiry}")
print(f"Client ID: {cred.client_id}")
print(f"Client Secret: {cred.client_secret}")  # Be cautious with this one!
print(f"Token URI: {cred.token_uri}")
# print(f"User Agent: {cred._user_agent}")  # If available
print(f"Scopes: {cred.scopes}")
