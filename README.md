# Introduction 
TODO: Give a short introduction of your project. Let this section explain the objectives or the motivation behind this project. 

# Getting Started
TODO: Guide users through getting your code up and running on their own system. In this section you can talk about:
1.	Installation process
2.	Software dependencies
3.	Latest releases
4.	API references

# Build and Test
TODO: Describe and show how to build your code and run the tests. 

# Contribute
TODO: Explain how other users and developers can contribute to make your code better. 

If you want to learn more about creating good readme files then refer the following [guidelines](https://docs.microsoft.com/en-us/azure/devops/repos/git/create-a-readme?view=azure-devops). You can also seek inspiration from the below readme files:
- [ASP.NET Core](https://github.com/aspnet/Home)
- [Visual Studio Code](https://github.com/Microsoft/vscode)

- [Chakra Core](https://github.com/Microsoft/ChakraCore)


### **Authentication APIs**

#### **1. OKTA Token Authentication**

**Purpose:** Retrieve an OKTA Bearer Token for authenticating with Epic API endpoints.
**Base URL:** `https://derivco.okta-emea.com`
**Endpoint:** Token request endpoint (managed internally via `Epic.WebApi.Client.DLL`)
**Authentication Type:** Username / Password
**Token Lifetime:** 2 hours

**Python Implementation:**

```python
import requests
from datetime import datetime, timedelta

def get_okta_token(username, password, client_id="0oa1l6jqqgsJbUHyu0i7", okta_host="derivco.okta-emea.com"):
    """
    Request an OKTA Bearer Token for Epic API access.
    """
    url = f"https://{okta_host}/oauth2/v1/token"
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'grant_type': 'password',
        'username': username,
        'password': password,
        'client_id': client_id,
        'scope': 'openid profile email'
    }
    response = requests.post(url, headers=headers, data=data)
    response.raise_for_status()
    token_data = response.json()
    return {
        'token': token_data['access_token'],
        'expiry': datetime.now() + timedelta(hours=2)
    }
```

**Response Structure:**

```json
{
    "token": "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9...",
    "expiry": "2025-10-31T14:30:00"
}
```
