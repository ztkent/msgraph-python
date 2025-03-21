# MSGraph-Python
Simplify python interactions with the [Microsoft Graph API](https://github.com/microsoftgraph).  
This module extends usability of the modern [Microsoft Graph Python SDK](https://github.com/microsoftgraph/msgraph-sdk-python).  

Start development for applications that:  
- Fetch user information. More permissions will provide more information.
- Fetch [Unread] Teams messages and Chats
- Fetch [Unread] Outlook emails.
- Fetch [Today's] Calendar events.

## Setup
- Assuming you belong to an organization that uses Microsoft, and you are not the admin, this project connects to the Graph API with 'Delegated' permissions.  
- This means that you will only be able to read and write the authenticated user's data, not the entire organizations.  
- Hopefully, this will make it easier to get the necessary permissions from your organization's admin.  

### Create an Application Connection via Microsoft Graph API
#### Register an application in Azure Active Directory  
- Go to the Azure portal, then to Azure Active Directory, then to App registrations, and click on New registration.
- Enter a name for the application, select the supported account types, and then click Register.
- After the app is registered, note down the Application (client) ID and the Directory (tenant) ID.

#### Set API permissions
- In the App registrations page, go to API permissions, and click on "Add a Permission".
- Select Microsoft Graph, then Delegated permissions, and then add the necessary permissions.
    - `User.Read`: Read basic profile data
    - `Mail.Read`: Read user Outlook mail
    - `Calendars.Read`: Read user calendar
    - `Chat.Read`: Read user Teams chat messages 
    - Admin consent is required for the following permissions:
        - `ChannelMessage.Read.All`: Read user Teams channel messages

### Install MSGraph-Python
```bash
pip install git+https://github.com/ztkent/msgraph-python.git
```

### Authorize the application
- When creating a new connection, your application will connect via the [Microsoft Graph Python SDK](https://github.com/microsoftgraph/msgraph-sdk-python).
- During the connection flow, a login URL and authentication code are logged to the user.
- Follow the provided link, enter the authorization code, and login with OAuth.
- After a successful login, an application connection is generated from the response.
- This connection is used for future requests.
```python
async def example_connection(client_id, tenant_id):
    try: 
        graph_api = await NewGraphAPI(
            client_id="YOUR_CLIENT_ID",
            tenant_id="YOUR_TENANT_ID",
            scopes=["mail", "calendar", "teams-chat", "teams-channel"])
    except AuthorizationException as e:
        print(f"{e}")
        return
```