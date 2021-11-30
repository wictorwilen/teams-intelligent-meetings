# meetai - Microsoft Teams App

Teams meeting and calling demo

## App setup

### Bot app

**Redirect URIs**

- https://token.botframework.com/.auth/web/redirect
- https://YOUR-URL/_api/consent/bot

**Application permissions**

- AppCatalog.ReadWrite.All
- Calls.Initiate.All
- Calls.InitiateGroupCall.All
- Calls.JoinGroupCall.All
- Chat.Create
- Chat.ReadWrite.All
- ChatMember.ReadWrite.All
- OnlineMeetings.ReadWrite.All
- TeamsAppInstallation.ReadWriteForChat.All
- TeamsTab.ReadWriteForChat.All

**Delegated permissions**

- Chat.ReadWrite

### Azure Bot Service

**OAuth Settings**

- name: AzureAD
- Provider: AAD v2
- Client ID: Tab App Id
- Client Secret: Tab app secret
- Tenant ID: common
- Scopes: Chat.ReadWrite
### Tab app

**Redirect URIs**

- https://YOUR-URL/_api/consent/tab
- https://token.botframework.com/.auth/web/redirect

**Application ID URI**

- api://YOUR-URL/botid-<BOTID>

**Delegated permissions**

- AppCatalog.Read.All
- Chat.ReadWrite
- Files.Read.All
- email
- offline_access
- openid
- profile
- User.Read

### Azure Services

- Azure Speech Service
- Azure Language Service