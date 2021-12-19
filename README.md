[![.NET](https://github.com/damienbod/AspNetCoreBlazorMicrosoftGraph/actions/workflows/dotnet.yml/badge.svg)](https://github.com/damienbod/AspNetCoreBlazorMicrosoftGraph/actions/workflows/dotnet.yml)

# ASP.NET Core Blazor with Microsoft Graph samples

![User Calendar](https://github.com/damienbod/AspNetCoreBlazorMicrosoftGraph/blob/main/images/BlazorGraph_03.png)


### Email client

The Azure App registration requires the Graph API delegated **Mail.Send** and the **Mail.ReadWrite** scopes.

```json
"User.read Mail.Send Mail.ReadWrite"
```

### Presence client

The Azure App registration requires the Graph API delegated **User.Read.All** and **Presence.Read.All** scope.

```json
"User.read User.Read.All Presence.Read.All"
```

### User Mailbox settings client

The Azure App registration requires the Graph API delegated **User.Read.All** **MailboxSettings.Read** scopes.

```json
"User.read User.Read.All MailboxSettings.Read"
```

### Calendar client

The Azure App registration requires the Graph API delegated **User.Read.All** **Calendars.Read **Calendars.Read.Shared** scopes.

```json
 "User.read User.Read.All Calendars.Read Calendars.Read.Shared"
```

# Links

https://blazorise.com/

https://github.com/AzureAD/microsoft-identity-web</p>

https://docs.microsoft.com/en-us/graph/api/user-get-mailboxsettings

https://docs.microsoft.com/en-us/graph/api/presence-get
