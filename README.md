[![.NET](https://github.com/damienbod/AspNetCoreBlazorMicrosoftGraph/actions/workflows/dotnet.yml/badge.svg)](https://github.com/damienbod/AspNetCoreBlazorMicrosoftGraph/actions/workflows/dotnet.yml)

# ASP.NET Core Blazor with Microsoft Graph samples

### Email client

The Azure App registration requires the Graph API delegated **Mail.Send** and the **Mail.ReadWrite** scopes.

```json
"User.read Mail.Send Mail.ReadWrite"
```

### Presence client

The Azure App registration requires the Graph API delegated **User.Read.All**  scope.

```json
"User.read User.Read.All"
```

### User Mailbox settings client

The Azure App registration requires the Graph API delegated **User.Read.All** **MailboxSettings.Read** scopes.

```json
"Scope" value="User.read User.Read.All MailboxSettings.Read"
```

### Calendar client

The Azure App registration requires the Graph API delegated **User.Read.All** **Calendars.Read** Calendars.ReadWrite** scopes.

```json
 "User.read User.Read.All Calendars.Read Calendars.ReadWrite"
```

# Links

https://blazorise.com/
