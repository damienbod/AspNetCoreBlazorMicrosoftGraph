[![.NET](https://github.com/damienbod/AspNetCoreBlazorMicrosoftGraph/actions/workflows/dotnet.yml/badge.svg)](https://github.com/damienbod/AspNetCoreBlazorMicrosoftGraph/actions/workflows/dotnet.yml)

# ASP.NET Core Blazor with Microsoft Graph samples

![User Calendar](https://github.com/damienbod/AspNetCoreBlazorMicrosoftGraph/blob/main/images/BlazorGraph_03.png)

## Blogs

[Use calendar, mailbox settings and Teams presence in ASP.NET Core hosted Blazor WASM with Microsoft Graph](https://damienbod.com/2021/12/20/use-calendar-mailbox-settings-and-teams-presence-in-asp-net-core-hosted-blazor-wasm-with-microsoft-graph/)

## History

- 2025-01-01 .NET 9
- 2024-07-03 Updated packages
- 2024-01-14 Updated packages, fix refresh bug
- 2024-01-14 Updated packages, improved CSP, now using a nonce
- 2023-11-22 Updated Graph 5 SDK
- 2023-11-21 Updated .NET 8
- 2023-03-02 Updated nuget packages, .NET 7
- 2022-07-09 Updated nuget packages
- 2022-01-28 Updated nuget packages
- 2022-01-14 Added Teams and Email APIs
- 2022-01-14 updated MailboxSettings and Calander to use application scopes
- 2021-12-19 first version

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

The Azure App registration requires the Graph API application **User.Read.All** **MailboxSettings.Read** scopes.

```json
"User.read User.Read.All MailboxSettings.Read"
```

### Calendar client

The Azure App registration requires the Graph API application **User.Read.All** **Calendars.Read** **Calendars.Read.Shared** scopes.

```json
 "User.read User.Read.All Calendars.Read Calendars.Read.Shared"
```

### Teams client

Requires the delegated **OnlineMeetings.ReadWrite** permission

```json
 "OnlineMeetings.ReadWrite"
```

# Links

https://blazorise.com/

https://github.com/AzureAD/microsoft-identity-web</p>

https://docs.microsoft.com/en-us/graph/api/user-get-mailboxsettings

https://docs.microsoft.com/en-us/graph/api/presence-get

https://docs.microsoft.com/en-us/aspnet/core/blazor/security/content-security-policy

https://github.com/andrewlock/NetEscapades.AspNetCore.SecurityHeaders
