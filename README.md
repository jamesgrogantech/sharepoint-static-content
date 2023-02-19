# SharePoint Static Content

This is a CLI based node library which takes a SharePoint list and extracts the content and images to be served on a static site.

This allows SharePoint lists to be used as a rudimentary CMS system.

## Prerequisites

- SharePoint site (SharePoint Online only)
- SharePoint list
- Active Directory service principle (Username and Password for MSAL login)
- Active Directory App Registration with API permissions for "Sites.Read.All"
- Client Secret for App Registration

## Environment Variables

This package requires these variables to be set first.

```dotenv
SHAREPOINT_USERNAME="service principle username"
SHAREPOINT_PASSWORD="service principle password"
CLIENT_ID="Azure AD app registration client id (application id)"
CLIENT_SECRET="Azure AD app registration client secret"
TENANT_ID="Azure AD tenant id"
SITE_ID="sharepoint site id"
LIST_ID="sharepoint list id"

```
