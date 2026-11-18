# YRDSB sign-in setup (do these in order)

The page now asks for a Google sign-in. Every request to your Apps Script carries the sign-in token, and Code.gs asks Google to verify it
and only accepts emails ending in `@gapps.yrdsb.ca` or `@yrdsb.ca`. Until you finish steps 1-3 nobody can get in (it fails closed).

## 1. Create a Google "client ID" (free, about 5 minutes)
1. Go to https://console.cloud.google.com and sign in (use your YRDSB account if the board allows it).
2. Create a project (any name, e.g. "Williams inventory").
3. APIs & Services > OAuth consent screen: choose **Internal** if offered (only your organization), otherwise **External**. Fill in the app name and your email; save.
4. APIs & Services > Credentials > Create credentials > **OAuth client ID** > Application type **Web application**.
5. Under **Authorized JavaScript origins** add your site address with no path, e.g. `https://YOUR-USERNAME.github.io`
6. Create, then copy the Client ID (ends in `.apps.googleusercontent.com`).

## 2. Put the client ID in the page
In `index.html`, find `const CLIENT_ID='PASTE-YOUR-CLIENT-ID.apps.googleusercontent.com';` and replace it with your client ID. Upload `index.html` to GitHub again.

## 3. Update Apps Script
1. Open your Apps Script project. Replace everything in `Code.gs` with the new `Code.gs`.
2. Project Settings > Script Properties > add `GOOGLE_CLIENT_ID` = the same client ID. (Optional: `ALLOWED_DOMAINS` = `gapps.yrdsb.ca,yrdsb.ca` to change who is allowed.)
3. Deploy > Manage deployments > pencil > Version: **New version** > Deploy. Keep "Who has access" as **Anyone**; the sign-in check happens inside the script.
4. If Google asks to authorize "connect to an external service", allow it.

## Check it works
Open the site on a phone: you should see the sign-in screen. A YRDSB account gets in; any other Google account sees "that account is not allowed".

## Notes
- The admin password still protects admin actions (add/edit, repair, log). Student accounts can view the list and sign out/return instruments.
- Confirm the exact email domains your students and staff use and set `ALLOWED_DOMAINS` accordingly.
- Sign-in lasts about an hour, then the page asks again (usually automatically).
