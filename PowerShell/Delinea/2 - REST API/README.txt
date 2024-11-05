This method uses REST API, but for this to work, you need to make sure you authenticate your session first.

When you first login in Chrome, you can bring up the dev tools by pressing F12.
Next click the "Application" tab, and from there look under Storage--> Cookies in the left side navigation.
Expand the cookie and the URL underneath. You should be able to get your token for ihawu and Thycotic_Location from here. Copy those values into the powershell script.

Keep in mind these cookies are stored as a session, so if you close out of Chrome, your session won't exist anymore and you will need to reauthenticate again.

The important thing to note here is since this data is time sensitive you need to keep chrome open. It doesn't necessarily need to be open at the delinea portal, just the process had to exist so the session is active. Also keep in mind that sometimes the session can expire even if you have your session still open/active. I believe every few days your session cookies need to reauthenticate for security purposes so just moniter the ihawu and thycotic_location before running your script. If the value for those 2 variables are different than the last time you ran your script, you will need to update them in your script to reflect the new values.

