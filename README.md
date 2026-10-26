# Bridmet Outlook signature add-in (starter)

This starter hosts an Outlook compose task pane and inserts the signature at the current cursor position.

## Included
- `index.html`, `styles.css`, `taskpane.js`: task pane and insertion logic
- `config.js`: optional email-to-name mapping
- `assets/bsi-certification.png`: BSI image provided for the signature
- `manifest.xml`: Outlook add-in manifest
- `staticwebapp.config.json`: basic Azure Static Web Apps configuration

## Before deployment
1. Create an Azure Static Web App and deploy the contents of this folder as the app's static files.
2. Replace every `YOUR-STATIC-WEB-APP` in `manifest.xml` with the exact HTTPS hostname Azure assigns.
3. Add your actual staff email/name mappings in `config.js` if you want automatic prefill. Otherwise users can type/check their name.
4. Create the three manifest icon files (`icon-16.png`, `icon-32.png`, `icon-80.png`, plus `icon-64.png` and `icon-128.png`) or update the manifest to point to real hosted icon files. The manifest references are placeholders and must be fixed before deployment.
5. Validate the manifest with Microsoft's Office Add-in manifest validation tools and test in your exact New Outlook build before centrally deploying it.

## Test flow
- In New Outlook, open a new message, place the cursor where the signature should go, and open **Insert my signature**.
- Confirm the name, click the button, and inspect the inserted signature.
- Send a test message to yourself and check the signature in the received email and on a mobile/web client if available.

## Important limitations
- The user-profile email and display name are used only to prefill the name. They are not proof of identity; users can edit the field.
- The image is referenced by an HTTPS URL hosted with the add-in. Some recipients' email clients may block remote images until they choose to display them.
- This inserts at the current cursor position. If you want the signature at the end of the message, place the cursor there before clicking the button.
- This starter is not production-validated and has not yet been tested against your tenant/client. Validate and test before organisation-wide deployment.
- The current template follows the screenshot's dark background and white text. Outlook dark mode may affect rendering, so test before rollout.

## Staff mapping format
```js
users: [
  { email: "person@bridmet.co.uk", name: "Person Name" },
  { email: "another@bridmet.co.uk", name: "Another Person" }
]
```
