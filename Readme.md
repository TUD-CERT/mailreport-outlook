# E-Mail Reporting Add-In for Microsoft Outlook (with on-premises Exchange)

Microsoft Outlook Add-In for reporting phishing or otherwise malicious E-Mails to an IT security department, developed by [TUD-CERT](https://tud.de/cert). This add-in is meant to be customized and deployed within well-defined organizational boundaries, such as for all employees covered by a CERT or members of a university.

## Features
![Add-In screenshots](docs/teaser.png?raw=true "The add-in in action")

* Support for reports via SMTP (summary mail with reported raw sample attached) or HTTP(S) to a [Lucy](https://lucysecurity.com)-compatible API (or both)
* User-provided optional comment for each report
* Configurable after-report action: move to junk/move to bin/keep mail
* Localization in English and German
* Respects selected UI theme
* Quickly adjustable organization-specific deployment settings
* Basic telemetry to report current add-in and MUA versions with each request
* Permission settings to disable unwanted features
* [Lucy](https://lucysecurity.com) phishing campaign detection

## Requirements
This add-in is specifically designed to be deployed in an on-premises Exchange environment. It is **not** compatible with Exchange Online or Microsoft 365.

On the client side, this add-in has been tested successfully with the following Outlook flavours:
* Microsoft Outlook 2016, 2019, 2021 and 2024 (Windows)
* Microsoft Outlook 2024 for Mac in Legacy Mode (MacOS)
* Outlook on the web / Outlook Web App

The project build system requires Node.js and npm.

## Technical Overview
This Outlook add-in adds a button to either the Ribbon (Outlook on Desktop) or to the message-specific buttons shown when reading a message in Outlook on the web (Outlook Web App). When clicked, the currently selected e-mail can be reported either as being malicious or spam. A third option launches a settings dialog. Reporting a malicious e-mail opens a task pane that enables users to attach an (optional) comment to their report. In contrast, reporting e-mails as spam happens immediately, doesn't open any views and can't be commented.

Reports can be sent either via e-mail/SMTP to a configurable reporting address or to a server that provides a [Lucy](https://lucysecurity.com)-compatible API (or both). The subjects of reports sent via SMTP use either `Phishing Report` or `Spam Report` as prefix to differentiate between the reporting options. Attached comments and basic telemetry (if enabled) are prepended to the e-mail body, while a raw sample of the reported e-mail is added as an attachment. The Lucy API doesn't support spam reports.

If a Lucy-style phishing campaign is detected - which is based on certain header fields present in the reported e-mail - the Lucy server is notified of the report via HTTP(S) and a dialog is shown to congratulate the reporter.

Since we strive to support a broad range of (older) Outlook versions that are still in use, the minimal required Mailbox API is 1.4. Due to many limitations in that old requirement set and the Outlook JavaScript API for on-premises Exchange/Outlook environments in general, missing functionality is implemented via EWS requests.

## How to build
Each organization using this reporting add-in has a specific set of requirements, such as the organization's spam reporting e-mail address, custom strings and messages shown within the add-in interface or custom icons. We call these organization-specific settings *deployment configurations* and place them inside the `configs/` directory. In there, each subfolder holds all modifications to the default configuration (which can be found in `templates/`) for a specific organization.

This project uses npm as build system, therefore the first step is to install or update dependencies with `npm install`. We defined various build and development commands that can be launched with `npm run <command> -- --env <params>`. When building the add-in, the name of the desired organization's folder denotes the deployment configuration to use. It has to be supplied as additional parameter to most build commands (via `--env` as environment variable). For example, to assemble the add-in with the deployment configuration of TU Dresden, which sits in `configs/tu-dresden.de`:

``` 
$ npm run build -- --env config=tu-dresden.de
> mailreport-outlook@0.0.1 build
> webpack --mode production --env config=tu-dresden.de

Deployment: tu-dresden.de
...
```

This command builds the project for distribution by parsing the default configuration in `templates/` and the deployment configuration in `configs/tu-dresden.de/` to produce add-in artifacts such as `defaults.json`, `manifest.xml` and `locales.json`. The finished build result is written to `dist/`. For deployment, serve the files in `dist/` on the same URL that was defined in the deployment configuration's `overrides.json` as `manifest.hosted_at`. To install the add-in, pick up the generated `manifest.xml` (from the top-level directory, it's not in `dist/`) and follow Microsoft documentation to either install a custom add-in from an XML file or distribute it to users directly from the Exchange server.

During development, use either `npm run build:dev -- --env config=<deployment config>` to write a one-shot development build to `dist/` or `npm run dev-server -- --env config=<development config>` to serve the add-in from a local web server that automatically picks up changes made to the sources during runtime. By default, this dev server listens at `https://localhost:3000`. The `manifest.xml` will also be updated to point to `localhost:3000`.

## Deployment Configurations
![Add-In settings screenshot](docs/settings.png?raw=true "Add-In settings")

Each organization's configuration directory requires at least a file named `overrides.json`. The JSON object in that file describes how the default configuration in `templates/` should be overwritten with organization-specific values. For reference, this repository includes a small example deployment configuration in `configs/example.com/` as a starting point for creating individual configurations.

To define values for the add-in's manifest file, which holds metadata such as the add-in identifier, name and version, modify the `manifest` key within `overrides.json`. A minimal example:

```
{
    "manifest": {
        "id": "<randomly-generated-id>",
        "provider_name": "Example Corp",
        "version": "1.0",
        "hosted_at": "https://example.com"
    },
    "defaults": {},
    "locales": {}
}
```

The add-in name and description are both set by overwriting their localizations strings. For an example, look further below.

The add-in's default configuration is kept in `templates/defaults.tpl` and can be overwritten in `overrides.json` within the top-level key `defaults`. Most of these settings can also be adjusted by users in the add-in configuration dialog from within Outlook. The following keys are essential for proper operation and should be reviewed thoroughly:

* **phishing_transport**: Defines which protocol(s) to use when reporting mails.
  * `"http"`: Send reports via HTTP(S) to a Lucy-compatible API.
  * `"smtp"`: Send reports as regular E-Mail with a summary of the reported E-Mail in the mail's body and the raw mail as attached EML file.
  * `"http+smtp"`: Send reports via HTTP(S) and SMTP simultaneously.
* **simulation_transport**: Protocol(s) to use when an E-Mail that belongs to a Lucy campaign is reported. Supports the same values as `phishing_transport`.
* **lucy_server**: Domain name of the Lucy API to send HTTP(s) reports to. Only required if HTTP(S) is set as phishing or simulation transport.
* **lucy_client_id**: The Lucy Client ID to send to the Lucy API with each report. Can be `null` to indicate *"all"* clients. Incidents on the Lucy server will then be shown as coming from client `N/A`.
* **smtp_to**: E-Mail address to send SMTP reports to. Only required if SMTP is set as phishing or simulation transport.

The remaining supported keys in `defaults` are
* **report_action**: How to deal with an E-Mail after it has been reported.
  * `"junk"`: Move it to the junk folder.
  * `"trash"`: Move it to the trash folder.
  * `"keep"`: Do nothing, keep it.
* **smtp_use_expressive_subject**: Determines which subject line to use when sending SMTP reports. If set to `false`, reports will simply use *Phishing Report* or *Spam Report* as subject lines. With this set to `true`, the subject line of the reported e-mail will be appended as well (e.g. *Phishing Report: Re: Urgent Letter*).
* **send_telemetry**: If set to `true`, this includes two header fields `Reporting-Agent` and `Reporting-Plugin` set to the current MUA and add-in identifier/version to all outgoing requests: Either as HTTP(S) header or preptended to the e-mail body. To disable, set to `false`. This setting can *not* be changed from within Outlook.
* **permit_advanced_config**: If set to `true`, users can modify `phishing_transport`, `simulation_transport`, `lucy_client_id`, `lucy_server`, `smtp_to`, `smtp_use_expressive_subject` and `update_url` from within Thunderbird (via *"Show advanced settings"* in the add-in's configuration dialog). Set to `false` so that user can only change the `report_action`. This setting can *not* be changed from within Outlook. **Notice**: This setting determines which options are stored locally within the MUA. If however this is set to `false`, future updates can transparently update the advanced settings (e.g. by switching to another SMTP reporting address). If this is set to `true`, all advanced settings are handled manually by users. Changing the SMTP reporting address would then require either a full redeployment of the add-in, ideally with a new identifier (to clear the local storage), or user's to manually update these settings.

Organizations can also overwrite individual localization strings from `templates/locales/` via the top-level key `locales`. For example:

```
{
  ...
  "locales": {
    "en": {
        "extensionName": "Phishing Report",
        "extensionDescription": "Enables reporting of suspicious e-mails to Example Corp."
    },
    "de": {
        "extensionName": "Phishing Report",
        "extensionDescription": "Ermöglicht das Melden auffälliger E-Mails an Example Corp."
    }
  }
  ...
}
```

To provide custom icons for the add-in, place them in a directory `images/` within your deployment configuration with names that mirror those in `templates/images/`. As with locales and default configuratino values, the template images in `templates/images/` will always be used as fallback if there is no custom override defined.
