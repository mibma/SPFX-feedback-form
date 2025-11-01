
# SPFX Feedback Form

A SharePoint Framework (SPFx) web part that collects user feedback (rating + comments) and stores responses in a SharePoint list. This web part is Teams-aware (supports `hasTeamsContext`) and can be run in SharePoint or inside Microsoft Teams.

---
<img width="968" height="984" alt="image" src="https://github.com/user-attachments/assets/5dd063a6-89ba-4601-aa6f-c7e51a46eb38" />

## Features

* Simple rating (1–5) and text comment form
* Stores responses to a SharePoint list (`cloudlist')
* Friendly for tenant trial/dev usage (can be adapted to use external email providers)

---

## Table of contents

1. [Prerequisites](#prerequisites)
2. [Getting started (developer)](#getting-started-developer)
3. [Configure SharePoint list](#configure-sharepoint-list)
4. [Build & bundle](#build--bundle)
5. [Package & deploy to App Catalog](#package--deploy-to-app-catalog)
6. [Add web part to site / Teams](#add-web-part-to-site--teams)
7. [Common issues & troubleshooting](#common-issues--troubleshooting)
8. [Contributing](#contributing)
9. [License & author](#license--author)

---

## Prerequisites

* Node.js LTS (recommended v16 or v18 depending on SPFx toolchain used)
* Yeoman and SharePoint generator (if you want to scaffold or modify):
  `npm i -g yo @microsoft/generator-sharepoint`
* Gulp CLI: `npm i -g gulp-cli`
* Office 365 tenant with App Catalog (or dev tenant for local testing)
* SharePoint Framework toolchain installed per your SPFx version
* Git (for source control)

---

## Getting started (developer)

1. Clone the repo:

```bash
git clone https://github.com/mibma/SPFX-feedback-form.git
cd SPFX-feedback-form
```

2. Install dependencies:

```bash
npm install
```

3. Start the local workbench (for SPFx web part local testing):

```bash
gulp serve --nobrowser
```

* Open the local workbench at: `https://localhost:5432/workbench` OR use SharePoint online workbench if connecting to tenant data.

---

## Configure SharePoint list

Create a SharePoint list to store feedback. Example list settings:

**List name:** `Feedback`
**Columns:**

* `Title` (single line) — can hold item identifier or subject
* `Rating` (Number or Choice 1–5)
* `Comments` (Multiple lines of text)
* `SubmittedBy` (Person or use Created By)
* `SubmittedOn` (Created date is fine)

> If the web part expects a specific list name or internal column names, update the `constants` or configuration in the source where the list name and field internal names are defined.

---

## Build & bundle

To create a production bundle and package:

1. Build:

```bash
gulp build --ship
```

2. Bundle:

```bash
gulp bundle --ship
```

3. Package solution:

```bash
gulp package-solution --ship
```

This creates the `.sppkg` package under `sharepoint/solution/`.

---

## Package & deploy to App Catalog

1. Upload the generated `.sppkg` file to your tenant App Catalog (SharePoint Admin Center → More features → Apps → App catalog).
2. When prompted, choose whether to make it available to all sites.
3. If you want to make the web part available in Microsoft Teams, enable the **Teams** option in the package settings or use the “Make this solution available in Microsoft Teams” checkbox (if included when packaging).

---

## Add web part to a site / Teams

* **SharePoint site:** Go to a modern page → Edit → `+` → Add the web part (search by its name).
* **Microsoft Teams:** If you enabled Teams packaging, you can add the app from the Teams App Catalog and then add it as a tab in a Team channel.

---

## Configuration & common code points

* `hasTeamsContext?: boolean;` — optional prop indicating whether the web part is running inside Teams. The web part uses this to toggle CSS or layout via `styles.teams`.
* Update list name or internal field names in the code if your SP list uses different columns. Typical locations:

  * `src/webparts/<your-webpart>/components/<ComponentName>.tsx`
  * `src/webparts/<your-webpart>/components/I<Interface>.ts` (prop definitions)
  * `src/common/constants.ts` (if present)

---

## Troubleshooting

### Flow shows “sent” but recipient doesn’t receive

* Check the flow run outputs (To, Subject, message content).
* If using tenant trial, outbound mail via the Outlook connector may be blocked. Use alternative such as Gmail connector or SMTP with external provider.
* Inspect recipient mailbox (junk/quarantine) and run message trace (admin).

### `gulp bundle --ship` or `package-solution` errors

* Ensure correct Node version for SPFx version.
* Clear `node_modules` and reinstall:

  ```bash
  rm -rf node_modules package-lock.json
  npm install
  ```
* Re-run build steps.

### Merge/push issues to GitHub

* If `git push` fails with non-fast-forward, run:

  ```bash
  git pull --rebase origin main
  # resolve conflicts if any
  git push origin main
  ```

  or use `--force-with-lease` only if you are sure.

---

## Testing

* Local: `gulp serve` + local workbench
* Tenant: Add to a test site page and verify list entries are created correctly
* Teams: Add as an app tab and confirm `hasTeamsContext` behavior

---

## Contributing

1. Fork the repo
2. Create a topic branch: `git checkout -b feat/my-change`
3. Make changes and commit: `git commit -am "Describe change"`
4. Push branch and open a Pull Request

