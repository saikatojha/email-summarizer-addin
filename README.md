# Email Summarizer – Outlook Add-in
### Netwoven | Powered by Microsoft Copilot Studio

One-click email summarization directly inside Outlook. No copy-paste, no tab switching.

---

## Files
| File | Purpose |
|---|---|
| `taskpane.html` | The add-in UI – reads the email, calls Copilot Studio, shows the summary |
| `manifest.xml` | Tells Outlook about the add-in, where to find the files, where to put the button |
| `commands.html` | Minimal required file (leave as-is) |

---

## Setup (one-time, ~20 minutes)

### Step 1 – Prepare your Copilot Studio agent for Direct Line

1. Go to **https://copilotstudio.microsoft.com**
2. Open your **Email Summarizer** agent
3. Go to **Settings → Security → Authentication**
4. Set to **No authentication** (for testing — we will secure this in Phase 2)
5. Save and **Publish** the agent
6. Go to **Channels → Direct Line**
7. Copy the **Secret Key** (the long string under "Secret keys")

### Step 2 – Put the secret into taskpane.html

Open `taskpane.html` in any text editor. Find this line near the bottom:

```
const DL_SECRET = 'REPLACE_WITH_YOUR_DIRECT_LINE_SECRET';
```

Replace the placeholder with your actual secret:

```
const DL_SECRET = 'abc123...your_actual_secret_here';
```

Save the file.

### Step 3 – Host the files on Azure Blob Storage (HTTPS required)

Outlook add-ins must be served over HTTPS. Azure Blob Static Website is the fastest option.

**In Azure Portal:**
1. Create a **Storage Account** (or use an existing one)
2. Go to the storage account → **Static website** (under Data Management)
3. Enable it — note the **Primary endpoint URL** (looks like `https://myaccount.z13.web.core.windows.net`)
4. Upload all 3 files (`taskpane.html`, `manifest.xml`, `commands.html`) to the **`$web`** container
5. Set container access to **Blob (anonymous read)**

**Alternatively — GitHub Pages:**
1. Push the 3 files to a public GitHub repo
2. Go to Settings → Pages → Deploy from branch (main / root)
3. Your URL will be: `https://yourusername.github.io/your-repo-name/`

### Step 4 – Update manifest.xml with your hosting URL

Open `manifest.xml` and replace every instance of `YOUR_HOST_URL` with your actual URL.

There are 3 places:
- Line with `SourceLocation` under `<Form xsi:type="ItemRead">`
- Line with `Commands.Url`
- Line with `Taskpane.Url`

Example — if your Azure Blob URL is `https://myaccount.z13.web.core.windows.net`:
```xml
<SourceLocation DefaultValue="https://myaccount.z13.web.core.windows.net/taskpane.html"/>
```

Re-upload the updated `manifest.xml` to the `$web` container.

### Step 5 – Sideload the add-in in Outlook on the Web

1. Open **https://outlook.office.com** (Outlook on the Web)
2. Open any email
3. Click **More options (...)** in the email toolbar → **Get Add-ins**
4. Click **My add-ins** tab on the left
5. Scroll to the bottom → **Add a custom add-in** → **Add from file...**
6. Upload your `manifest.xml` file
7. Accept the warning dialog

Done! You should now see a **"Summarize Email"** button in the Outlook toolbar when reading emails.

---

## Using the add-in

1. Open any email in Outlook on the Web (or New Outlook for Windows)
2. Click **Summarize Email** in the toolbar — a side panel opens
3. Click the **Summarize this email** button in the panel
4. Wait ~5–15 seconds for the Copilot Studio agent to respond
5. The structured summary appears — you can also click **Copy summary**

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| "Could not connect" error | Check that `DL_SECRET` is correct and agent is published |
| "Agent did not respond" | Open Copilot Studio and test the agent in the built-in test chat first |
| Button not appearing | Reload Outlook — the add-in sometimes needs a fresh page load after sideloading |
| Blank task pane | Check browser console (F12) for CORS or CSP errors; ensure files are served over HTTPS |
| Summary looks like raw Markdown | The `marked.js` CDN script may have failed to load — check your network |

---

## Phase 2 Roadmap

- [ ] Move Direct Line secret to an Azure Function proxy (security hardening)
- [ ] Add Entra authentication so the agent can verify the caller's identity
- [ ] Deploy org-wide via M365 Admin Center (no sideloading needed)
- [ ] Add a "Reply with summary" button that pre-populates a reply with the summary prepended
- [ ] Replace placeholder icons with Netwoven branding

---

*Built by Netwoven Inc. | Powered by Microsoft Copilot Studio + Office JS API*
