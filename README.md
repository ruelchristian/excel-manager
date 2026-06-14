# Excel Cloud Sidebar Add-in Sideloading Guide

This folder contains your Microsoft Excel Add-in, which replicates the exact **ADB Subscription Manager** interface from Google Sheets inside Excel for the Web (OneDrive/Excel Cloud).

## Folder Contents
* **`manifest.xml`**: Registers the add-in inside Excel and defines sidebar settings.
* **`taskpane.html`**: The HTML frontend layout for your sidebar (search list and form).
* **`taskpane.js`**: The Javascript engine using the Microsoft `Office.js` API to manage your spreadsheet rows, sort records, and apply pastel formats.

---

## How to Set It Up (Free & No Credit Card)

To run custom scripts and sidebars in Excel Cloud, the HTML and JS files must be hosted on a secure `https://` address. The easiest and most permanent free way to do this is **GitHub Pages**.

### Step 1: Host the HTML and JS Files on GitHub Pages
1. Go to [GitHub](https://github.com/) and sign up for a free account (no credit card or trial registrations).
2. Create a new repository named `excel-manager` (make it **Public**).
3. Upload **`taskpane.html`** and **`taskpane.js`** to this new repository.
4. In your repository, go to **Settings** (tab at the top) > **Pages** (in the left menu).
5. Under **Build and deployment** > **Branch**:
   * Set it to `main` (or `master`).
   * Select `/ (root)`.
   * Click **Save**.
6. Wait 30 seconds, then refresh the page. GitHub will show a message at the top: **"Your site is live at: https://YOUR_USERNAME.github.io/excel-manager/"**

---

### Step 2: Edit your Manifest File
Now that your files are hosted on the internet, update your local `manifest.xml` to point to them:

1. Open **`manifest.xml`** in a text editor.
2. Locate line 15:
   ```xml
   <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
   ```
3. Replace it with your live HTML file link:
   ```xml
   <SourceLocation DefaultValue="https://YOUR_USERNAME.github.io/excel-manager/taskpane.html"/>
   ```
4. Save the `manifest.xml` file.

---

### Step 3: Upload the Add-in to Excel Cloud (OneDrive)
1. Open your Excel workbook (`ADB MONTHLY (MASTER FILE) FINAL.xlsx`) in your web browser via **OneDrive**.
2. Click the **Insert** tab in the top ribbon menu.
3. Click **Add-ins** (or **Office Add-ins**).
4. In the top-right corner of the Add-ins modal, click **Upload My Add-in**.
5. Choose your edited **`manifest.xml`** file from your computer and click **Upload**.
6. A sidebar titled **ADB Subscription Manager** will open on the right side of Excel!

---

## How It Works
* **View & Search Tab**: Lists all subscriptions starting from row 4 on sheet `"ADB MASTER LIST MONTHLY"`. As you search, the cards filter in real-time. Click any card to edit it.
* **Add/Edit Form**: Allows you to enter details and click **Save Subscription**. This writes directly to the row (or appends a new one), automatically applies formatting settings, and triggers the sort function.
* **Automatic Sorting & Color Formatting**: After saving a record, the add-in automatically sorts the workbook (New/Renewal on top, Cancelled at bottom, and soonest expiries first) and applies soft pastel colors based on their status.
* **Re-Format Button**: Click **Re-Format Sheet Colors** to scan the sheet and re-apply colors at any time.
"# excel-manager" 
