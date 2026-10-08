# 🌟 Kid Activity Tracker ("Chore Hero") — Multi-iPad Sync

A kid-friendly web app designed for shared or private iPads. Kids can select their name, tap an activity they completed, and the app instantly logs the event (with timestamp, child name, and task) directly into your private **Google Sheet**.

Profiles and activities are **centrally synchronized via Google Sheets**: when you add or remove an activity, every iPad receives the update automatically without needing to be reconfigured.

---

## 📁 Files Included

1. [**`index.html`**](file:///home/gregory/Documents/Development/Kids_Chores_Tracker/index.html): The complete frontend web application.
2. [**`google-apps-script.js`**](file:///home/gregory/Documents/Development/Kids_Chores_Tracker/google-apps-script.js): The Google Apps Script backend managing logs, kids, and tasks tabs.

---

## 📊 How the Centralized Storage Works

Your Google Sheet is organized into 3 tabs:
1. **`Logs`**: Every time a kid completes a chore, a new timestamped row is recorded here.
2. **`Kids`**: The list of children profiles (`ID`, `Name`, `Avatar`, `Color`).
3. **`Tasks`**: The list of chores (`ID`, `Title`, `Icon`, `Color`).

Whenever an iPad opens the app or reconnects to Wi-Fi, it automatically fetches the latest **Kids** and **Tasks** from Google Sheets.

---

## 🚀 Setup Guide (Takes 3 Minutes)

### Step 1: Set Up Your Google Sheet

1. Open a new Google Sheet by visiting [sheets.new](https://sheets.new).
2. Name your spreadsheet (e.g., *"Kids Chore Tracker"*).
3. In the top menu, go to **Extensions** → **Apps Script**.
4. Delete any sample code in the editor.
5. Copy all code from [**`google-apps-script.js`**](file:///home/gregory/Documents/Development/Kids_Chores_Tracker/google-apps-script.js), paste it in, and click the **Save** (disk) icon.

### Step 2: Deploy as a Web App

1. In the upper-right corner of Apps Script, click **Deploy** → **New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Configure the settings:
   - **Description**: `Chore Hero API v2`
   - **Execute as**: `Me (your email address)`
   - **Who has access**: **`Anyone`** *(⚠️ Crucial: Allows iPads to log events and fetch chores without Google sign-in)*.
4. Click **Deploy**, authorize permissions when prompted, and copy the **Web app URL** (ends in `/exec`).

---

### Step 3: Connect & Sync Any iPad

1. Open [**`index.html`**](file:///home/gregory/Documents/Development/Kids_Chores_Tracker/index.html) in your browser.
2. Click the ⚙️ gear icon in the top right to open **Parent Settings**.
3. Enter PIN: **`1234`**.
4. Paste your **Google Sheets Web App URL** and click **Save URL**.
5. The app will immediately sync and pull the kids and chores from your Google Sheet.

*(Tip: When you host this on GitHub Pages or Netlify, you only have to paste the Web App URL once on each iPad, and thereafter all chores/kids update automatically!)*

---

### Step 4: Adding or Editing Chores Across All iPads

You have two convenient ways to update chores:

- **Method A (Direct in Google Sheets)**:
  Open your Google Sheet on your phone or computer, go to the **`Tasks`** tab, and add or change any row (e.g. add `Walk the dog | 🐕`). Next time any iPad loads the app, the new chore appears automatically!
- **Method B (Inside the App)**:
  Open Parent Settings on any iPad, add chores or kids, and tap **☁️ Push to Cloud**. All other iPads will receive the changes.

---

### Step 5: iPad Home Screen Installation

1. In Safari on each iPad, open your hosted app URL.
2. Tap the **Share** button (box with an arrow pointing up).
3. Tap **"Add to Home Screen"** → **Add**.
4. The app opens full-screen without Safari browser bars, exactly like a native iPad app.
