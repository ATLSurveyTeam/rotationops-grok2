# RotationOps Scheduler

Simple scheduling tool for continuous AM/PM coverage (3:45 AM – 11:00 PM).  
Focuses on keeping all Tier 1 positions covered and assigning break relief correctly.

## Files in this folder

- `app.py` – the main application screen
- `scheduler.py` – the scheduling engine (first version)
- `RotationOps_v1.1_Coverage_Rules.xlsx` – employee and position data
- `requirements.txt` – list of required packages

---

## How to put this on Streamlit Cloud (free)

You already have a GitHub account. Follow these steps carefully.

### Step 1 – Create a new GitHub repository
1. Go to https://github.com
2. Click the **+** button (top right) → **New repository**
3. Name it something like `rotationops-scheduler`
4. Leave it **Public**
5. Click **Create repository**

### Step 2 – Upload the files
1. On the new repository page, click **uploading an existing file**
2. Drag all four files from this folder into the browser:
   - app.py
   - scheduler.py
   - requirements.txt
   - RotationOps_v1.1_Coverage_Rules.xlsx
3. Click **Commit changes**

### Step 3 – Deploy on Streamlit Cloud
1. Go to https://share.streamlit.io
2. Sign in with the **same GitHub account**
3. Click **New app**
4. Choose the repository you just created
5. Main file path: `app.py`
6. Click **Deploy**

Wait 1–2 minutes. Streamlit will give you a public link.

### Step 4 – Test on a City computer
Open the link on a City of Atlanta computer.  
If the page loads, the network allows it.  
If it is blocked, we will switch to the “run on personal laptop + export Excel” method.

---

## Current status of the engine
- First working version only
- Assigns people to positions (Tier 1 first)
- Tries to give each Tier 1 a break relief agent
- Still needs stronger Divest rules, fairness ranking, and better lunch separation

These improvements will be added after the app is successfully running on Streamlit Cloud.
