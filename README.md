# Group Meeting Slot Booking (Python + Streamlit)

34 groups · 15-min slots · Week 1 + Week 2 · Sun–Thu office hours (from your Excel).
70 slots in total. Each group books ONE slot; a booked slot shows "Booked · G5" to everyone.

## A. Run on your own PC (test)
```
pip install -r requirements.txt
streamlit run app.py
```
Opens at http://localhost:8501. Bookings are saved in `bookings.xlsx`.
Admin password (sidebar) = `admin123` until you set your own in secrets.

## B. Put it online and share in Google Classroom
1. **GitHub**: create a free account, make a new repository, upload `app.py`, the `assets` folder (logo), `.streamlit/config.toml`,
   `requirements.txt` and `.gitignore` (NOT secrets.toml).
2. **Streamlit Cloud**: go to share.streamlit.io, sign in with GitHub,
   click "Create app", choose your repo, main file = `app.py`, Deploy.
3. **Google Sheet storage** (important: online, a plain file can be wiped when the
   app restarts, so store bookings in a Google Sheet):
   1. console.cloud.google.com, create a project.
   2. Enable "Google Sheets API".
   3. IAM & Admin, Service Accounts, Create, then Keys, Add key, JSON (downloads a file).
   4. Create an empty Google Sheet, click Share, and add the `client_email`
      from the JSON as **Editor**.
   5. In Streamlit Cloud: your app, Settings, Secrets. Paste the content of
      `.streamlit/secrets.toml.example` filled with your JSON values, sheet URL and
      your own admin password. Save (the app restarts).
   The sidebar in admin mode will now show "Storage: Google Sheet".
4. **Google Classroom**: Stream, Announce (or Classwork, Material), Add, Link.
   Paste your app URL (https://your-app.streamlit.app) and post.

## How groups use it
1. Open the link and select their group.
2. Click any "Free" slot. It turns into "Booked".
3. Want to change? Press "Cancel my booking" and pick again.

## Admin (sidebar, password)
- Download the full schedule as Excel (all slots + by group)
- Cancel any booking
- Reset all bookings

## Change settings
At the top of `app.py`: `NUM_GROUPS`, `SLOT_MINUTES`, `WEEKS`, `OFFICE_HOURS`.
