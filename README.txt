Pricing App (ready to run)

This folder contains:
- streamlit_app.py  (the web page)
- run_pricing.py    (the pricing engine)
- requirements.txt  (Python deps)
- RETOOL ALL COST UPLOAD 2026 WITH FUEL TYPE.xlsx (rates)

Run locally (Windows or Mac)
1) Install Python 3.10+
2) In a terminal, cd into this folder, then run:

   pip install -r requirements.txt
   streamlit run streamlit_app.py

Host it (staff-only)
- Streamlit Cloud or Render work well.
- Set an environment variable APP_PASSWORD to require a password.
- Updating rates centrally = replace the rates xlsx in the hosted app and redeploy.
