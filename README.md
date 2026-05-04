# AGE_24 Web Processor

A local web app that processes your AGE_24 Excel files.
One person runs the server — everyone on the same WiFi uses it via browser.

---

## SETUP (one-time, on the host PC)

**Requirements:** Python 3.9+ installed on the host PC.

1. Place all files in a folder:
   ```
   age24_webapp/
   ├── app.py
   ├── START_SERVER.bat
   ├── README.md
   └── templates/
       └── index.html
   ```

2. Double-click **START_SERVER.bat**
   - It installs Flask, pandas, openpyxl automatically
   - It starts the server and prints the WiFi address

3. You'll see something like:
   ```
   Open in browser on this PC : http://localhost:5000
   Open from other PCs on WiFi: http://192.168.1.42:5000
   ```

4. Share the WiFi address (e.g. `http://192.168.1.42:5000`) with your team.

---

## HOW TO USE (anyone on the network)

1. Open the address in any browser (Chrome, Edge, Firefox)
2. Drag & drop your `.xlsx` file (or click to browse)
3. Click **PROCESS FILE**
4. Review the summary table
5. Click **DOWNLOAD UPDATED EXCEL** to save the result

---

## STOPPING THE SERVER

Close the command window (or press Ctrl+C in it).

---

## TROUBLESHOOTING

- **"No raw data sheet found"** → The data sheet tab must be named MM-DD-YYYY (e.g. `04-06-2026`)
- **"Date not found in Summary"** → The Summary sheet's header row must include the same date
- **Coworkers can't connect** → Check Windows Firewall; allow Python on private networks
- **Wrong IP address shown** → Open Command Prompt and run `ipconfig` to find your IPv4 address
