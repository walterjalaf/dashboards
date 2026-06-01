import urllib.request
import urllib.parse
import json

api_key = 'AIzaSyAduvvGyXxyntmoMJQR_pnczbqatXW8q0c'
file_id = '1fYH5RByHm428L0FowKBPnVu6TX6a-hejdhriqgCFybc' # FEB

ranges = "ranges=" + urllib.parse.quote("DATA STUDIO") + "&ranges=" + urllib.parse.quote("ESTADOS") + "&ranges=" + urllib.parse.quote("ECONOMICO")
url = f"https://sheets.googleapis.com/v4/spreadsheets/{file_id}/values:batchGet?{ranges}&valueRenderOption=UNFORMATTED_VALUE&key={api_key}"

try:
    req = urllib.request.urlopen(url)
    data = json.loads(req.read())
    print(json.dumps([v.get('range', 'NO_RANGE') for v in data.get('valueRanges', [])]))
    
    for v in data.get('valueRanges', []):
        if 'DATA STUDIO' in v.get('range', ''):
            print("DATA STUDIO Sample:", v.get('values', [])[:5])
        
except Exception as e:
    if hasattr(e, 'read'):
        print("Error:", e.read().decode('utf-8'))
    else:
        print("Error:", e)
