import urllib.request
import json

api_key = 'AIzaSyAduvvGyXxyntmoMJQR_pnczbqatXW8q0c'
folder_id = '18dbLxdOuAe8koxgPJ6IvmLpEZer4X7-e'
url = f"https://www.googleapis.com/drive/v3/files?q='{folder_id}'+in+parents+and+trashed=false&key={api_key}&fields=files(id,name,mimeType)"

try:
    req = urllib.request.urlopen(url)
    data = json.loads(req.read())
    print(json.dumps(data, indent=2))
except Exception as e:
    print("Error:", e)
