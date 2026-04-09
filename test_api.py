import requests

url = "http://localhost:8000/api/process"
files = {'file': open('data.xlsx', 'rb')}

try:
    response = requests.post(url, files=files)
    if response.status_code == 200:
        with open('test_output.xlsx', 'wb') as f:
            f.write(response.content)
        print("Success! Processed file saved as test_output.xlsx")
    else:
        print(f"Error: {response.status_code}")
        print(response.json())
except Exception as e:
    print(f"Request failed: {e}")
