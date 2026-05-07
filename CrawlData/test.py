import requests
import html
import re
import time
from requests.exceptions import RequestException


# Student prefix: the year of the student ID
student_prefix = "390"

# Path to the CSV file
filePath = 'NN' + str(student_prefix) + '.csv'
logPath = 'log' + str(student_prefix) + '.txt'
errorlog = "error_log.txt"

# Save rate: save data every 10 students
SAVE_RATE = 10
SKIP_RATE = 100  # Skip rate for consecutive "NOT FOUND" results

try_time = 5  # Number of retry attempts for each student ID
losts = []  # List to store student IDs that could not be processed

url = 'https://diemthi.hcm.edu.vn/2018/Home/Show'

headers = {
    'Content-Type': 'application/x-www-form-urlencoded',
    'Cookie': '__RequestVerificationToken_LzIwMTg1=7syHtgfx4c4ERDKLNsc1J12E6QJsZHL6yKpbKFPNHn8EVU44uG3_TY2334cvvoRGcT9kG96f1MAQnxERI59c6coIfQXhwZI9LgVlZMwMEUY1',
}

data = {
    '__RequestVerificationToken': 'NkAu0icXFOtpDq73BkKh-4DkmmNbwrKBXzRpmynCMdjVWfocv0BYX7YO86jbCiSn_2VVpad-2eoDrhTMBkEq3SLcYnhj3ALCHKH0QN3c8OU1',
    'SoBaoDanh': 39000001,
}

response = requests.post(url, headers=headers, data=data)
response.encoding = 'utf-8'

decoded_text = html.unescape(response.text)

with open(filePath, 'w', encoding='utf-8') as file:
    file.write(decoded_text)
    file.flush()
    file.close()