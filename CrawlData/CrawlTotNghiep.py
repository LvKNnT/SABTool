import requests
import html
import re
import time
from requests.exceptions import RequestException


# Student prefix: the year of the student ID
student_prefix = "020"

# Path to the CSV file
filePath = 'NN' + str(student_prefix) + '.csv'
logPath = 'log' + str(student_prefix) + '.txt'
errorlog = "error_log.txt"

# Save rate: save data every 10 students
SAVE_RATE = 10

try_time = 5  # Number of retry attempts for each student ID
losts = []  # List to store student IDs that could not be processed

def getDataOfStudent(studentId, timeout=10):
    for i in range(try_time):
        try:
            url = 'https://diemthi.hcm.edu.vn/2018/Home/Show'

            headers = {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Cookie': '__RequestVerificationToken_LzIwMTg1=7syHtgfx4c4ERDKLNsc1J12E6QJsZHL6yKpbKFPNHn8EVU44uG3_TY2334cvvoRGcT9kG96f1MAQnxERI59c6coIfQXhwZI9LgVlZMwMEUY1',
            }

            data = {
                '__RequestVerificationToken': 'NkAu0icXFOtpDq73BkKh-4DkmmNbwrKBXzRpmynCMdjVWfocv0BYX7YO86jbCiSn_2VVpad-2eoDrhTMBkEq3SLcYnhj3ALCHKH0QN3c8OU1',
                'SoBaoDanh': studentId,
            }

            response = requests.post(url, headers=headers, data=data, timeout=timeout)
            response.encoding = 'utf-8'

            decoded_text = html.unescape(response.text)

            if "table" in decoded_text:
                matches = re.findall(
                    r'<tr>\s*<td[^>]*>\s*([^<]+)\s*</td>\s*<td[^>]*>\s*([^<]+)\s*</td>\s*<td[^>]*>\s*([^<]+)\s*</td>',
                    decoded_text, re.DOTALL
                )
                match = matches[1] if len(matches) > 1 else None
                if match:
                    name = match[0].strip()
                    birthday = match[1].strip()
                    result = match[2].strip()
                    print(f"Student ID: {studentId}, Name: {name}, Birthday: {birthday}, Result: {result}")
                    return [studentId, name, birthday, result]

        except RequestException as e:
            print(f"[Timeout/Error] {studentId} attempt {i+1}: {e}")
            time.sleep(1)  # Wait before retrying

    # Add fail attempt to losts
    with open(logPath, mode='a', encoding='utf-8') as log:
        log.write(f"Student ID {studentId} could not be processed after {try_time} attempts.\n")
        log.flush()

    losts.append(studentId)

    print(f"Student ID {studentId} could not be processed after {try_time} attempts.")
    return [studentId, "NOT FOUND", "NOT FOUND", "NOT FOUND"]

# Clear the log file
with open(logPath, mode='w', encoding='utf-8') as log:
    log.write("Student groups:\n")
    log.flush()
    log.close()

# Search for valid student groups in the range of 00 to 99
student_groups = []
# # checking mode
# for i in range(100):
#     studentId = student_prefix + str(i).zfill(3) + "01"
#     result = getDataOfStudent(studentId, 2)

#     if result[1] == "NOT FOUND":
#         continue

#     # Check if the student ID is valid
#     student_groups.append(i)
#     with open(logPath, mode='a', encoding='utf-8') as log:
#         log.write(f"{i}\n")
#         log.flush()
#         log.close()
# brute force mode
student_groups = list(range(100))
# student_groups = [i for i in range(6,100)]

with open(logPath, mode='a', encoding='utf-8') as log:
    log.write("Brute force mode\n")
    log.flush()
    log.close()

# Create a list of student IDs with the first two digits as the value of student_prefix
student_ids = [f"{student_prefix}{group:02d}" for group in student_groups]

# Open the CSV file and write data into it
with open(filePath, mode='w', newline='', encoding='utf-8') as file:
    file.write("StudentID,Name,Birthday,Result\n")
    
    # Search for and write data for each thousand IDs
    for prefix in student_ids:
        count = 0

        for i in range(1000):
            studentId = prefix + str(i).zfill(3)

            result = getDataOfStudent(studentId,3)

            if result[1] == "NOT FOUND":
                count += 1
                if count == 100:  # Stop if 100 consecutive IDs are not found
                    break
                continue

            file.write(f"{result[0]},{result[1]},{result[2]},{result[3]}\n")

            if int(studentId) % SAVE_RATE == 0:
                file.flush()

    with open(logPath, mode='a', encoding='utf-8') as log:
        log.write(f"Total student groups: {len(student_groups)}\n")
        log.write(f"Total student IDs: {len(student_ids) * 1000}\n")
        log.write(f"Total losts: {len(losts)}\n")
        log.flush()

    for lost in losts:
        result = getDataOfStudent(lost,5)

        if result[1] != "NOT FOUND":
            file.write(f"{result[0]},{result[1]},{result[2]},{result[3]}\n")
        else:
            with open(logPath, mode='a', encoding='utf-8') as log:
                log.write(f"Student ID {studentId} has request error.\n")
                log.flush()


    file.flush()
    file.close()
