# Remember to pip install selenium before running the code
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Path to WebDriver
# Download the Edge WebDriver from https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
# If using a different browser, search for its WebDriver. Ensure compatibility between the WebDriver and the browser version.
# Extract and replace the path here
msdriver = "D:/Downloads/edgedriver_win64/msedgedriver.exe"

# Save rate: save data every 10 students
SAVE_RATE = 10

# Student prefix: the year of the student ID
student_prefix = "22"

# Path to the CSV file
filePath = 'NN' + str(student_prefix) + '.csv'
logPath = 'log' + str(student_prefix) + '.txt'

# Initialize WebDriver
service = Service(executable_path=msdriver)
driver = webdriver.Edge(service=service)
wait = WebDriverWait(driver, 10)

try_time = 5  # Number of retry attempts for each student ID
losts = []  # List to store student IDs that could not be processed

def getDataOfStudent(studentId, delay=10):
    try:
        # Wait for and click into the input field
        inputElement = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "topSearchInput"))
        )

        # Repeat typing the student ID until the value matches
        for _ in range(try_time):
            # Clear old input first
            inputElement.send_keys(Keys.CONTROL + "a")
            inputElement.send_keys(Keys.BACKSPACE)

            inputElement.send_keys(f"{studentId}")
            
            # Ensure the value in the class "uz227 GgHzM" matches the current student ID
            def value_matches(driver):
                try:
                    element = driver.find_element(By.ID, "topSearchInput")
                    return str(studentId) in element.get_attribute("value")
                except:
                    return False

            if WebDriverWait(driver, delay).until(value_matches):
                break

            print(f"Attempt {_ + 1} for student ID {studentId} failed. Retrying...")

            if _ == try_time - 1:
                losts.append(studentId)
                with open(logPath, mode='a', encoding='utf-8') as log:
                    log.write(f"Student ID {studentId} could not be processed after {try_time} attempts.\n")
                    log.flush()

        # Wait for the suggestion to appear AND be updated for this studentId
        def suggestion_matches(driver):
            try:
                element = driver.find_element(By.CLASS_NAME, "erWUf")
                return str(studentId) in element.get_attribute("aria-label")
            except:
                return False

        WebDriverWait(driver, delay).until(suggestion_matches)

        # Now safely extract info
        accountElement = driver.find_element(By.ID, "searchSuggestion-0")
        full_text = accountElement.get_attribute("aria-label")
        info = full_text.replace("People Suggestion - ", "")
        name, email = info.rsplit(" ", 1)

        print(f"Student ID: {studentId}, Name: {name}, Email: {email}")

        # Click on the suggestion to open the profile
        return [studentId, name, email]

    except Exception as e:
        print(f"Error for student ID {studentId}: {e}")
        return [studentId, "NOT FOUND", "NOT FOUND"]

# Open the website
driver.get("https://outlook.office365.com/mail/")
print("Please log in manually. Press Enter when done.")
input()  # Waits for user to press Enter before proceeding

# Clear the log file
with open(logPath, mode='w', encoding='utf-8') as log:
    log.write("Student groups:\n")
    log.flush()
    log.close()

# Search for valid student groups in the range of 000 to 999
student_groups = []
for i in range(1000):
    studentId = student_prefix + str(i).zfill(3)
    result = getDataOfStudent(studentId, 3)

    if result[1] == "NOT FOUND":
        continue

    # Check if the student ID is valid
    student_groups.append(i)
    with open(logPath, mode='a', encoding='utf-8') as log:
        log.write(f"{i}\n")
        log.flush()
        log.close()

# Create a list of student IDs with the first two digits as the value of student_prefix
student_ids = [f"{student_prefix}{group:03d}" for group in student_groups]

with open(logPath, mode='a', encoding='utf-8') as log:
    log.write("Lost student IDs:\n")
    log.flush()
    log.close()

# Open the CSV file and write data into it
with open(filePath, mode='w', newline='', encoding='utf-8') as file:
    file.write("StudentID,Name,Email\n")
    
    # Search for and write data for each thousand IDs
    for prefix in student_ids:
        count = 0

        for i in range(1000):
            studentId = prefix + str(i).zfill(3)

            result = getDataOfStudent(studentId, 3)

            if result[1] == "NOT FOUND":
                count += 1
                if count == 100:  # Stop if 100 consecutive IDs are not found
                    break
                continue

            file.write(f"{result[0]},{result[1]},{result[2]}\n")

            if int(studentId) % SAVE_RATE == 0:
                file.flush()

    file.flush()
    file.close()

driver.quit()

# Run the script using the terminal: python CrawlStudentData.py
