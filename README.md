# Whatsapp Messaging Automation


This program will run and send whatsapp messages from excel sheet devs.xslx and read all the names and phone numbers and email and send a custom message to each row in that excel file to the corresponding whatsapp phone number on that sheet


Instructions:

1. must best logged in on your whatsapp web on the default browser
2. update the message if you wish
3. update the excel files with valid phone numbers at least, update names or emails if you wish 
4. Must have python 3.9 installed 
5. Must have pip installed
6. Navigate to the project folder from command line
```commandline
cd path/to/project/folder
```
4. Install all the required packages (run for the first time only)

```commandline
pip install -r requirements.txt
```
5. run the script
```commandline
py main.py
```

Notes:
- handle errors in case something went bad and repeat sending the message to that number in case not sent, then from excel mark as done in a column
- only send messages to names which doesn't have a column marked as done
- format texts:
  - names to be caplitalized 
  - emails to be lower letter
  - phone with '+' sign if not existed



```python 
wh = pyautogui.size()
```