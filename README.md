# Report Automation with Python
This is an automation of operations with excel that I made using Python, it aims to automatically send an email to the recipient containing the Billing, the quantity of products sold and the average ticket per product.

## How to install:
### In your terminal, install the following libraries:
```
pip install pandas
pip install openpyxl
```
### In ```main.py``` file... change the following variables:
```
42. msg['Subject'] = 'Enter the email subject here'
43. msg['From'] = 'Enter sender here'
44. msg['To'] = 'Enter recipient here'
45. password = 'Enter the token here'  (To get the token go to: https://security.google.com/settings/security/apppasswords.)
```
