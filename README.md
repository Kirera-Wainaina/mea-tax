# Mea-tax

## About the project
This project came about as result of manual and repetitive work. An accountant friend of mine was tasked with creating tax files for employees in the organization. 

It all starts from a base file that contains all the records of the employees. It has the salaries and chargeable taxes per month, their identification information such as names, pins and id. This base file is reference in the code as 'records.xlsx'. I can't share it for confidential reasons.

## How it works
I used python to make the script. It has a package, [Openpyxl](https://openpyxl.readthedocs.io/en/stable/index.html), which is great for manipulating excel documents.

Using Openpyxl, I retrieved the employee data from records.xlsx. I then created each employee's file from a [template](./static/p9.xlsx). I then added the data into the file.

When this is done, I used headless Libreoffice to convert the xlsx files to pdf format.

All of this is in [main.py](./main.py).

## Other Information
Prior versions of this program did not incorporate tests. As an engineer, it is important to test your code. Turns out it is an important antidote to anxiety. You are certain your code will do what it is supposed to do. Check out the [test file](./test.py) for the tests.

'Ancora imparo'