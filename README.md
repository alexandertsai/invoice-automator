# invoice-automator
(For Mac) Uses Applescript and Python to automate the generation of custom invoices using your invoice template (Word file), and the sending of these invoices through Outlook.

### Generating the invoices

Steps to use:

1. Open your terminal in VScode or whichever IDE you use.
2. Install requirements!

```bash
pip install -r requirements.txt
# pip3 for Python 3 specifically
```

3. Before running the script, prepare the following a template word document with your desired variables in these {brackets}. For example: 


> Dear {name},
>
> Your invoice this month is {amount}. Please pay by {date}.

4. Prepare a CSV file with your various values of variables. (to create a CSV, go to Excel or Google sheets, key in your data with the headings (name, amount, date), key in the data below, and press download as CSV)

| Name | Amount | Date |
| ---- | --- | ----- |
| John | $60 | 23/5 |
| Walt | $55 | 22/2 |

5. Have a folder where you want to generate your invoices in.
6. Now you're ready! Open `app v1.py` and run the script (F5 is usually the shortcut) and select your files.

### Sending the invoices

7. Set up your `email_list.csv`. This should contain the emails and the company code that references your invoices.

| Company Code | receiver | cc recipients | salutation |
| ----- | -------- | -------------- | ---- |
| APPL | stevejobs@gmail.com | timcook@hotmail.com; warrenbuffet@hotmail.com | Steve |
| WMT | walmartboss@gmail.com | employee1@gmail.com | Mr Wall |

8. Open invoiceSender.scpt (AppleScript).
9. Modify the fields with the emojis next to them. To get a path to your file on Mac: click on the file and press Command+I. The file path is under "where". 
10. Compile with command + k and run with command + r