# Invoice Generator and automated Email sending

This is a project I worked on whilst teaching where I was required to bill the parents of my students for the lessons they had in a term. The invoices were billed termly and had two possible rates of $15 for a half hour lesson in a group or $30 for a half hour lesson as an indiviudual. This program takes three arguments to run:

### CSV File

The file containing the information about each student. It is formatted as follows:

| Child Name       | Lesson Type          | Instrument           | Skill     | Email                      | Parent          | School                | Lessons Completed      | Term          |
| ---------------- | -------------------- | -------------------- | --------- | -------------------------- | --------------- | --------------------- | ---------------------- | ------------- |
| <Student's Name> | <Individual / Group> | <Trumpet / Trombone> | <N/A>     | <Parent's email address>   | <Parent's Name> | <Hukanui / Fairfield> | <No. of weeks in term> | <No. of term> |
| Campbell Smith   | Individual           | Trumpet              | 15+ years | campbellsmith116@gmail.com | Campbell Smith  | Hukanui               | 10                     | 1             |

These categories are filled in for each student and is exported as a .csv file when the invoices are needed to be generated.

### User Email

The email you want to send the invoices from. I used my personal email address.

### User Password

I needed to create an app password for my google account in order for this to work. Instructions about how to create an app password can be found here: https://support.google.com/accounts/answer/185833?hl=en

![Screenshot 2024-06-01 171242](https://github.com/cs1522561/invoice_generator/assets/91705168/7017a550-920a-49dc-8136-e891485c6b25)


When creating an app password it is generated in the format `XXXX XXXX XXXX XXXX`. When including the app password in the arguments it needs to have the spaces removed to be `XXXXXXXXXXXXXXXX`.


## Running the program

Two libraries need to be installed in order for this program to run. 

* FPDF - generates the invoice using cells to format as a table. Saves the invoice as a pdf in the same folder.
* Pandas - Extracts the data from the csv in an easy to use and manipulate format.

This is done by writing the following line into a console, in my case this was done in Visual Studio Code after installing python through the microsoft store.

```
pip install fpdf pandas
```

To run the program, navigate to the folder where `invoice_generator.py` is located. Ensure the .csv file is in the same folder for ease of running. Assuming the .csv is in the same folder the command to run the program is as follows:

```
python .\invoice_generator.py .\<name of csv>.csv <emailaddress> <apppassword>
```

If you run the program with the included test files the command is as follows:

```
python .\invoice_generator.py .\term1_test.csv youremail@gmail.com XXXXXXXXXXXXXXXX
```

This will create an invoice and save it as a pdf in the same folder, and also send the invoice to the email specificed in the csv file.
