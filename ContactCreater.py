"""
Required Fields for contact

First Name
Last Name
Company
Zip Code
County (work your magic)
Mobile Number
Personas
Industry
Email
Job Title
(PROBABLY) Job Level - need to have some internal discussion here
Employees
Revenue


"""

import uszipcode, openpyxl, csv
import pandas as pd
from xlsxwriter import Workbook


class Contact:

    def __init__(self, firstName, lastName, company, zip, mobileNumber, personas, industry, email, jobTitle, revenue):
        self.firstName = firstName
        self.lastName = lastName
        self.fullName = f"{self.firstName} {self.lastName}"
        self.company = company
        self.zip = zip
        self.mobileNumber = mobileNumber
        self.personas = personas 
        self.industry = industry 
        self.email = email 
        self.jobTitle = jobTitle 
        self.revenue = revenue 

    def getCounty(self):
        # obtain registrant county data from user inputted zipcode
        search = uszipcode.SearchEngine()
        try:
            zipcode = search.by_zipcode(self.zip)
            return zipcode.county
        except AttributeError:
            return "Invalid zip: no county"

    def getCity(self):
        search = uszipcode.SearchEngine()
        try:
            zipcode = search.by_zipcode(self.zip)
            return zipcode.city
        except AttributeError:
            return "Invalid zip"

    def getState(self):
        search = uszipcode.SearchEngine()
        try:
            zipcode = search.by_zipcode(self.zip)
            return zipcode.state
        except AttributeError:
            return "Invalid zip"

    def toDict(self):
        contactDict = {"name": self.fullName, "company": self.company, "zip": self.zip, "county": self.getCounty(), 
        "phone number": self.mobileNumber, "personas": self.personas, "industry": self.industry, "email": self.email, 
        "job title": self.jobTitle, "revenue": self.revenue}

        return contactDict


class xlToContact:

    def __init__(self, file):
        self.file = file # takes in input file 
        self.contacts = [] 

        file = openpyxl.load_workbook(self.file)
        sheet = file.active 

        for row in sheet.iter_rows():
            name = (row[0].value, row[1].value) # firstname, lastname
            # print(name)
            company = row[2].value
            zip = row[5].value
            number = row[6].value
            personas = row[7].value
            industry = row[9].value
            email = row[10].value
            jobTitle = row[11].value
            revenue = row[13].value

            newContact = Contact(name[0], name[1], company, zip, number, personas, industry, 
            email, jobTitle, revenue)

            self.contacts.append(newContact)
    
    def toExcel(self):
        data = []
        for contact in self.contacts:
            data.append(contact.toDict())
        df = pd.DataFrame.from_dict(data)
        print (df)
        df.to_excel('Output.xlsx')     


running = True
while running:
    print("Type 'q' to quit.")
    userInput = input("Enter path to input file: ")
    if userInput == 'q':
        running = False
    else:
        x = xlToContact(userInput)
        print("1: Convert input file to excel")
        userInput = input("\nEnter command: ")
        if userInput == '1':
            x.toExcel()
