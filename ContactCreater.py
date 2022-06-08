"""

Required Fields for contact:
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

import uszipcode, openpyxl, os, json
import pandas as pd

# import tkinter as tk
# from tkinter import filedialog

class Contact:
    # definition for contact with 
    def __init__(self, firstName, lastName, company, zip, mobileNumber, personas, industry, email, jobTitle, revenue):
        self.firstName = firstName
        self.lastName = lastName
        # self.fullName = f"{self.firstName} {self.lastName}"
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
            return -1

    def toDict(self):
        contactDict = {"firstName": self.firstName, "lastName": self.lastName, "company": self.company, "zip": self.zip, "county": self.getCounty(), 
        "phone number": self.mobileNumber, "personas": self.personas, "industry": self.industry, "email": self.email, 
        "job title": self.jobTitle, "revenue": self.revenue}

        return contactDict


class xlToContact:

    def __init__(self, file):
        self.file = file # takes in input file 
        self.contacts = []
        self.data = [] 
        self.sortedDataByRegion = []
        self.regions = {}

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

            county = newContact.getCounty()
            data = newContact.toDict()
            if county not in self.regions:
                self.regions[county] = []
                self.regions[county].append(data)
            else:
                self.regions[county].append(data)

    def toJson(self, file="ContactsByRegion.json"):
        with open(file, 'w') as f:
            json.dump(self.regions, f, indent=4)

    def toExcel(self, file="Output.xlsx"):
        for contact in self.contacts:
            self.data.append(contact.toDict())
        df = pd.DataFrame.from_dict(self.data)
        print (df)
        df.to_excel(file)     
        filePath = os.path.abspath(file)
        print("Output file location: "  + filePath)

    def sortByCounty(self, file="ContactsByRegion.xlsx"):
        for region in self.regions:
            for contact in self.regions[region]:
                self.sortedDataByRegion.append(contact)
        df = pd.DataFrame.from_dict(self.sortedDataByRegion)
        df.to_excel(file)

    def getCountyNumbers(self, county):
        # pass county name and return number of contacts from designated region
        formattedString = ""
        for i in range(len(county)):
            if i == 0:
                formattedString += county[i].upper()
            elif county[i - 1] == ' ':
                formattedString += county[i].upper()
            else:
                formattedString += county[i]
        if " County" not in formattedString:
            formattedString += " County"
        try:
            return formattedString + ": " + str(len(self.regions[formattedString])) + " contacts"
        except:
            return 0


def tester(file="test.xlsx"):
    x = xlToContact(file=file)
    x.toExcel()
    x.sortByCounty()


def main():
    running = True
    convertedToExcel = False
    print("Intructions: \nEnter full path to bizzabo .xlsx file. \nIf unsure of path right click on file, select get info, and the \"where\" section gives the full file path.")
    filePath = input("Enter path to input file: ")
    while running:
        try:
            if not convertedToExcel:
                x = xlToContact(file=filePath)
                x.toExcel()
                convertedToExcel = True
            print("Options:\n1) Type name of county to return number of contacts from designated region")
            print("2) Would you like to sort by region type 'y'")
            print("3) If you are finished type 'q'\n")
            userInput = input("Enter action: ")
            if userInput == 'q':
                running = False
            elif userInput == 'y':
                x.sortByCounty()
                filePath = os.path.abspath("ContactsByRegion.xlsx")
                print("Sorted contacts by region file location: " + filePath)
            else:
                print(x.getCountyNumbers(userInput))
        except:
            print("Invalid file name. ")

main()
# tester()


            
            