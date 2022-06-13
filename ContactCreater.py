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
        self.county = self.getCounty()
        self.state = self.getState()
        self.city = self.getCity()

        self.mobileNumber = mobileNumber
        # self.personas = self.matchPersona(personas) 
        self.personas = personas
        self.industry = industry 
        self.email = email.lower() 
        self.jobTitle = jobTitle 
        self.revenue = revenue 
        

    def matchPersona(self, persona):
        # match persona type inputted by registrant with hubspot predefinded personas
        personaTypes = ["Entrepreneur/Founder", "Talent", "Investor", "Governement", "Employee", 
                        "Innovation Enabler", "Higher Education"]
        for p in personaTypes:
            if str(persona) in p:
                return p
        
        return persona

    def getCounty(self):
        # uses zip code to find regitrants county
        search = uszipcode.SearchEngine()
        try:
            zipcode = search.by_zipcode(self.zip)
            return zipcode.county
        except AttributeError:
            return "Invalid zip: no county"

    def getCity(self):
        # uses zip code to find city
        search = uszipcode.SearchEngine()
        try:
            zipcode = search.by_zipcode(self.zip)
            return zipcode.city
        except AttributeError:
            return "Invalid zip"

    def getState(self):
        # uses zip code to find state
        search = uszipcode.SearchEngine()
        try:
            zipcode = search.by_zipcode(self.zip)
            return zipcode.state
        except AttributeError:
            return -1

    def toDict(self):
        # seperate registrant information into dict data structure
        contactDict = {"firstName": self.firstName, "lastName": self.lastName, "company": self.company, "zip": self.zip, "county": self.county, 
        "phone number": self.mobileNumber, "personas": self.personas, "industry": self.industry, "email": self.email, 
        "job title": self.jobTitle, "revenue": self.revenue, "state": self.state, "city": self.city}

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

        headerNames = self.getHeaders(sheet)
        # print(headerNames)

        for row in sheet.iter_rows():
            # obtain registrant properties
            firstName = row[headerNames.get("First Name")].value
            lastName = row[headerNames.get("Last Name")].value
            # print(name)
            company = row[headerNames.get("Company")].value
            zip = row[headerNames.get("Zip Code")].value
            number = row[headerNames.get("Mobile number")].value
            personas = row[headerNames.get("Personas")].value
            industry = row[headerNames.get("Industry Selection")].value
            email = row[headerNames.get("Email Address")].value
            jobTitle = row[headerNames.get("Job Title")].value
            revenue = row[headerNames.get("Revenue")].value

            newContact = Contact(firstName, lastName, company, zip, number, personas, industry, 
            email, jobTitle, revenue)

            self.contacts.append(newContact)

            county = newContact.getCounty()
            data = newContact.toDict()
            if county not in self.regions:
                self.regions[county] = []
                self.regions[county].append(data)
            else:
                self.regions[county].append(data)

    def getHeaders(self, sheet):
        # get headers and their location from inputted .xlsx file 
        # pass in the active sheet as parameter and loop through 1st column and map index of header to head in dict
        # by default header location set to index 0
        headers = {"First Name": 0, "Last Name": 0, "Job Title": 0, "Company": 0, "Company Website": 0, "Country Industry": 0, 
        "Zip Code": 0, "Mobile number": 0, "Personas": 0, "Industry Selection": 0, "Revenue": 0, "Email Address": 0}

        value = 0
        for cell in sheet[1]:
            cell = cell.value
            headers[cell] = value 
            value += 1

        return headers

    def toJson(self, file="ContactsByRegion.json"):
        # because registrant data is mapped to a dict then data can easily be made into .json form
        with open(file, 'w') as f:
            json.dump(self.regions, f, indent=4)

    def toExcel(self, file="Output.xlsx"):
        # create .xlsx file of all registrants with added regional information
        for contact in self.contacts:
            self.data.append(contact.toDict())
        df = pd.DataFrame.from_dict(self.data)
        print (df)
        df.to_excel(file)     
        filePath = os.path.abspath(file)
        print("Output file location: "  + filePath)
        return filePath


    def sortByCounty(self, file="ContactsByRegion.xlsx"):
        # sort registrants by region
        # map all registrants to dictionary containing all counties
        # loop through each county in dict and outputted order is sorted by county
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
    # tester function
    x = xlToContact(file=file)
    x.toExcel()
    x.sortByCounty()


def main():
    # run main function to use program in terminal window
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

# main()
# tester()


from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfile 
import time

def GUI():

    ws = Tk()
    ws.title('Bizzabo --> Hubspot')
    ws.geometry('700x400') 

    global inputFile
    def open_file():
        file_path = askopenfile(mode='r', filetypes=[('Excel Files', '*xlsx')])
        if file_path is not None:
            Label(ws, text='File Selected Successfully', foreground='black').grid(row=0, column=2)
            pass
        file_path = str(file_path)
        file_path = file_path[25:len(file_path) - 28]
        global inputFile 
        inputFile = file_path

    def uploadFiles():

        file = inputFile
        x = xlToContact(file)

        pb1 = Progressbar(
            ws, 
            orient=HORIZONTAL, 
            length=300, 
            mode='determinate'
            )
        fileLocation = x.toExcel()
        x.sortByCounty()
        pb1.grid(row=4, columnspan=3, pady=20)
        for _ in range(5):
            ws.update_idletasks()
            pb1['value'] += 20
            time.sleep(0.2)
        pb1.destroy()
        Label(ws, text='Files Uploaded Successfully!', foreground='green').grid(row=4, columnspan=3, pady=10)
        Label(ws, text='Files located in \n' + fileLocation, foreground='green').grid(row=5, columnspan=3, pady=10)
        # os.system(fileLocation[0:len(fileLocation)-11])
            
    adhar = Label(
        ws, 
        text='Upload .xlsx file from Bizzabo: '
        )
    adhar.grid(row=0, column=0, padx=10)

    adharbtn = Button(
        ws, 
        text ='Choose File', 
        command = lambda:open_file()
        ) 
    adharbtn.grid(row=0, column=1)

    upld = Button(
        ws, 
        text='Upload Files', 
        command=uploadFiles
        )
    upld.grid(row=3, columnspan=3, pady=10)

    instructions = Label(
        ws,
        text="Download contact list from bizzabo with the following properties:\n     First Name\n     Last Name\n     Company\n     Zip Code\n     County\n     Mobile Number\n     Personas\n     Industry\n     Email\n     Job Title\n     Employees\n     Revenue"
    )
    instructions.grid(row=4, column=0)


    ws.mainloop()
    
GUI()