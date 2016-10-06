import sqlite3
import wx
import datetime
import re
import xlwt

class invalidEmailError(Exception):
    pass

class invalidFirstNameError(Exception):
    pass

class invalidLastNameError(Exception):
    pass
    
class invalidPhoneNoError(Exception):
    pass
    
class SubmitButton(wx.Button):    
    def __init__(self, par):
        wx.Button.__init__(self, par, wx.ID_OK, "Confirm", (275, 400))

class viewLogRecordButton(wx.Button):
    def __init__(self, par):
        wx.Button.__init__(self, par, label = "View Log Record", pos = (125,400))
        
class LogFrame(wx.Frame):
    def __init__(self):
        self.Employees = ['John Doe', 'Jane Doe', 'Jeffrey Zhang', 'Keith Chung', 'Jason Ye', 'Wesley Lim', 'Bob Richards', 'Nina Williams', 'Dean Earwicker', 'Don Burdette', 'Tim Kristian', 'KC Walker', 'Wesley Taylor']        
        wx.Frame.__init__(self, None, -1, 'Information Confirmed')
        self.execute()
        
    def execute(self):
        self.panel = wx.Panel(self, -1) 
        self.askForInfo(self.checkFirstName)       
        self.askForInfo(self.checkLastName)
        self.askForInfo(self.checkEmail)
        self.askForInfo(self.checkPhoneNo)
        self.askForInfo(self.checkPersonVisited)
        self.EnteredName = self.FirstName + " " + self.LastName
        self.dateTime = str(datetime.datetime.now())
        wx.StaticText(self.panel, -1, 'Name: '+self.EnteredName, (15,10))
        wx.StaticText(self.panel, -1, 'Email: '+self.Email, (15,50))        
        wx.StaticText(self.panel, -1, 'Phone Number: '+self.PhoneNo, (15,90))
        wx.StaticText(self.panel, -1, 'Person you are Visited: '+self.PersonVisited, (15,130))
        wx.StaticText(self.panel, -1, 'Date and Time: ' + self.dateTime, (15, 170))
        Submit = SubmitButton(self.panel)
        self.Bind(wx.EVT_BUTTON, self.close, Submit)
        
    def getFirstName(self):
        return self.FirstName
    
    def getLastName(self):
        return self.LastName 
           
    def getEmail(self):
        return self.Email
    
    def getPhoneNo(self):
        return self.PhoneNo
    
    def getPersonVisited(self):
        return self.PersonVisited
    
    def getDateTime(self):
        return self.dateTime
    
    def checkEmail(self):
        self.EmailBox = wx.TextEntryDialog(None, 'Enter your email: ', '', '123@yahoo.com')
        if self.EmailBox.ShowModal() == wx.ID_OK:
            self.Email = self.EmailBox.GetValue()
        if not re.match('[^@]+@[^@]+\.[^@]+', self.Email):
            raise invalidEmailError
                
    def checkFirstName(self):
        self.firstNameBox = wx.TextEntryDialog(None, 'Enter your first name: ', '', 'John')
        if self.firstNameBox.ShowModal() == wx.ID_OK:
            self.FirstName = self.firstNameBox.GetValue()
        if self.FirstName.isalpha()==False:
            raise invalidFirstNameError 
            
    def checkLastName(self):
        self.lastNameBox = wx.TextEntryDialog(None, 'Enter your last name: ', '', 'John')
        if self.lastNameBox.ShowModal() == wx.ID_OK:
            self.LastName = self.lastNameBox.GetValue()
        if self.LastName.isalpha()==False:
            raise invalidLastNameError
                
    def checkPhoneNo(self):
        self.PhoneNoBox = wx.TextEntryDialog(None, 'Enter your phone number (ie 1234567890): ', '', '6504567890')
        if self.PhoneNoBox.ShowModal() == wx.ID_OK:
           self.PhoneNo = self.PhoneNoBox.GetValue()        
        int(self.PhoneNo)
        if len(self.PhoneNo)!=10:
            raise invalidPhoneNoError
    
    def checkPersonVisited(self):        
        self.PersonVisitedBox = wx.SingleChoiceDialog(None, 'Select the person you are visiting: ', '', self.Employees)
        if self.PersonVisitedBox.ShowModal() == wx.ID_OK:
            self.PersonVisited = self.PersonVisitedBox.GetStringSelection()
            
    def askForInfo(self, checkInfo):      
        def run(self):
            exception = True
            while exception == True:
                try:
                    checkInfo()
                    exception = False
                except invalidFirstNameError:
                    self.box = wx.MessageDialog(None, 'The first name given is invalid! It must only consist of alphabetical letters! Please re-enter your first name', 'Error!', wx.OK)
                    self.box.ShowModal()
                except invalidLastNameError:
                    self.box = wx.MessageDialog(None, 'The last name given is invalid! It must only consist of alphabetical letters! Please re-enter your last name', 'Error!', wx.OK)
                    self.box.ShowModal()
                except invalidEmailError:
                    self.box = wx.MessageDialog(None, 'The email given is invalid! It must be in the format username@domain.com! Please re-enter your email', 'Error!', wx.OK)
                    self.box.ShowModal()
                except invalidPhoneNoError:
                    self.box = wx.MessageDialog(None, 'The phone number given is invalid! It must only consist of numbers, start with the area code, and be of 10 digits! Please re-enter your phone number', 'Error!', wx.OK)
                    self.box.ShowModal()
                except ValueError:
                    self.box = wx.MessageDialog(None, 'The phone number given is invalid! It must only consist of numbers, start with the area code, and be of 10 digits! Please re-enter your phone number', 'Error!', wx.OK)
                    self.box.ShowModal()
        return run(self)
                
    def viewRecord(self, event):
        self.allLines = self.cursor.execute("select firstName, lastName and dateTime from visitorlog")
        for lines in self.allLines:
            print lines[0]+' '+lines[1]+', '+lines[2]
        
    def addToDB(self):
        self.db = sqlite3.connect("testing.db" )
        self.cursor = self.db.cursor()
        self.cursor.execute('''create table if not exists visitorlog(
            firstName text, 
            lastName text,
            email text, 
            phoneNo char(10), 
            personVisited text,
            dateTime text)''')
        self.sql = "insert into visitorlog values('"+self.getFirstName()+"', '" + self.getLastName() + "', '"  + self.getEmail() + "', '"  + self.getPhoneNo()+ "', '"  + self.getPersonVisited()+"', '"  + self.getDateTime()+"')"
        self.cursor.execute(self.sql)
        self.allLines = self.cursor.execute("select firstName, lastName and dateTime from visitorlog")
        for lines in self.allLines:
            print lines[0], ' ', lines[1], ', ', lines[2]

    def close(self, event):
        self.Close(True)
               
def addToDB(logFrame):
    db = sqlite3.connect("registration.db" )
    cursor = db.cursor()
    cursor.execute('''create table if not exists visitorlog(
        firstName text, 
        lastName text,
        email text, 
        phoneNo char(10), 
        personVisited text,
        dateTime text)''')
    sql = "insert into visitorlog values('"+logFrame.getFirstName()+"', '" + logFrame.getLastName() + "', '"  + logFrame.getEmail() + "', '"  + logFrame.getPhoneNo()+ "', '"  + logFrame.getPersonVisited()+"', '"  + logFrame.getDateTime()+"')"
    cursor.execute(sql)
    db.commit()
    allLines = cursor.execute('select * from visitorlog')
    return allLines
        
def writeToExcelFile(allLines):
    reg = xlwt.Workbook()
    sheet = reg.add_sheet('Sheet 1')
    g = ['First Name', 'Last Name', 'Email', 'Phone Number', 'Person Visited', 'Date and Time']

    for i in range(len(g)):
        sheet.write(0,i,g[i])
    i = 1
    for lines in allLines:
        sheet.write(i, 0, lines[0])
        sheet.write(i, 1, lines[1])
        sheet.write(i, 2, lines[2])
        sheet.write(i, 3, lines[3])
        sheet.write(i, 4, lines[4])
        sheet.write(i, 5, lines[5])
        i+=1
    try:
        reg.save('Visitor Registration Log.xls')
    except:
        box = wx.MessageDialog(None, 'Permission to the file is denied since you are running it in the background. Please close "Visitor Registration Log.xls" and try again', 'Error!', wx.OK)
        box.ShowModal()
    
log = wx.App()
logFrame = LogFrame()
logFrame.Show()
allLines = addToDB(logFrame)
writeToExcelFile(allLines)
log.MainLoop()


