# ATS
## Contents
[automated-teller-machine simulator](#automated-teller-simulator)

[installing-the-application](#installing-the-application)

[login-names-and-pins](#login-names-and-pins)

[exiting-the-application](#exiting-the-application)


Automated Teller Simulator
------------------------------------------------------------
This project was designed to give the student some code to work with.

The entire project worked, but kept it's data in the text files included in the Data folder

The student was asked to meet with the client (teacher) and discuss changes required, then make and present the changes.

## Change Request Premise
> the "product owner" wanted to change the logins to a database

> project was to remove the uneeded code and allow the login form to check a central database (MSAccess) for login, 
presumably so that mutliple machines could access an "online" database for login information.

## What's new in the latest version (Version #2.0)
- Data files are copied to the ProgramData folder where programs are also allowed to write, due to the code writing a new file then deleting the old one, and renaming the new.  This file operation requires delete permissions, and win10 does not allow that in the ProgramFiles folders.

- Data filenames are set dynamically, depending on the system drive. 

- added an installer using Package & Deploy Wizard

- a couple of minor bugs in the transaction file interaction for reports.

## What what changed originally
- database file (MDF) was kept in the program files folder and accessed using an ODBC connection string to find the file, as the file would be in the same place on every computer it was run on.
- frmPINS would access the database instead of a PINS.txt file when checking login information.


# Installing the application
Run ATS_Setup.exe
- it's a self-extracting zipfile, and will unpack to the temporary folder, and then execute the setup.exe, which will copy the required files to your machine, register them as required.

# Login names and PINs


	NAME			PIN
	----------------------------
	Dallas			D001
	Cann			C001
	Clapton			C002
	Santana			S001
	John			J001
	Stevens			S002
	Lennon			L001
	Vaughan			V001
	Connery			C003
	Steele			S003

# Exiting the Application

	a) Login to the application using the login name Dallas and the PIN D001.
	b) Click FILE > CLOSE ATS.
