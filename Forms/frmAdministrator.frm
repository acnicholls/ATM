VERSION 5.00
Begin VB.Form frmAdministrator 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Teller Simulator Administrator Menu"
   ClientHeight    =   4005
   ClientLeft      =   6330
   ClientTop       =   6420
   ClientWidth     =   7380
   Icon            =   "frmAdministrator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuFileCloseATS 
         Caption         =   "&Close ATS"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuMoney 
      Caption         =   "&Money"
      Begin VB.Menu mnuMoneyFillUp 
         Caption         =   "&Fill Up"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportsTransactions 
         Caption         =   "&Transactions"
         Begin VB.Menu mnuReportsTransactionsDate 
            Caption         =   "By &Date ..."
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuReportsTransactionsAccount 
            Caption         =   "By &Account Number ..."
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuReportsTransactionsAll 
            Caption         =   "All &Transactions"
            Shortcut        =   ^T
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "?"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Aotomatic Teller simulator ..."
      End
   End
End
Attribute VB_Name = "frmAdministrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'set the shutdown boolean to false, as the machine is running
Private blnExitATS As Boolean

Private Sub Form_Unload(Cancel As Integer)
    'if admin wants to shut down machine then shut down
    If blnExitATS = True Then
        End
    Else
        'unload admin menu, ask for user assurance
        gstrMessage = "Are you sure you want to quit administrator menu?"
        gstrTitle = "Automatic Teller Simulator ..."
        gintStyle = vbYesNo + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        'stop the unload of the form if admin selects no
        If gintAnswer = vbNo Then
            Cancel = True
        End If
    End If

End Sub

Private Sub mnuFileCloseATS_Click()
    'menu selection of shutdown to true, and unload form calling the function that will end the program
    blnExitATS = True
    Unload Me

End Sub

Private Sub mnuFileExit_Click()
    'only close down the admin menu and leave the machine running
    blnExitATS = False
    Unload Me

End Sub

Private Sub mnuMoneyFillUp_Click()
    'declare variables for the sub
    Dim blnFound As Boolean
    Dim dtmToday As Date
    Dim strDate As String
    Dim intBalance As Integer
    Dim intFileHandle As Integer
    'set testing boolean, and today's date
    blnFound = False
    dtmToday = Date
    'open file and read teller's balance for today's date
    intFileHandle = FreeFile
    Open gstrDailyBalanceFile For Input As #intFileHandle
    Do While Not EOF(intFileHandle)
        Input #intFileHandle, strDate, intBalance
        'test date on file for equality to today's date
        If Format(strDate, "mm/dd/yyyy") = Format(dtmToday, "mm/dd/yyyy") Then
            blnFound = True
            Exit Do
        End If
    Loop
    Close #intFileHandle
    'if the teller has at least 5000, then ask confirmation message of admin
    If intBalance > 5000 Then
        gstrMessage = "There is already a minimum of $5,000 avalaible, do you want to continue?"
        gstrTitle = "Automatic Teller Simulator ... Warning."
        gintStyle = vbYesNo + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        If gintAnswer = vbNo Then
            Exit Sub
        End If
    End If
    'if the teller has over 20000 then report to admin and cancel
    If intBalance = 20000 Then
        gstrMessage = "The machine has the maximum of $20,000"
        gintStyle = vbOKOnly
        gstrTitle = "Automatic Teller Simulator ... Warning."
        MsgBox gstrMessage, gintStyle, gstrTitle
        Exit Sub
    End If
    intBalance = intBalance + 5000
    'if the teller goes over 20000 then report to admin
    If intBalance >= 20000 Then
    'only allow the machine to be filled with 20000...
        intBalance = 20000
        'send transaction report to file
        WriteAdministratorTransaction ("R")
        'tell admin
        gstrMessage = "The machine is filled to the maximum of $20,000"
        gintStyle = vbOKOnly + vbCritical
        gstrTitle = "Automatic Teller Simulator ... Warning."
        MsgBox gstrMessage, gintStyle, gstrTitle
        'make global variable for teller balacne equal to 20000
        gintTellerDailyBalance = intBalance
        'send new teller balance to file
        SaveTellerBalance
        Exit Sub
    End If
    'if all is well then allow the operation to continue
    'add 5000 to the daily balance
    gintTellerDailyBalance = intBalance
    'call the write function and send the parameter value
    WriteAdministratorTransaction ("R")
    'call to write new teller balance to file
    SaveTellerBalance
    'send message to admin reporting success
    gstrMessage = "An amount of "
    gstrMessage = gstrMessage & Format(5000, "currency")
    gstrMessage = gstrMessage & " had been added to today's balance."
    gstrTitle = "Automatic Teller Simulator ... Fill Up Report."
    gintStyle = vbOKOnly
    gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
    
End Sub

Private Sub mnuReportsTransactionsAccount_Click()
    'call the report selector and send it the value of account report
    gstrReportType = "Account"
    frmReportSelector.Show vbModal
End Sub

Private Sub mnuReportsTransactionsAll_Click()
    'print report header
    PrintReportHeader
    'print the contents of the transaction file
    PrintAllTransactions
    'send the report to printer
    Printer.EndDoc
    'notify admin the report is ready
    gstrMessage = "Your report is available at the printer."
    gstrTitle = "Printing completed ..."
    gintStyle = vbOKOnly
    MsgBox gstrMessage, gintStyle, gstrTitle
    
End Sub

Private Sub mnuReportsTransactionsDate_Click()
    'call the report selector and send it the value of date report
    gstrReportType = "Date"
    frmReportSelector.Show vbModal
End Sub
