VERSION 5.00
Begin VB.Form frmReportSelector 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1785
   ClientLeft      =   -2220
   ClientTop       =   9210
   ClientWidth     =   7410
   Icon            =   "frmReportSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox cboSelection 
      Height          =   315
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1080
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3060
   End
End
Attribute VB_Name = "frmReportSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'declare transaction file variables and form wide testing variables, counters
    Dim strDate As String
    Dim lngTransactionID As Long
    Dim strUserName As String
    Dim strUserAccountNumber As String
    Dim strAccountType As String
    Dim strTransactionCode As String
    Dim curTransactionAmount As Currency
    Dim curUserChequingAccountBalance As Currency
    Dim curUserSavingsAccountBalance As Currency
    Dim intTellerDailyBalance As Integer
    Dim intCounter As Integer

Private Sub cboSelection_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    'unload the report selector and return to admin form
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'set type of report to print and give slection information
    Select Case gstrReportType
        Case "Date"
            'send date selected to global variable
            gstrTransactionReportDate = cboSelection
            'print report header
            PrintReportHeader
            'print report based on date
            PrintDateReport
            'send report to printer
            Printer.EndDoc
        Case "Account"
            'send account number selected to global variable
            gstrTransactionReportAccountNumber = cboSelection
            'print report header
            PrintReportHeader
            'print report based on account number
            PrintAccountReport
            'send report to printer
            Printer.EndDoc
     End Select
    'notify admin the report is printing
    gstrMessage = "Your report is now available at the printer."
    gstrTitle = "Printing Completed ..."
    gintStyle = vbOKOnly
    MsgBox gstrMessage, gintStyle, gstrTitle
     'return to main admin form
     Unload Me
End Sub

Private Sub Form_Load()
    'declare the file number variable for this sub, and a boolean for testing
    Dim intFileHandle As Integer
    Dim blnFound As Boolean
    'change the form contents depending on the type of report to be printed
    Select Case gstrReportType
            Case "Date"
            'if it is a date report, then add the title, and the message, and fill the combo box with
            'date values
            frmReportSelector.Caption = "Date Transaction Report ..."
            lblMessage.Caption = "Please select the date for which the transactions should be printed: "
            intFileHandle = FreeFile
            Open gstrTransactionFile For Input As #intFileHandle
            Do While Not EOF(intFileHandle)
                'read a record
                 Input #intFileHandle, strDate, lngTransactionID, strUserName, strUserAccountNumber, strAccountType, _
                    strTransactionCode, curTransactionAmount, curUserChequingAccountBalance, curUserSavingsAccountBalance, _
                    intTellerDailyBalance
                blnFound = False
                'test for each date already in the combobox to see if the date value has been found
                For intCounter = 0 To cboSelection.ListCount
                    If Format(strDate, "yyyy/dd/mm") = Format(cboSelection.List(intCounter), "yyyy/dd/mm") Then
                        'if yes exit
                        blnFound = True
                        Exit For
                    Else
                        'if not found continue
                        blnFound = False
                    End If
                Next intCounter
                'if the current record date was not found add to the combo box list
                If blnFound = False Then
                    cboSelection.AddItem Format(strDate, "yyyy/mm/dd")
                End If
            Loop
            Close
        Case "Account"
            'if it is account report add form title and message, and fill combo box with account values
            frmReportSelector.Caption = "Account Transaction Report ..."
            lblMessage.Caption = "Please select the Account # for which the transactions should be printed: "
            intFileHandle = FreeFile
            Open gstrTransactionFile For Input As #intFileHandle
            Do While Not EOF(intFileHandle)
                'read a record
                 Input #intFileHandle, strDate, lngTransactionID, strUserName, strUserAccountNumber, strAccountType, _
                    strTransactionCode, curTransactionAmount, curUserChequingAccountBalance, curUserSavingsAccountBalance, _
                    intTellerDailyBalance
                'don't allow admin transactions to be printed
                'if more than one admin account change to variable
                If strUserAccountNumber <> "00000" Then
                    'test the account numberagainst each account number already in the combobox list
                    'to see if it has already been found
                    For intCounter = 0 To cboSelection.ListCount
                        If strUserAccountNumber = cboSelection.List(intCounter) Then
                            'if yes then exit
                            blnFound = True
                            Exit For
                        Else
                            'if no continue
                            blnFound = False
                        End If
                    Next intCounter
                    'if the current account number was not found add it to the combo box list
                    If blnFound = False Then
                        cboSelection.AddItem strUserAccountNumber
                    End If
                End If
            Loop
            Close
    End Select
End Sub

