VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8985
   ClientLeft      =   6195
   ClientTop       =   645
   ClientWidth     =   11985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   45000
      Left            =   600
      Top             =   480
   End
   Begin VB.Frame fraTransactions 
      BackColor       =   &H8000000C&
      Height          =   5295
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CommandButton cmdTransaction 
         Height          =   375
         Index           =   5
         Left            =   2160
         Picture         =   "frmMain.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3840
         Width           =   495
      End
      Begin VB.CommandButton cmdTransaction 
         Height          =   375
         Index           =   6
         Left            =   2160
         Picture         =   "frmMain.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4440
         Width           =   495
      End
      Begin VB.CommandButton cmdTransaction 
         Height          =   375
         Index           =   4
         Left            =   2160
         Picture         =   "frmMain.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton cmdTransaction 
         Height          =   375
         Index           =   3
         Left            =   2160
         Picture         =   "frmMain.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2640
         Width           =   495
      End
      Begin VB.CommandButton cmdTransaction 
         Height          =   375
         Index           =   2
         Left            =   2160
         Picture         =   "frmMain.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdTransaction 
         Height          =   375
         Index           =   1
         Left            =   2160
         Picture         =   "frmMain.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblTransaction 
         BackColor       =   &H00C0FFFF&
         Caption         =   "  Account Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   5
         Left            =   2640
         TabIndex        =   17
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label lblTransaction 
         BackColor       =   &H00C0FFFF&
         Caption         =   "  Quit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   6
         Left            =   2640
         TabIndex        =   13
         Top             =   4440
         Width           =   2415
      End
      Begin VB.Label lblTransaction 
         BackColor       =   &H00C0FFFF&
         Caption         =   "  Bill Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   11
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label lblTransaction 
         BackColor       =   &H00C0FFFF&
         Caption         =   "  Transfer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   9
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label lblTransaction 
         BackColor       =   &H00C0FFFF&
         Caption         =   "  Withdrawal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   7
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblTransaction 
         BackColor       =   &H00C0FFFF&
         Caption         =   "  Deposit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   5
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   615
         Left            =   1200
         TabIndex        =   3
         Top             =   480
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   7440
      Width           =   6015
   End
   Begin VB.Label lblATS 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "A U T O M A T I C    T E L L E R   S I M U L A T O R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   1440
      TabIndex        =   14
      Top             =   720
      Width           =   9255
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2175
      Left            =   3720
      TabIndex        =   0
      Top             =   3000
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogin_Click()
    'load frmPIN and allow user to enter login info
    frmPIN.Show vbModal
    Set frmPIN = Nothing

End Sub

Private Sub cmdTransaction_Click(Index As Integer)
'this sub brings up the appropriate screen for the users desired transaction
'as well as disabling the timer so the user can complete their transaction
    Timer1.Enabled = False
    Select Case lblTransaction(Index).Caption
        Case "  Deposit"
            frmDepositTransaction.Show vbModal
        Case "  Withdrawal"
            frmWithdrawalTransaction.Show vbModal
        Case "  Transfer"
            frmTransferTransaction.Show vbModal
        Case "  Bill Payment"
            frmBillPaymentTransaction.Show vbModal
        Case "  Account Info"
            frmAccountInfo.Show vbModal
        Case "  Quit"
            SaveAccountsBalance
            fraTransactions.Visible = False
            frmMain.Refresh
    End Select
    
End Sub

Private Sub Form_Load()
    'set timer value to false to prevent timer activation without login
    Timer1.Enabled = False
    '  set filenames for all data files
    gstrDataFileLocation = GetSystemDrive & "\ProgramData\ATS"
    gstrAccountFile = gstrDataFileLocation & "\Accounts.dat"
    gstrTempAccountFile = gstrDataFileLocation & "\TempAccounts.dat"
    gstrTransactionFile = gstrDataFileLocation & "\Transactions.dat"
    gstrTransactionIdFile = gstrDataFileLocation & "\TransactionIdGenerator.dat"
    gstrDailyBalanceFile = gstrDataFileLocation & "\DailyBalances.dat"
    gstrTempDailyBalanceFile = gstrDataFileLocation & "\TempDailyBalances.dat"
    gstrDatabaseFile = gstrDataFileLocation & "\ATM.mdb"
    
    If Not FolderExists("C:\ProgramData\ATS") Then
        MkDir gstrDataFileLocation
        FileCopy App.Path & "\Data\Accounts.dat", gstrAccountFile
        FileCopy App.Path & "\Data\Transactions.dat", gstrTransactionFile
        FileCopy App.Path & "\Data\TransactionIDGenerator.dat", gstrTransactionIdFile
        FileCopy App.Path & "\Data\DailyBalances.dat", gstrDailyBalanceFile
        FileCopy App.Path & "\Data\ATM.mdb", gstrDatabaseFile
    End If
    

End Sub

Private Sub Form_Paint()
   'call to the function
    CheckTellerBalance
    'special label captions with specific format or including vbNewLine
    'if teller balance zero report to users
     If gintTellerDailyBalance <= 0 Then
        lblMessage.Caption = "This Automatic Teller Simulator is out of available funds."
        lblMessage.Caption = lblMessage.Caption & "Please use another machine for withdrawals." & vbNewLine
        cmdTransaction(2).Enabled = False
    Else
        lblMessage.Caption = vbNewLine & "To use this Automatic Teller Simulator, you must first login." & vbNewLine
        cmdTransaction(2).Enabled = True
    End If
    lblMessage.Caption = lblMessage.Caption & " Please click on the Login button."
    lblTitle.Caption = "S E L E C T   A" & vbNewLine & "T R A N S A C T I O N"
    lblDate.Caption = Format(Now, "dddd, mmmm d yyyy")

End Sub

Private Sub Timer1_Timer()
'when the timer times out, return to login screen and reset timer
    fraTransactions.Visible = False
    Timer1.Enabled = False
    frmMain.Refresh
End Sub


Private Function FolderExists(sFullPath As String) As Boolean
    Dim myFSO As Object
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = myFSO.FolderExists(sFullPath)
End Function
