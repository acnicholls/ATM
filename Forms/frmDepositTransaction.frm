VERSION 5.00
Begin VB.Form frmDepositTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deposit Transaction ..."
   ClientHeight    =   4290
   ClientLeft      =   9735
   ClientTop       =   1530
   ClientWidth     =   7380
   Icon            =   "frmDepositTransaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame fraDeposit 
      Caption         =   "Deposit:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   3615
      Begin VB.OptionButton optType 
         Caption         =   " Savings account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optType 
         Caption         =   " Chequing account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   0
         Top             =   480
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame fraAmount 
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   3615
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
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
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   $"frmDepositTransaction.frx":0442
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
      Height          =   2295
      Left            =   4440
      TabIndex        =   7
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "frmDepositTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'cancel transaction and return to main form
    Unload Me

End Sub

Private Sub cmdOK_Click()
    'delcare transaction variables
    Dim curTransactionAmount As Currency
    'send transaction amount entered by user to variable for ease of use
    curTransactionAmount = Val(txtAmount.Text)
    'test for zero amount
    If curTransactionAmount <= 0 Then
        gstrMessage = "You must enter an amount larger than zero!"
        gstrTitle = "Invalid data ..."
        gintStyle = vbOKOnly + vbExclamation
        MsgBox gstrMessage, gintStyle, gstrTitle
        Exit Sub
    End If
    'depending on account type modify account accordingly and write to transaction file
    If optType(1).Value = True Then
        gcurUserChequingAccountBalance = gcurUserChequingAccountBalance + curTransactionAmount
        WriteTransaction "C", "D", curTransactionAmount
        gstrMessage = "New Chequing account balance:" & Format(gcurUserChequingAccountBalance, "currency")
    Else
        gcurUserSavingsAccountBalance = gcurUserSavingsAccountBalance + curTransactionAmount
        WriteTransaction "S", "D", curTransactionAmount
        gstrMessage = "New Savings account balance:" & Format(gcurUserSavingsAccountBalance, "currency")
    End If
    'remind user to deposit the envelope containing funds
    gstrMessage = gstrMessage & vbNewLine & "Don't forget to deposit your envelope!"
    gstrTitle = "Transaction completed ..."
    gintStyle = vbOKOnly + vbInformation
    MsgBox gstrMessage, gintStyle, gstrTitle
    'return to main form
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
'activate time on return to main form
    frmMain.Timer1.Enabled = True
    
End Sub


Private Sub optType_Click(Index As Integer)
'when user selects account ype send focus to text box
    txtAmount.SetFocus
    
End Sub

Private Sub txtAmount_GotFocus()
'select all text when focus arrives
    txtAmount.SelStart = 0
    txtAmount.SelLength = Len(txtAmount.Text)
    
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    ' set the global variable for message style
    gintStyle = vbOKOnly + vbExclamation
    Select Case KeyAscii
        Case 8
        'accept backspace
        Case 46
        'disallow decimal seperator
            MsgBox "Please do not deposit change.", gintStyle, "Invalid data ..."
            KeyAscii = 0
        Case 48 To 57
        'accept numbers 0 to 9
        Case Else
        'all else gets ignored
            KeyAscii = 0
    End Select
End Sub
