VERSION 5.00
Begin VB.Form frmWithdrawalTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Withdrawal Transaction ..."
   ClientHeight    =   4290
   ClientLeft      =   -2430
   ClientTop       =   3300
   ClientWidth     =   7380
   Icon            =   "frmWithdrawalTransaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   6
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
         MaxLength       =   4
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame fraWithdrawal 
      Caption         =   "Withdrawal:"
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
      TabIndex        =   5
      Top             =   360
      Width           =   3615
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   $"frmWithdrawalTransaction.frx":0442
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
      Height          =   2055
      Left            =   4440
      TabIndex        =   7
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "frmWithdrawalTransaction"
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
    'declare variable for the amount of transaction and variable to test against
    Dim curTransactionAmount As Currency
    Dim curTransactionTest As Currency
    'move transaction amount to variable for ease of use
    curTransactionAmount = Val(txtAmount.Text)
    'test for zero amount
    If curTransactionAmount <= 0 Then
        gstrMessage = "You must enter an amount larger than zero!"
        gstrTitle = "Invalid data ..."
        gintStyle = vbOKOnly + vbExclamation
        MsgBox gstrMessage, gintStyle, gstrTitle
        Exit Sub
    End If
    'check to see if the teller has enough available funds
    If curTransactionAmount > gintTellerDailyBalance Then
        gstrMessage = "The amount of your withdrawal exceeds the amount avalaible in this teller device." & vbNewLine
        gstrMessage = gstrMessage & "unfortunately, for the moment, your are limited to a maximum of : "
        gstrMessage = gstrMessage & Format(gintTellerDailyBalance, "currency")
        gstrTitle = "Invalid data ..."
        gintStyle = vbOKOnly + vbExclamation
        MsgBox gstrMessage, gintStyle, gstrTitle
        txtAmount.Text = gintTellerDailyBalance
        txtAmount.SetFocus
        Exit Sub
    End If
    'check to see if amount exceeds the maximum limit
    If curTransactionAmount > 1000 Then
        gstrMessage = "The amount of your withdrawal exceeds the allowed per transaction maximum" & vbNewLine
        gstrMessage = gstrMessage & "of $1000,  please select a new amount within the allowable limit."
        gstrTitle = "Invalid data ..."
        gintStyle = vbOKOnly + vbExclamation
        MsgBox gstrMessage, gintStyle, gstrTitle
        txtAmount.SetFocus
        Exit Sub
    End If
    'check to see if amount is multiple of 10
    curTransactionTest = curTransactionAmount / 10
    If InStr(1, curTransactionTest, ".", 0) Then
        gstrMessage = "This machine only dispenses $10 bills, please enter an amount that is a multiple of ten"
        gstrTitle = "Invalid data ..."
        gintStyle = vbOKOnly + vbExclamation
        MsgBox gstrMessage, gintStyle, gstrTitle
        txtAmount.SetFocus
        Exit Sub
    End If
    'send user a message stating no more withdrawals if machine will empty with this transaction
    If curTransactionAmount = gintTellerDailyBalance Then
        gstrMessage = "Your withdrawal reduces the available funds to zero." & vbNewLine
        gstrMessage = gstrMessage & "This transaction WILL proceed" & vbNewLine
        gstrMessage = gstrMessage & "Please use another machine for more withdrawals"
        gstrTitle = "Warning ..."
        gintStyle = vbOKOnly + vbCritical
        MsgBox gstrMessage, gintStyle, gstrTitle
    End If
    'after testing to see if user has avilable funds
    'subtract transaction amount from users account and tellers balance
    'and write to transaction file
    If optType(1).Value = True Then
        If curTransactionAmount < gcurUserChequingAccountBalance Then
            gcurUserChequingAccountBalance = gcurUserChequingAccountBalance - curTransactionAmount
            'subtract transaction amount from tellers available funds
            gintTellerDailyBalance = gintTellerDailyBalance - curTransactionAmount
            WriteTransaction "C", "W", curTransactionAmount
            gstrMessage = "New Chequing account balance: " & Format(gcurUserChequingAccountBalance, "currency")
        Else
            gstrMessage = "The amount you entered exceeds available funds in the specified account." & vbNewLine
            gstrMessage = gstrMessage & "Available funds: " & Format(gcurUserChequingAccountBalance, "currency")
            gstrTitle = "Invalid data ..."
            gintStyle = vbOKOnly + vbCritical
            MsgBox gstrMessage, gintStyle, gstrTitle
            Exit Sub
        End If
    Else
        If curTransactionAmount < gcurUserSavingsAccountBalance Then
            gcurUserSavingsAccountBalance = gcurUserSavingsAccountBalance - curTransactionAmount
            'subtract transaction amount from tellers available funds
            gintTellerDailyBalance = gintTellerDailyBalance - curTransactionAmount
            WriteTransaction "S", "W", curTransactionAmount
            gstrMessage = "New Savings account balance: " & Format(gcurUserSavingsAccountBalance, "currency")
        Else
            gstrMessage = "The amount you entered exceeds available funds in the specified account." & vbNewLine
            gstrMessage = gstrMessage & "Available funds: " & Format(gcurUserSavingsAccountBalance, "currency")
            gstrTitle = "Invalid data ..."
            gintStyle = vbOKOnly + vbCritical
            MsgBox gstrMessage, gintStyle, gstrTitle
            Exit Sub
        End If
    End If
    'remind user to pickup their money from the machine
    gstrMessage = gstrMessage & vbNewLine & "Don't forget to pick-up your money!"
    gstrTitle = "Transaction complete ..."
    gintStyle = vbOKOnly + vbInformation
    MsgBox gstrMessage, gintStyle, gstrTitle
    'return to main form
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
'activate timer on return to main form
    frmMain.Timer1.Enabled = True
    
End Sub

Private Sub optType_Click(Index As Integer)
'when user selects account type send focus to amount textbox
    txtAmount.SetFocus
    
End Sub

Private Sub txtAmount_GotFocus()
    'select text in amount box for user ease of information change
    txtAmount.SelStart = 0
    txtAmount.SelLength = Len(txtAmount.Text)

End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    gintStyle = vbOKOnly + vbExclamation
    Select Case KeyAscii
        Case 8
            'accept backspace
        Case 46
            'do not accept decimal place holder, as change cannot be dispensed
            MsgBox "Change cannot be dispensed, Please do not enter an amount with cents", gintStyle, "Invalid data ..."
            KeyAscii = 0
        Case 48 To 57
            'accepts numers 0 to 9
        Case Else
            'anything else gets ignored
            KeyAscii = 0
    End Select
    
End Sub
