VERSION 5.00
Begin VB.Form frmBillPaymentTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill Payment Transaction ..."
   ClientHeight    =   3780
   ClientLeft      =   -1710
   ClientTop       =   -1170
   ClientWidth     =   7380
   Icon            =   "frmBillPaymentTransaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
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
      TabIndex        =   3
      Top             =   2160
      Width           =   3615
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
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
         MaxLength       =   8
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   $"frmBillPaymentTransaction.frx":0442
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
      Height          =   1455
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "frmBillPaymentTransaction"
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
    'declare transaction varibales
    Dim curTransactionAmount As Currency
    'take user input and send to variable
    curTransactionAmount = Val(txtAmount.Text)
    'test against zero amount
    If curTransactionAmount <= 0 Then
        gstrMessage = "You must enter an amount larger than zero!"
        gstrTitle = "Invalid data ..."
        gintStyle = vbOKOnly + vbExclamation
        MsgBox gstrMessage, gintStyle, gstrTitle
        Exit Sub
    End If
    'test for amount between  zero and 1 dollars
    If curTransactionAmount < 1 Then
        gstrMessage = "Are you sure you want to pay a bill for " & Format(curTransactionAmount, "currency")
        gstrTitle = "Verification ..."
        gintStyle = vbYesNo + vbQuestion
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        If gintAnswer = vbNo Then
            Exit Sub
        End If
    End If
    'don't allow anymore than $10000
    If curTransactionAmount > 10000 Then
        gstrMessage = "The maximum allowable payment is $10,000"
        gstrTitle = "Invalid data ..."
        gintStyle = vbOKOnly + vbExclamation
        MsgBox gstrMessage, gintStyle, gstrTitle
        Exit Sub
    End If
    'make sure amount is in users account
    If curTransactionAmount > gcurUserChequingAccountBalance Then
        gstrMessage = "The amount you entered exceeds the amount available in your account" & vbNewLine
        gstrMessage = gstrMessage & "Amount available: " & Format(gcurUserChequingAccountBalance, "currency")
        gstrTitle = "Invalid data ..."
        gintStyle = vbOKOnly + vbCritical
        MsgBox gstrMessage, gintStyle, gstrTitle
        Exit Sub
    End If
    'subtract transaction amount from users chequing account balance with service charge
    gcurUserChequingAccountBalance = gcurUserChequingAccountBalance - (curTransactionAmount + 1.25)
    'send appropriate variable to function to write to file
    WriteTransaction "C", "P", curTransactionAmount
    'remind user to insert bill stub
    gstrMessage = "New Chequing account balance: " & Format(gcurUserChequingAccountBalance, "currency")
    gstrMessage = gstrMessage & vbNewLine & "Don't forget to include your payment envelope!"
    gstrTitle = "Transaction completed ..."
    gintStyle = vbOKOnly + vbExclamation
    MsgBox gstrMessage, gintStyle, gstrTitle
    'return to main form
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
'activate timer on return to amin form
    frmMain.Timer1.Enabled = True
    
End Sub

Private Sub txtAmount_GotFocus()
'select all text when focus arrives
    txtAmount.SelStart = 0
    txtAmount.SelLength = Len(txtAmount.Text)

End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
'declare search start integer
    Dim intStart As Integer
    'set global message style and title for this sub
    gintStyle = vbOKOnly + vbExclamation
    gstrTitle = "Invalid data ..."
    Select Case KeyAscii
            Case 8
                'accept backspace
            Case 46
                'accept one decimal separator only
                If InStr(txtAmount.Text, ".") > 0 Then
                    MsgBox "You already typed a decimal separator.", gintStyle, gstrTitle
                    KeyAscii = 0
                End If
            Case 48 To 57
                'accepts numers 0 to 9, but disallows more than two decimal places
                If InStr(txtAmount.Text, ".") Then
                    intStart = InStr(1, txtAmount.Text, ".")
                    If Len(Mid(txtAmount.Text, intStart)) > 2 Then
                        gstrMessage = "Please enter only two digits to the right of the decimal seperator."
                        MsgBox gstrMessage, gintStyle, gstrTitle
                        KeyAscii = 0
                    End If
                End If
            Case Else
            'all else is ignored
                KeyAscii = 0
        End Select
    
End Sub
