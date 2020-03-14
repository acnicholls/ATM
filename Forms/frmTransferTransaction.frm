VERSION 5.00
Begin VB.Form frmTransferTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Transaction ..."
   ClientHeight    =   4290
   ClientLeft      =   -2835
   ClientTop       =   6000
   ClientWidth     =   7380
   Icon            =   "frmTransferTransaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTransfer 
      Caption         =   "Transfer:"
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
      TabIndex        =   7
      Top             =   360
      Width           =   3615
      Begin VB.OptionButton optType 
         Caption         =   " Savings to Chequing"
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
         Width           =   2535
      End
      Begin VB.OptionButton optType 
         Caption         =   " Chequing to Savings"
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
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
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
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   $"frmTransferTransaction.frx":0442
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
      TabIndex        =   6
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "frmTransferTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    'declare a transaction amount variable
    Dim curTransactionAmount As Currency
    Dim lngCurrentTransactionID As Long
    'send the textbox info to the variable for ease of use
    curTransactionAmount = Val(txtAmount.Text)
    'grab the transactionID to pass to user for confirmation
    lngCurrentTransactionID = CLng(GetNextTransactionID)
    'test againt zero amount
    If curTransactionAmount <= 0 Then
        gstrMessage = "You must enter an amount larger than zero!"
        gstrTitle = "Invalid data ..."
        gintStyle = vbOKOnly + vbExclamation
        MsgBox gstrMessage, gintStyle, gstrTitle
        Exit Sub
    End If
    'test to see if the amount exceeds the maximum limit
    If curTransactionAmount > 100000 Then
        gstrMessage = "The amount you entered exceeds the allowable limit of $100,000.00" & vbNewLine
        gstrMessage = gstrMessage & "Please enter an amount within acceptable limits."
        gstrTitle = "Invalid data ..."
        gintStyle = vbOKOnly + vbExclamation
        MsgBox gstrMessage, gintStyle, gstrTitle
        Exit Sub
    End If
    'test to see if user has funds available in specified account then
    'depending on the direction of transfer credit the proper account, and debit the other, then
    'write the transaction to file
    gstrMessage = Format(curTransactionAmount, "currency") & " sent from "
    If optType(1) = True Then
        If curTransactionAmount < gcurUserChequingAccountBalance Then
            gcurUserChequingAccountBalance = gcurUserChequingAccountBalance - curTransactionAmount
            gcurUserSavingsAccountBalance = gcurUserSavingsAccountBalance + curTransactionAmount
            WriteTransaction "C", "T", curTransactionAmount
            gstrMessage = gstrMessage & "Chequing account to Savings account." & vbNewLine & vbNewLine
            gstrMessage = gstrMessage & "New Chequing account balance:" & Format(gcurUserChequingAccountBalance, "currency")
            gstrMessage = gstrMessage & vbNewLine
            gstrMessage = gstrMessage & "New Savings account balance:" & Format(gcurUserSavingsAccountBalance, "currency") & vbNewLine
            gstrMessage = gstrMessage & vbNewLine
        Else
            gstrMessage = "The amount you entered exceeds your available funds in the specified account." & vbNewLine
            gstrMessage = gstrMessage & "Available funds: " & Format(gcurUserChequingAccountBalance, "currency")
            gstrTitle = "Invalid data ..."
            gintStyle = vbOKOnly + vbCritical
            MsgBox gstrMessage, gintStyle, gstrTitle
            Exit Sub
        End If
    Else
        If curTransactionAmount < gcurUserSavingsAccountBalance Then
            gcurUserSavingsAccountBalance = gcurUserSavingsAccountBalance - curTransactionAmount
            gcurUserChequingAccountBalance = gcurUserChequingAccountBalance + curTransactionAmount
            WriteTransaction "S", "T", curTransactionAmount
            gstrMessage = gstrMessage & "Savings account to Chequing account." & vbNewLine & vbNewLine
            gstrMessage = gstrMessage & "New Chequing account balance:" & Format(gcurUserChequingAccountBalance, "currency")
            gstrMessage = gstrMessage & vbNewLine
            gstrMessage = gstrMessage & "New Savings account balance:" & Format(gcurUserSavingsAccountBalance, "currency") & vbNewLine
            gstrMessage = gstrMessage & vbNewLine
        Else
            gstrMessage = "The amount you entered exceeds available funds in the specified account." & vbNewLine
            gstrMessage = gstrMessage & "Available funds: " & Format(gcurUserSavingsAccountBalance, "currency")
            gstrTitle = "Invalid data ..."
            gintStyle = vbOKOnly + vbCritical
            MsgBox gstrMessage, gintStyle, gstrTitle
            Exit Sub
        End If
    End If
    'add transaction number for user confirmation
    gstrMessage = gstrMessage & "Confirmation #" & lngCurrentTransactionID
    gstrTitle = "Transfer complete ..."
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
'select all text in textbox when focus arrives
    txtAmount.SelStart = 0
    txtAmount.SelLength = Len(txtAmount.Text)
    
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    'declare local values of global message variables
    Dim intStart As Integer
    gintStyle = vbOKOnly + vbExclamation
    gstrTitle = "Invalid data ..."
    Select Case KeyAscii
        Case 8
            'accept backspace
        Case 46
            'accept only one decimal separator
            If InStr(1, txtAmount.Text, ".", 0) Then
                gstrMessage = "You have already entered a decimal separator!"
                MsgBox gstrMessage, gintStyle, gstrTitle
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
            KeyAscii = 0
    End Select
    
End Sub
