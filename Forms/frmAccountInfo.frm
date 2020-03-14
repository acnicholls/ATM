VERSION 5.00
Begin VB.Form frmAccountInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Information ..."
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7680
   Icon            =   "frmAccountInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraChangePIN 
      Caption         =   "PIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   6855
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   5400
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdChangePIN 
         Caption         =   "OK"
         Height          =   495
         Left            =   5400
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtCONFIRMPIN 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3840
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtNEWPIN 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3840
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtOLDPIN 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Confirm New PIN: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1920
         TabIndex        =   15
         Top             =   1200
         Width           =   1860
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "New PIN: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   14
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Current PIN: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   1305
      End
   End
   Begin VB.Frame fraDisplay 
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   6015
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Savings Account Balance: "
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
         Height          =   240
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chequing Account Balance: "
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
         Height          =   240
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   2910
      End
      Begin VB.Label lblSavingsBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblChequingBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdPIN 
      Caption         =   "Change PIN"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmAccountInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare form wide variables for string testing
    Dim strOLDPIN As String * 4
    Dim strNEWPIN As String * 4
    Dim strConfirm As String * 4


Private Sub cmdCancel_Click()
'if the user cancels the attempt to change PIN then hide that frame and show the account balances
    fraChangePIN.Visible = False
    cmdReturn.Visible = True
    cmdPIN.Visible = True
    
End Sub

Private Sub cmdPIN_Click()
'when the user commands to change his/her PIN the interface changes, the frame containing the
'PIN textboxes becomes available, and the user is prompted to enter their current PIN
'the command from the interface are made invisible so the user cannot cancel from another location

    frmAccountInfo.Caption = "Change PIN ..."
    fraChangePIN.Visible = True
    'the NEW PIN and CONFIRM text boxes are disabled until a valid current PIN is entered
    Label7.Visible = False
    Label8.Visible = False
    txtNEWPIN.Visible = False
    txtCONFIRMPIN.Visible = False
    'the command buttons are also hidden or disabled until valid data is entered
    cmdChangePIN.Enabled = False
    cmdReturn.Visible = False
    cmdPIN.Visible = False
    txtOLDPIN.SetFocus
   
End Sub

Private Sub cmdReturn_Click()
'return to main form and enable timer
    frmMain.Timer1.Enabled = True
    Unload Me
    
End Sub

Private Sub cmdChangePIN_Click()
'declare values for file handling and string testing
    Dim intFileHandleInput As Integer
    Dim intFileHandleOutput As Integer
    Dim strUserName As String
    Dim strUserPIN As String
    Dim strUserAccountNumber As String
    'if all is matching then perform final test and change PIN
    If UCase(strConfirm) = UCase(strNEWPIN) And UCase(strOLDPIN) = UCase(gstrUserPIN) Then
    'open the PINs.dat file to read from
        intFileHandleInput = FreeFile
        Open App.Path & "\Data\PINs.dat" For Input As #intFileHandleInput
        'open a temporary file to write to
        intFileHandleOutput = FreeFile
        Open App.Path & "\Data\TempPINs.dat" For Output As #intFileHandleOutput
        Do While Not EOF(intFileHandleInput)
        'read a record
            Input #intFileHandleInput, strUserName, strUserPIN, strUserAccountNumber
            'test user name
            If strUserName <> gstrUserName Then
            'if not current user then write to temp file
                Write #intFileHandleOutput, strUserName, strUserPIN, strUserAccountNumber
            End If
        Loop
        'after all but current user are written to file add current user /w new PIN
        Write #intFileHandleOutput, gstrUserName, strNEWPIN, gstrUserAccountNumber
        Close
        'set global varibale to new PIN
        gstrUserPIN = strNEWPIN
        'erase old file and replace with temp file /w new values
        Kill App.Path & "\Data\PINs.dat"
        Name App.Path & "\Data\TempPINs.dat" As App.Path & "\Data\PINs.dat"
        'confirm the change to user
        gstrMessage = "Valid PIN, PIN changed successfully."
        gstrTitle = "PIN changed ..."
        gintStyle = vbOKOnly + vbInformation
        MsgBox gstrMessage, gintStyle, gstrTitle
        'hide the PIN change frame and allow user to check balances
        fraChangePIN.Visible = False
        cmdReturn.Visible = True
        Exit Sub
    End If
    'redundent error message
    gstrMessage = "Your PIN could not be changed as some data you entered was invalid"
    gstrTitle = "Invalid data ..."
    gintStyle = vbOKOnly + vbCritical
    MsgBox gstrMessage, gintStyle, gstrTitle
    ClearForm
End Sub

Private Sub Form_Load()
    'hide frame with pin change
    fraChangePIN.Visible = False
    'display current users balances
    lblChequingBalance.Caption = Format(gcurUserChequingAccountBalance, "currency")
    lblSavingsBalance.Caption = Format(gcurUserSavingsAccountBalance, "currency")
    
End Sub

Private Sub ClearForm()
    'clear the form after an error
    txtOLDPIN.Text = ""
    Label7.Visible = False
    txtNEWPIN.Visible = False
    txtNEWPIN.Text = ""
    Label8.Visible = False
    txtCONFIRMPIN.Visible = False
    txtCONFIRMPIN.Text = ""
    'set variable to zero values to ensure security
    strOLDPIN = ""
    strNEWPIN = ""
    strConfirm = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'activate time on return to main form
    frmMain.Timer1.Enabled = True
End Sub

Private Sub txtCONFIRMPIN_KeyPress(KeyAscii As Integer)
'prevent the user from entering invalid characters in the PIN
    Select Case KeyAscii
        Case 8
        'accept backspace
        Case 48 To 57
        'accept numbers
        Case 65 To 90
        'accept capital letters
        Case 97 To 122
        'accept small letters
        Case Else
        'everything else is ignored with an error msg
            gstrMessage = "Your PIN must be 4 characters long, and" & vbNewLine
            gstrMessage = gstrMessage & "Only number and letters may be used in your PIN."
            gintStyle = vbOKOnly + vbExclamation
            gstrTitle = "Invalid data ..."
            MsgBox gstrMessage, gintStyle, gstrTitle
            KeyAscii = 0
    End Select

End Sub

Private Sub txtCONFIRMPIN_KeyUp(KeyCode As Integer, Shift As Integer)
'when each character is entered test the length of the string..when it reaches 4 it's valid, test and move on
    If Len(txtCONFIRMPIN.Text) = 4 Then
    'send textbox to variable for testing
        strConfirm = txtCONFIRMPIN.Text
        'test new pin and confirm for equality
        If UCase(strConfirm) <> UCase(strNEWPIN) Then
            gstrMessage = "Incorrect  new PIN, Try again"
            gintStyle = vbOKOnly + vbCritical
            gstrTitle = "Invalid data ..."
            MsgBox gstrMessage, gintStyle, gstrTitle
            ClearForm
            Exit Sub
        End If
        ' if the new PIN and confirming PIN match then allow user to accept change
        'disable confirm text box so it cannot be changed
        cmdChangePIN.Enabled = True
        txtCONFIRMPIN.Enabled = False
        cmdChangePIN.SetFocus
    End If

End Sub

Private Sub txtNEWPIN_KeyPress(KeyAscii As Integer)
'prevent user from entering invalid characters in the PIN
    Select Case KeyAscii
        Case 8
        'accept backspace
        Case 48 To 57
        'accept numbers
        Case 65 To 90
        'accept capital letters
        Case 97 To 122
        'accept small letters
        Case Else
        'everything else is ignored with an error msg
            gstrMessage = "Your PIN must be 4 characters long, and" & vbNewLine
            gstrMessage = gstrMessage & "Only number and letters may be used in your PIN."
            gintStyle = vbOKOnly + vbExclamation
            gstrTitle = "Invalid data ..."
            MsgBox gstrMessage, gintStyle, gstrTitle
            KeyAscii = 0
    End Select

End Sub

Private Sub txtNEWPIN_KeyUp(KeyCode As Integer, Shift As Integer)
'when PIN reaches 4 characters long test and allow to continue if valid
    If Len(txtNEWPIN.Text) = 4 Then
    'send textbox variable to variable for testing
        strNEWPIN = txtNEWPIN.Text
        'test old PIN and new PIN for equality
        If UCase(strOLDPIN) = UCase(strNEWPIN) Then
            gstrMessage = "Old PIN and new PIN same, PIN not changed"
            gintStyle = vbOKOnly + vbCritical
            gstrTitle = "Invalid data ..."
            MsgBox gstrMessage, gintStyle, gstrTitle
            ClearForm
            Exit Sub
        End If
        'if the pins are different allow the user to confirm new PIN
        'disable new pin text box so it cannot be changed
        Label8.Visible = True
        txtCONFIRMPIN.Visible = True
        txtNEWPIN.Enabled = False
        txtCONFIRMPIN.SetFocus
    End If

End Sub

Private Sub txtOLDPIN_KeyPress(KeyAscii As Integer)
'prevent the user from entering invalid characters in the PIN
    Select Case KeyAscii
        Case 8
        'accept backspace
        Case 48 To 57
        'accept numbers
        Case 65 To 90
        'accept capital letters
        Case 97 To 122
        'accept small letters
        Case Else
        'everything else is ignored with an error msg
            gstrMessage = "Please enter your current PIN" & vbNewLine
            gstrMessage = gstrMessage & "Only number and letters may be used in your PIN."
            gintStyle = vbOKOnly + vbExclamation
            gstrTitle = "Invalid data ..."
            MsgBox gstrMessage, gintStyle, gstrTitle
            KeyAscii = 0
    End Select
   
End Sub

Private Sub CheckPIN()
End Sub

Private Sub txtOLDPIN_KeyUp(KeyCode As Integer, Shift As Integer)
'when PIN is 4 characters long send to test
     If Len(txtOLDPIN.Text) = 4 Then
    'send textbox value to variable for testing
        strOLDPIN = txtOLDPIN.Text
        'test oldpin for validity
        If UCase(strOLDPIN) <> UCase(gstrUserPIN) Then
            gstrMessage = "Incorrect PIN, Please enter correct previous PIN!"
            gintStyle = vbOKOnly + vbCritical
            gstrTitle = "Invalid data ..."
            MsgBox gstrMessage, gintStyle, gstrTitle
            ClearForm
            Exit Sub
        Else
        'if old pin is correct then allow user to enter new pin
        'disable old pin textbox so it cannot be changed
            Label7.Visible = True
            txtNEWPIN.Visible = True
            txtOLDPIN.Enabled = False
            txtNEWPIN.SetFocus
        End If
    End If

End Sub
