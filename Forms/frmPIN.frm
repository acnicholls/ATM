VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPIN 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Login ...."
   ClientHeight    =   3255
   ClientLeft      =   3135
   ClientTop       =   2745
   ClientWidth     =   6060
   ForeColor       =   &H00000000&
   Icon            =   "frmPIN.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3255
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adoPINS 
      Height          =   855
      Left            =   120
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data\ATM.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Data\ATM.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblPINS"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtPIN 
      ForeColor       =   &H00800080&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox picKeys 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   600
      Picture         =   "frmPIN.frx":27A2
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblUserName 
      Caption         =   "&Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblPIN 
      Caption         =   "&PIN:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblSecurityMessage 
      Caption         =   $"frmPIN.frx":4F44
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmPIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form gives the user or administrator the ability to log into the ATS, validates the PIN and sends the
'user to the proper form, be it the main, or administrator form
Dim mintCounter As Integer


Private Sub cmdCancel_Click()
'on cancel return to main form
    Unload Me
    Set frmPIN = Nothing

End Sub

Private Sub cmdOK_Click()
'this sub varifies lgoin data and directs the user an appropriate screen based on security level
    Dim blnLoginOkay As Boolean
'    Dim strUserType As String
'    Dim intFileHandle As Integer
    Dim strName As String
    Dim strPIN As String
    Dim strAccountNumber As String
'****************************************REMOVED CODE
                                                
                                                '    'test user name field for info
                                                '    If txtUserName.Text = "" Then
                                                '        MsgBox " You must enter your name to proceed", vbOKOnly, "Missing field ..."
                                                '        txtUserName.SetFocus
                                                '        Exit Sub
                                                '    End If
                                                '    'test PIN field for info
                                                '    If txtPIN.Text = "" Then
                                                '        MsgBox "You must enter your PIN to proceed", vbOKOnly, "Missing field ..."
                                                '        txtPIN.SetFocus
                                                '        Exit Sub
                                                '    End If
    'if both fields are filled with info, count this as a login attempt
    'Add one to counter and ready variables for testing
    mintCounter = mintCounter - 1
    blnLoginOkay = True
'    strUserType = ""
    '''''REady the recordset
    adoPINS.Refresh
    'goto first record
    adoPINS.Recordset.MoveFirst
    'find the record that matches the pin
    adoPINS.Recordset.Find "fldPIN = '" & txtPIN.Text & "'"
    'if not found then login not okay
    If adoPINS.Recordset.EOF Then
        blnLoginOkay = False
    Else
    'if found test against username
        If adoPINS.Recordset.Fields("fldLastName") <> txtUserName.Text Then
        'no match login NOT okay
            blnLoginOkay = False
        Else
            ' load the user's account details
            gstrUserAccountNumber = adoPINS.Recordset.Fields("fldAccount")
            gstrUserPIN = txtPIN.Text
            gstrUserName = txtUserName.Text
        End If
    End If
    
                                                                       
'''''''''''''''''''''''''''''''''Code to check how many attempts have been made no more that 3 allowed
' if the user name and pin have not been validated then notify the user and begin again
    If blnLoginOkay = False Then
            If mintCounter = 0 Then
                'if the user has attempted 3 logins return to main screen
                gintStyle = vbOKOnly + vbCritical
                gstrTitle = "Automatic Teller Simulator ... Account Locked."
                gstrMessage = "You are restricted to three (3) attempts to login to the system." & vbNewLine
                gstrMessage = gstrMessage & " Please try again later or verify your PIN with the system administrator!"
                MsgBox gstrMessage, gintStyle, gstrTitle
                Unload Me
            Else
                'if the user has more attempts allowed clear the form and allow another attempt
                gintStyle = vbOKOnly + vbExclamation
                gstrTitle = "Automatic Teller Simulator ... Login System."
                gstrMessage = "Login information is not valid!" & vbNewLine & vbNewLine
                gstrMessage = gstrMessage & "You are still allowed " & mintCounter
                gstrMessage = gstrMessage & " attempt(s) to login to the system." & vbNewLine
                gstrMessage = gstrMessage & " Please try again later or verify your PIN with the system administrator!"
                If DEMO_MODE Then
                    gstrMessage = gstrMessage & vbNewLine & vbNewLine
                    gstrMessage = gstrMessage & "For this demo, if you want to login as the administrator use:" & vbNewLine
                    gstrMessage = gstrMessage & "Name: Dallas" & vbTab & "PIN: D001" & vbNewLine & vbNewLine
                    gstrMessage = gstrMessage & "If you want to login as a regular client use:" & vbNewLine
                    gstrMessage = gstrMessage & "Name: Clapton  " & vbTab & "PIN: C002"
                End If
                MsgBox gstrMessage, gintStyle, gstrTitle
                'focus on user name
                txtUserName.SetFocus
                Exit Sub
            End If
        End If

'**********
'********if it's not a user, and there's enuf attempts left, is it an admin???
'*******

'   code to validate ADMIN user
    If adoPINS.Recordset.Fields("fldPIN") = ADMINISTRATOR_PIN Then
            gintStyle = vbOKOnly + vbExclamation
            gstrTitle = "Successful Login to Automatic Teller Simulator ..."
            gstrMessage = "Welcome  " & strName & vbNewLine & vbNewLine
            gstrMessage = gstrMessage & " Access to the system has been granted with administrator rights."
            MsgBox gstrMessage, gintStyle, gstrTitle
            Unload Me
            Set frmPIN = Nothing
            'show admin screen only allowing admin form to be accessed
            frmAdministrator.Show vbModal
            Set frmAdministrator = Nothing
    Else
            gintStyle = vbOKOnly + vbExclamation
            gstrTitle = "Successful Login to Automatic Teller Simulator ..."
            gstrMessage = "Welcome  " & strName & vbNewLine & vbNewLine
            gstrMessage = gstrMessage & " You have been granted access to the Simulator."
            'display a message stating the user has been accepted
            MsgBox gstrMessage, gintStyle, gstrTitle
            'show the transactions frame
            frmMain.fraTransactions.Visible = True
            'start the timer
            frmMain.Timer1.Enabled = True
            'refresh the main screen to show the transaction frame
            frmMain.Refresh
            'call the sub to read the users account balances
            ReadAccountsBalance
            'unload the login form and disallow it to show again
            Unload Me
    End If
    Set frmPIN = Nothing

End Sub


Private Sub Form_Load()
   mintCounter = 3

   
End Sub

Private Sub txtPIN_GotFocus()
'select all text in box when focus arrives
    txtPIN.SelStart = 0
    txtPIN.SelLength = Len(txtPIN.Text)
End Sub

Private Sub txtPIN_KeyPress(KeyAscii As Integer)
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
            gstrMessage = "Only Numbers and letters may be used in your PIN!"
            gintStyle = vbOKOnly + vbExclamation
            gstrTitle = "Invalid data ..."
            MsgBox gstrMessage, gintStyle, gstrTitle
            KeyAscii = 0
    End Select

End Sub

Private Sub txtUserName_GotFocus()
'select all text in box when focus arrives
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        'accept backspace
        Case 32
        'do NOT accept space
            gstrMessage = "You must enter only your last name."
            gstrTitle = "Invalid Data ..."
            gintStyle = vbOKOnly + vbExclamation
            MsgBox gstrMessage, gintStyle, gstrTitle
            KeyAscii = 0
        Case 65 To 90
        'accept capital letters
        Case 97 To 122
        'accept small letters
        Case Else
            gstrMessage = "You must enter only your last name."
            gstrTitle = "Invalid Data ..."
            gintStyle = vbOKOnly + vbExclamation
            MsgBox gstrMessage, gintStyle, gstrTitle
            KeyAscii = 0
    End Select

End Sub
