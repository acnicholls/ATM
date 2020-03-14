Attribute VB_Name = "LIB_ATS"
Option Explicit

'various global constants
Public Const DEMO_MODE As Boolean = True

'global constants used to identify the system administrator
Public Const ADMINISTRATOR_NAME As String = "Korben Dallas"
Public Const ADMINISTRATOR_PIN As String = "D001"

'global variables used to display message to user
Public gstrMessage As String
Public gstrTitle As String
Public gintStyle As Integer
Public gintAnswer As Integer

'various global variables
Public gstrReportType As String
Public gstrTransactionReportDate As String
Public gstrTransactionReportAccountNumber As String
Public gintTellerDailyBalance As Integer

'global variables used to identify the current user
Public gstrUserName As String
Public gstrUserPIN As String * 4
Public gstrUserAccountNumber As String * 5
Public gcurUserChequingAccountBalance As Currency
Public gcurUserSavingsAccountBalance As Currency

Public Sub ReadAccountsBalance()
    'this sub is used to  read a users account balances from the file
    'it stores these values in global variables
    'declare sub variables
    Dim intFileHandleInput As Integer
    Dim strAccountNumber As String
    Dim curChequingAccountBalance As Currency
    Dim curSavingsAccountBalance As Currency
    'open the file and read a record, test record against current user account number
    intFileHandleInput = FreeFile
    Open App.Path & "\Data\Accounts.dat" For Input As #intFileHandleInput
    'while there is information to be read input each record and test the User account number
    'if it tests true then apply users account balances to global variables and exit
    Do While Not EOF(intFileHandleInput)
        Input #intFileHandleInput, strAccountNumber, curChequingAccountBalance, curSavingsAccountBalance
        If strAccountNumber = gstrUserAccountNumber Then
            gcurUserChequingAccountBalance = curChequingAccountBalance
            gcurUserSavingsAccountBalance = curSavingsAccountBalance
            Exit Do
        End If
    Loop
    
End Sub

Public Sub WriteTransaction(ByVal strAccountType, ByVal strTransactionCode, ByVal curTransactionAmount)
' this sub takes values passed from other subs and writes a record in the Transactions.dat file concerning the
'current transaction
    Dim intFileHandle As Integer
    Dim blnResult As Boolean
    Dim lngCurrentTransactionID As Long
    'acquire a transaction ID to identify this transaction to the computer
    lngCurrentTransactionID = CLng(GetNextTransactionID)
    intFileHandle = FreeFile
    'open the file and write current transaction record
    Open App.Path & "\Data\Transactions.dat" For Append As #intFileHandle
    Write #intFileHandle, Format(Date, "mm/dd/yyyy"); lngCurrentTransactionID; gstrUserName; gstrUserAccountNumber; _
                                      strAccountType; strTransactionCode; curTransactionAmount; gcurUserChequingAccountBalance; _
                                      gcurUserSavingsAccountBalance; gintTellerDailyBalance
    Close #intFileHandle
    'call the transactionID generator
    blnResult = GenerateNextTransactionID(lngCurrentTransactionID)
    'save all acounts to file
    SaveAccountsBalance
    
End Sub

Public Sub WriteAdministratorTransaction(ByVal strTransactionCode)
'this sub writes the administrator fills, and the daily auto fill to the transaction file
    Dim intFileHandle As Integer
    Dim blnResult As Boolean
    Dim lngCurrentTransactionID As Long
    'get newest transaction ID and write transaction to file depending on transaction code
    lngCurrentTransactionID = CLng(GetNextTransactionID)
    intFileHandle = FreeFile
    Open App.Path & "\Data\Transactions.dat" For Append As #intFileHandle
    Write #intFileHandle, Format(Date, "mm/dd/yyyy"); lngCurrentTransactionID; "ADMINISTRATOR"; "00000"; _
                                      ""; strTransactionCode; 5000, 0; 0; gintTellerDailyBalance
    Close #intFileHandle
    'write the new balance to dailybalances.dat
    SaveTellerBalance
    'call the transactionID generator
    blnResult = GenerateNextTransactionID(lngCurrentTransactionID)
    
End Sub

Public Sub CheckTellerBalance()
'this sub checks to see how much money is currently in the machine
    Dim blnFound As Boolean
    Dim dtmToday As Date
    Dim strDate As String
    Dim intBalance As Integer
    Dim intFileHandle As Integer
    
    blnFound = False
    dtmToday = Date
    
    'read Teller's balance
    intFileHandle = FreeFile
    Open App.Path & "\Data\DailyBalances.dat" For Input As #intFileHandle
    Do While Not EOF(intFileHandle)
        Input #intFileHandle, strDate, intBalance
        'if found use global variable to hold the amount found for today's balance
        If Format(strDate, "mm/dd/yyyy") = dtmToday Then
            blnFound = True
            gintTellerDailyBalance = intBalance
            Exit Do
        End If
    Loop
    Close
    intFileHandle = FreeFile
    'if ATS is accessed for the first time today apply automatic fill up of $5,000 and write to file
    If Not blnFound Then
        gintTellerDailyBalance = 5000
        strDate = Format(dtmToday, "mm/dd/yyyy")
        WriteAdministratorTransaction ("F")
        Open App.Path & "\Data\DailyBalances.dat" For Append As #intFileHandle
        Write #intFileHandle, strDate, gintTellerDailyBalance
        Close #intFileHandle
    End If

End Sub

Public Sub SaveTellerBalance()
'this sub saves the current amount of money in the teller machine to a file for later retrieval

    Dim intFileHandleInput As Integer
    Dim intFileHandleOutput As Integer
    Dim dtmDate As Date
    Dim strDate As String
    Dim intBalance As Integer
    'give the current date to a variable
    dtmDate = Now
    'open a file for reading
    intFileHandleInput = FreeFile
    Open App.Path & "\Data\DailyBalances.dat" For Input As #intFileHandleInput
    'open a file for writing to
    intFileHandleOutput = FreeFile
    Open App.Path & "\Data\TempDailyBalances.dat" For Output As #intFileHandleOutput
    'while there is still information to be read, test the date against today's date and write to file if not equal
    Do While Not EOF(intFileHandleInput)
        Input #intFileHandleInput, strDate, intBalance
        If strDate <> Format(dtmDate, "mm/dd/yyyy") Then
        'write all but today's record to temp file
            Write #intFileHandleOutput, strDate, intBalance
        End If
    Loop
    'write today's balance to the file
    Write #intFileHandleOutput, Format(dtmDate, "mm/dd/yyyy"); gintTellerDailyBalance
    Close
    'change the temp output file to the permanent DailyBalances.dat file
    Kill App.Path & "\Data\DailyBalances.dat"
    Name App.Path & "\Data\TempDailyBalances.dat" As App.Path & "\Data\DailyBalances.dat"

End Sub

Public Function GetNextTransactionID() As Long
'this function passes the next transaction ID from the file to the sub that called it
    Dim intFileHandle As Integer
    Dim strNextTransactionID As String
    'open the file get the ID and close
     intFileHandle = FreeFile
     Open App.Path & "\Data\TransactionIDGenerator.dat" For Input As #intFileHandle
     Input #intFileHandle, strNextTransactionID
     Close #intFileHandle
     'make sure it's a long value, no decimals allowed
     GetNextTransactionID = CLng(strNextTransactionID)

End Function

Public Function GenerateNextTransactionID(ByVal strCurrentTransactionID As String) As Boolean
'this function passes the number of the next transaction to file
    Dim intFileHandle As Integer
        
     intFileHandle = FreeFile
    'add one to the current transaction ID and save to file for retrieval
     strCurrentTransactionID = strCurrentTransactionID + 1
     Open App.Path & "\Data\TransactionIDGenerator.dat" For Output As #intFileHandle
     Write #intFileHandle, Format(Val(strCurrentTransactionID), "0000000000")
     Close #intFileHandle
    'set the boolean concerning the next transaction ID number generated
     GenerateNextTransactionID = True

End Function

Public Sub SaveAccountsBalance()
'this sub saves to file all account balances of the current user on exit
    Dim intFileHandleInput As Integer
    Dim intFileHandleOutput As Integer
    Dim strUserAccountNumber As String
    Dim curChequingAccountBalance As Currency
    Dim curSavingsAccountBalance As Currency
    Dim dtmToday As Date
    Dim strDate As String
    Dim intTellerBalance As Integer
    'declare today's date to variable
    dtmToday = Date
    'open Accounts.dat file to read from
    intFileHandleInput = FreeFile
    Open App.Path & "\Data\Accounts.dat" For Input As #intFileHandleInput
    'create temporary file to write to
    intFileHandleOutput = FreeFile
    Open App.Path & "\Data\TempAccounts.dat" For Output As #intFileHandleOutput
    Do While Not EOF(intFileHandleInput)
    'read a record
        Input #intFileHandleInput, strUserAccountNumber, curChequingAccountBalance, curSavingsAccountBalance
        'test for current user..all but current get written to file
        If strUserAccountNumber <> gstrUserAccountNumber Then
            'write all other users to file
            Write #intFileHandleOutput, strUserAccountNumber, curChequingAccountBalance, curSavingsAccountBalance
        End If
    Loop
    'update current users balances
    Write #intFileHandleOutput, gstrUserAccountNumber, gcurUserChequingAccountBalance, gcurUserSavingsAccountBalance
    Close
    'delete old file and replace with temp one
    Kill App.Path & "\Data\Accounts.dat"
    Name App.Path & "\Data\TempAccounts.dat" As App.Path & "\Data\Accounts.dat"
    'send call to save teller's balance
    SaveTellerBalance

End Sub

Public Sub PrintReportHeader()
    Dim intMyTab As Integer
    Dim strTitle As String
    Dim strTitle1 As String
    
    'set landscape tab
    intMyTab = 95
    
    'set printer options
    With Printer
        .PrintQuality = vbPRPQDraft
        .FontName = "courier new"
        .FontBold = True
        .Orientation = vbPRORLandscape
    End With
    'set report titles
    strTitle = "***** T R A N S A C T I O N   R E P O R T *****"
    strTitle1 = "All Transactions"
    Select Case gstrReportType
        Case "Date"
            strTitle1 = strTitle1 & " for " & Format(gstrTransactionReportDate, "mm/dd/yyyy")
        Case "Account"
            strTitle1 = strTitle1 & " for Account #" & gstrTransactionReportAccountNumber
        Case "All"
            strTitle1 = strTitle1 & " on file"
    End Select
    
    
    'Print Report Header
    Printer.FontSize = 40
    Printer.Print Space(2); "CONFIDENTIAL"
    Printer.FontSize = 14
    Printer.Print
    Printer.Print Space((intMyTab - Len("Automatic Teller Simulator")) / 2); "Automatic Teller Simulator"
    Printer.Print
    Printer.Print Space((intMyTab - Len(strTitle)) / 2); strTitle
    Printer.Print Space((intMyTab - Len(strTitle1)) / 2); strTitle1
    Printer.Print
    Printer.FontSize = 12
    Printer.Print Space(3) & "DATE:"; Tab(intMyTab + 6); "TIME:"
    Printer.Print Space(3) & Format(Date, "mm/dd/yyyy"); Tab(intMyTab + 3); Format(Time, "hh:mm:ss")
    Printer.Print
    Printer.Print
    Printer.FontSize = 10
    Printer.Print Space(3); "Date";
    Printer.Print Tab(16); "Number";
    Printer.Print Tab(30); "Name";
    Printer.Print Tab(53); "Acc #";
    Printer.Print Tab(62); "Type";
    Printer.Print Tab(68); "Code";
    Printer.Print Tab(78); "Amount";
    Printer.Print Tab(89); "Chequing $";
    Printer.Print Tab(102); "Savings $";
    Printer.Print Tab(115); "ATS Balance"
    Printer.Print Space(2); String(124, "_")
    Printer.Print
    
End Sub

Public Sub PrintDateReport()
'declare file handles, variables for file reading, and variable for data printing
    Dim intFileHandleTransaction As Integer
    Dim strDate As String
    Dim lngTransactionID As Long
    Dim strUserName As String
    Dim strUserAccountNumber
    Dim strAccountType As String
    Dim strTransactionCode As String
    Dim curTransactionAmount As Single
    Dim curUserChequingAccountBalance As Currency
    Dim curUserSavingsAccountBalance As Currency
    Dim intTellerBalance As Integer
    Dim strDataToPrint As String * 14
    Dim strShortDataToPrint As String * 6
    Dim strLongDataToPrint As String * 23
    'open the transactions file
    intFileHandleTransaction = FreeFile
    Open App.Path & "\Data\Transactions.dat" For Input As #intFileHandleTransaction
    Do While Not EOF(intFileHandleTransaction)
        Input #intFileHandleTransaction, strDate, lngTransactionID, strUserName, strUserAccountNumber, _
            strAccountType, strTransactionCode, curTransactionAmount, curUserChequingAccountBalance, _
            curUserSavingsAccountBalance, intTellerBalance
        'after readin 1 record test the date value for validity, if the value matches the desired report
        'enter the variable value into print variables to be justified and sent to printer
        If Format(strDate, "mm/dd/yyyy") = Format(gstrTransactionReportDate, "mm/dd/yyyy") Then
            Printer.Print Space(2); Format(strDate, "mm/dd/yyyy"); Space(2);
            LSet strDataToPrint = lngTransactionID
            Printer.Print strDataToPrint;
            LSet strLongDataToPrint = strUserName
            Printer.Print strLongDataToPrint;
            RSet strShortDataToPrint = strUserAccountNumber
            Printer.Print strShortDataToPrint;
            RSet strShortDataToPrint = strAccountType
            Printer.Print strShortDataToPrint;
            RSet strShortDataToPrint = strTransactionCode
            Printer.Print strShortDataToPrint;
            RSet strDataToPrint = Format(curTransactionAmount, "currency")
            Printer.Print strDataToPrint;
            RSet strDataToPrint = Format(curUserChequingAccountBalance, "currency")
            Printer.Print strDataToPrint;
            RSet strDataToPrint = Format(curUserSavingsAccountBalance, "currency")
            Printer.Print strDataToPrint;
            RSet strDataToPrint = Format(intTellerBalance, "currency")
            Printer.Print strDataToPrint
        End If
    Loop
    'after all transaction records have been tested, close file and add end of report
    Close
    Printer.Print
    Printer.FontSize = 16
    Printer.Print Space(6); "End of Report"
End Sub

Public Sub PrintAllTransactions()
'declare filehandle variable, input variables, and printer justification variables
    Dim intFileHandleTransaction As Integer
    Dim strDate As String
    Dim lngTransactionID As Long
    Dim strUserName As String
    Dim strUserAccountNumber
    Dim strAccountType As String
    Dim strTransactionCode As String
    Dim curTransactionAmount As Currency
    Dim curUserChequingAccountBalance As Currency
    Dim curUserSavingsAccountBalance As Currency
    Dim intTellerBalance As Integer
    Dim strDataToPrint As String * 14
    Dim strShortDataToPrint As String * 6
    Dim strLongDataToPrint As String * 23
    'open the transaction file
    intFileHandleTransaction = FreeFile
    Open App.Path & "\Data\Transactions.dat" For Input As #intFileHandleTransaction
    Do While Not EOF(intFileHandleTransaction)
    'input and print all records on file...no testing here folks!!
        Input #intFileHandleTransaction, strDate, lngTransactionID, strUserName, strUserAccountNumber, _
            strAccountType, strTransactionCode, curTransactionAmount, curUserChequingAccountBalance, _
            curUserSavingsAccountBalance, intTellerBalance
        Printer.Print Space(2); Format(strDate, "mm/dd/yyyy"); Space(2);
        LSet strDataToPrint = lngTransactionID
        Printer.Print strDataToPrint;
        LSet strLongDataToPrint = strUserName
        Printer.Print strLongDataToPrint;
        RSet strShortDataToPrint = strUserAccountNumber
        Printer.Print strShortDataToPrint;
        RSet strShortDataToPrint = strAccountType
        Printer.Print strShortDataToPrint;
        RSet strShortDataToPrint = strTransactionCode
        Printer.Print strShortDataToPrint;
        RSet strDataToPrint = Format(curTransactionAmount, "currency")
        Printer.Print strDataToPrint;
        RSet strDataToPrint = Format(curUserChequingAccountBalance, "currency")
        Printer.Print strDataToPrint;
        RSet strDataToPrint = Format(curUserSavingsAccountBalance, "currency")
        Printer.Print strDataToPrint;
        RSet strDataToPrint = Format(intTellerBalance, "currency")
        Printer.Print strDataToPrint
    Loop
    'close file and add end of report
    Close
    Printer.Print
    Printer.FontSize = 16
    Printer.Print Space(6); "End of Report"
        
End Sub

Public Sub PrintAccountReport()
'delcare file handlers, input variables, and print justification variables
    Dim intFileHandleTransaction As Integer
    Dim strDate As String
    Dim lngTransactionID As Long
    Dim strUserName As String
    Dim strUserAccountNumber
    Dim strAccountType As String
    Dim strTransactionCode As String
    Dim curTransactionAmount As Currency
    Dim curUserChequingAccountBalance As Currency
    Dim curUserSavingsAccountBalance As Currency
    Dim intTellerBalance As Integer
    Dim strDataToPrint As String * 14
    Dim strShortDataToPrint As String * 6
    Dim strLongDataToPrint As String * 23
    'open transaction file
    intFileHandleTransaction = FreeFile
    Open App.Path & "\Data\Transactions.dat" For Input As #intFileHandleTransaction
    Do While Not EOF(intFileHandleTransaction)
        Input #intFileHandleTransaction, strDate, lngTransactionID, strUserName, strUserAccountNumber, _
            strAccountType, strTransactionCode, curTransactionAmount, curUserChequingAccountBalance, _
            curUserSavingsAccountBalance, intTellerBalance
        'after readin 1 record test the account number variable for validity against the desired
        'account records to be printed
        If strUserAccountNumber = gstrTransactionReportAccountNumber Then
            Printer.Print Space(2); Format(strDate, "mm/dd/yyyy"); Space(2);
            LSet strDataToPrint = lngTransactionID
            Printer.Print strDataToPrint;
            LSet strLongDataToPrint = strUserName
            Printer.Print strLongDataToPrint;
            RSet strShortDataToPrint = strUserAccountNumber
            Printer.Print strShortDataToPrint;
            RSet strShortDataToPrint = strAccountType
            Printer.Print strShortDataToPrint;
            RSet strShortDataToPrint = strTransactionCode
            Printer.Print strShortDataToPrint;
            RSet strDataToPrint = Format(curTransactionAmount, "currency")
            Printer.Print strDataToPrint;
            RSet strDataToPrint = Format(curUserChequingAccountBalance, "currency")
            Printer.Print strDataToPrint;
            RSet strDataToPrint = Format(curUserSavingsAccountBalance, "currency")
            Printer.Print strDataToPrint;
            RSet strDataToPrint = Format(intTellerBalance, "currency")
            Printer.Print strDataToPrint
       End If
    Loop
    'close the file and add end of report
    Close
    Printer.Print
    Printer.FontSize = 16
    Printer.Print Space(6); "End of Report"
End Sub
