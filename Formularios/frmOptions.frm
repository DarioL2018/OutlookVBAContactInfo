VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptions 
   Caption         =   "Configuración"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5235
   OleObjectBlob   =   "frmOptions.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objExcelApp As Excel.Application

Private Sub btnOk_Click()
    Dim myFile As String, text As String, textline As String
    Dim myNameSpace As Outlook.NameSpace
    Dim myFolder As Outlook.Folder
    Dim myDistList As Outlook.DistListItem
    Dim myFolderItems As Outlook.Items
    
    'Contact Folders
    Dim colAL As Outlook.AddressLists
    'Contact Folder
    Dim oAL As Outlook.AddressList
    Dim colAE As Outlook.AddressEntries
    Dim oAE As Outlook.AddressEntry
    Dim oExUser As Outlook.ExchangeUser

    Dim x As Integer
    Dim y As Integer
  
    Dim email As String
    Dim strPath As String
    Dim strFilename As String
    
    Dim objExcelWorkBook As Excel.Workbook
    Dim objExcelWorkSheet As Excel.Worksheet
    Dim excelRow As Integer
    
    Dim dict As Scripting.Dictionary

    If txtOriginFile.text = "" Or txtReportPath.text = "" Or txtReport.text = "" Then
        MsgBox "Please, complete the information"
    
    Else
    
    myFile = txtOriginFile.text
    
    'Open Plain text File
    'Open myFile For Input As #1
    
    excelRow = 1
    strPath = txtReportPath.text & "\"
    strFilename = strPath & txtReport.text
    
    'Create Sheet on Excel
    Set objExcelWorkBook = objExcelApp.Workbooks.Add
    Set objExcelWorkSheet = objExcelWorkBook.Worksheets(1)
    
    'Insert Header
    insertHeaderOnExcel objExcelWorkSheet
            
    Set colAL = Application.Session.AddressLists
    loadFile dict
    searchListContacts dict, colAL, objExcelWorkSheet
    
    
    'read file
'    Do Until EOF(1)
'        Line Input #1, textline
        'read each email on file
'        email = getMail(textline)
        'search user
'        If Len(email) > 0 Then
'        Set oExUser = searchContact(colAL, email)
'        If (Not oExUser Is Nothing) Then
            'Add to Excel Report
 '           excelRow = excelRow + 1
 '           insertRowOnExcel objExcelWorkSheet, excelRow, oExUser
'        End If
 '       End If
'    Loop
'    Close #1
    'AutofitAllUsed objExcelWorkSheet
    'Save Excel Report
    objExcelWorkBook.SaveAs strFilename
    'Close Excel Application
    objExcelWorkBook.Close True
    
    MsgBox "Report generated"
    End If
End Sub

Sub loadFile(dict As Scripting.Dictionary)
    
    Dim myFile As String
    Dim email As String
    Dim textline As String
    
    Set dict = New Scripting.Dictionary
    email = ""
    texline = ""
    myFile = txtOriginFile.text
    
    'Open Plain text File
    Open myFile For Input As #1
    
    'Read File
    Do Until EOF(1)
        Line Input #1, textline
        'read each email on file
        email = getMail(textline)
        'search user
        If Len(email) > 0 Then
            dict.Add UCase(email), email
        End If
    Loop
    Close #1

End Sub

Sub searchListContacts(hashList As Scripting.Dictionary, colAL As Outlook.AddressLists, objExcelWorkSheet As Excel.Worksheet)
     'Contact Folder
    Dim oAL As Outlook.AddressList
    Dim colAE As Outlook.AddressEntries
    Dim oAE As Outlook.AddressEntry
    Dim oExUser As Outlook.ExchangeUser
    Dim excelRow As Integer
       
    excelRow = 1

    Set oExUser = Nothing
    
 For Each oAL In colAL
'Address list is an Exchange Global Address List
    If oAL.AddressListType = olExchangeGlobalAddressList Then
        Set colAE = oAL.AddressEntries
        'Loop each user
        For Each oAE In colAE
            If oAE.AddressEntryUserType = olExchangeUserAddressEntry Then
                Set oExUser = oAE.GetExchangeUser
                'If hashList.Exists(UCase(oExUser.PrimarySmtpAddress)) Then
                 If existMail(oExUser, hashList) Then
                    'return user object
                    excelRow = excelRow + 1
                    insertRowOnExcel objExcelWorkSheet, excelRow, oExUser
                    'hashList.Remove (UCase(oExUser.PrimarySmtpAddress))
                    If hashList.Count <= 0 Then
                        GoTo ExitLoop:
                    End If
                 End If
                 Set oExUser = Nothing
            End If
        Next
    End If
    Next
ExitLoop:
 'Set searchContact = oExUser
    
End Sub

Function existMail(oExUser As Outlook.ExchangeUser, hashList As Scripting.Dictionary) As Boolean
    Dim stringArray() As String
    Dim splitArray() As String
    Dim mailStr As String
    Dim y As Integer
    Dim resultado As Boolean
    resultado = False
    
    If hashList.Exists(UCase(oExUser.PrimarySmtpAddress)) Then
        hashList.Remove (UCase(oExUser.PrimarySmtpAddress))
        resultado = True
    Else
    stringArray() = oExUser.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x800F101F")
        For y = LBound(stringArray) To UBound(stringArray)
            splitArray = Split(UCase(stringArray(y)), "SMTP:")
            If (UBound(splitArray) > 0) Then
                If hashList.Exists(splitArray(1)) Then
                    resultado = True
                    hashList.Remove (UCase(splitArray(1)))
                    GoTo ExitLoopMail:
                End If
            End If
        Next
    End If
    
ExitLoopMail:
existMail = resultado

End Function

Function searchContact(colAL As Outlook.AddressLists, ByVal email As String) As Outlook.ExchangeUser
    'Contact Folder
    Dim oAL As Outlook.AddressList
    Dim colAE As Outlook.AddressEntries
    Dim oAE As Outlook.AddressEntry
    Dim oExUser As Outlook.ExchangeUser
    Set oExUser = Nothing
 
 For Each oAL In colAL
'Address list is an Exchange Global Address List
    If oAL.AddressListType = olExchangeGlobalAddressList Then
        Set colAE = oAL.AddressEntries
        'Loop each user
        For Each oAE In colAE
            If oAE.AddressEntryUserType = olExchangeUserAddressEntry Then
                Set oExUser = oAE.GetExchangeUser
                If oExUser.PrimarySmtpAddress = email Then
                    'return user object
                   GoTo ExitLoop:
                 End If
            End If
        Next
    End If
    Next
ExitLoop:
 Set searchContact = oExUser
End Function

'Insert header on excel file
Sub insertHeaderOnExcel(ws As Excel.Worksheet)
    ws.Cells(1, 1) = "Name"
    ws.Cells(1, 2) = "Email"
    ws.Cells(1, 3) = "Location"
    ws.Cells(1, 4) = "Grade Global"
    ws.Cells(1, 5) = "Grade Local"
    ws.Cells(1, 6) = "membership list"
    ws.Range("A1:E1").Font.Bold = True
End Sub
'Insert contact information
Sub insertRowOnExcel(ws As Excel.Worksheet, ByVal rowIndex As Integer, oExUser As Outlook.ExchangeUser)
    Dim y As Integer
    Dim group As Outlook.AddressEntry
    y = 0
    ws.Cells(rowIndex, 1) = oExUser.Name
    ws.Cells(rowIndex, 2) = oExUser.PrimarySmtpAddress
    ws.Cells(rowIndex, 3) = oExUser.OfficeLocation
    
    For Each group In oExUser.GetMemberOfList
        y = y + 1
        ws.Cells(rowIndex, 5 + y) = group.Name
        If (InStr(UCase(group.Name), "GLOBAL GRADE") > 0) Then
             ws.Cells(rowIndex, 4) = group.Name
        End If
        If (InStr(UCase(group.Name), "CAPGEMINI.GRADO") > 0) Then
             ws.Cells(rowIndex, 5) = group.Name
        End If
    Next
End Sub


Sub AutofitAllUsed(ws As Excel.Worksheet)
 
Dim x As Integer
 
For x = 1 To ws.UsedRange.Columns.Count
 
     Columns(x).EntireColumn.AutoFit
 
Next x
 
End Sub
Private Sub btnPath_Click()
    Dim fd As Office.FileDialog
    'Set fd = objExcelApp.Application.FileDialog(msoFileDialogFilePicker)
    Set fd = objExcelApp.FileDialog(msoFileDialogFilePicker)
    fd.Show
    
    If fd.SelectedItems.Count > 0 Then
        txtOriginFile.text = fd.SelectedItems(1)
    End If
    
End Sub

Private Sub btnPath2_Click()
    Dim fd As Office.FileDialog
    Set fd = objExcelApp.FileDialog(msoFileDialogFolderPicker)
    fd.Show
    If fd.SelectedItems.Count > 0 Then
        txtReportPath.text = fd.SelectedItems(1)
    End If
End Sub


Private Sub txtCancel_Click()
Unload Me
End Sub
Private Sub UserForm_Initialize()
    Set objExcelApp = CreateObject("Excel.Application")
End Sub

Function getMail(substring As String) As String
Dim startSymbol As Integer '<
Dim endSymbol As Integer   '>
Dim result As String
Dim LArray() As String
 
result = ""
 
startSymbol = InStr(substring, "<") - 1
endSymbol = InStr(substring, ">") - 1
 
If startSymbol > 0 And endSymbol > startSymbol Then
    LArray = Split(substring, "<")
    result = Split(LArray(1), ">")(0)
Else
    result = substring
End If
getMail = result
End Function

