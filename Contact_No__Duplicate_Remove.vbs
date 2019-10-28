Option Explicit

Sub OpenWorkbook()
    If MsgBox("Are you sure to continue?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    If Cells(1, 4) <> "" Then
        Call ReadDataFromCloseFile(Cells(1, 4))
        
        MsgBox "Duplicate phone numbers are removed successfully!", vbInformation
    Else
        MsgBox "File path is not given!", vbExclamation
    End If
End Sub


Sub ReadDataFromCloseFile(mFileName As String)
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    Dim src As Workbook
    Dim app As New Excel.Application
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = app.Workbooks.Open(mFileName, True, True)
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''Read All Phone Number''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''Get the total rows from the source workbook.
    Dim iCnt As Long, jCnt As Long, i As Long
    
    '''''-----------Get All Master Phone Numbers in Array-----------------''''
    Dim masterPhoneNo() As String
    Dim totalNumbers As Long
    
    totalNumbers = src.Worksheets("Phones").Cells(Rows.Count, 1).End(xlUp).Row
    totalNumbers = totalNumbers - 1
    
    ReDim masterPhoneNo(1 To totalNumbers + 1)
    
    For i = 1 To totalNumbers
        masterPhoneNo(i) = src.Worksheets("Phones").Range("A" & i + 1).Value
    Next i
    ''''----------End of getting all phone numbers from Master Excel----------''''''
    
    
    '''''-----------Get Phone to check in Array to check in master file-----------------''''
    Dim checkPhoneNo() As String
    Dim totalCheckNumbers As Long
    
    totalCheckNumbers = ActiveWorkbook.Worksheets("Main_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
    totalCheckNumbers = totalCheckNumbers - 1
    
    ReDim checkPhoneNo(1 To totalCheckNumbers + 1)
    
    For i = 1 To totalCheckNumbers
        checkPhoneNo(i) = ActiveWorkbook.Worksheets("Main_Sheet").Range("A" & i + 1).Value
        
        ''Remove entry
        ActiveWorkbook.Worksheets("Main_Sheet").Range("A" & i + 1).Value = ""
    Next i
    ''''----------End of getting all check phone numbers from working Sheet----------''''''
    
    
    ''''''''''Check for duplicates phone numbers and remove them'''''''''''''
    Dim uniquePhoneNo() As String
    Dim uniquePhoneNoCount As Long
    
    ReDim uniquePhoneNo(1 To totalCheckNumbers + 1)
    
    Dim foundPhone As Boolean
    uniquePhoneNoCount = 0
    
    For iCnt = 1 To totalCheckNumbers
        foundPhone = False
        For jCnt = 1 To totalNumbers
            If checkPhoneNo(iCnt) = masterPhoneNo(jCnt) Then
                foundPhone = True
                Exit For
            End If
        Next jCnt
        
        '''''''Number is not duplicate, add to new list
        If foundPhone = False Then
            uniquePhoneNoCount = uniquePhoneNoCount + 1
            uniquePhoneNo(uniquePhoneNoCount) = checkPhoneNo(iCnt)
        End If
    Next iCnt
    
     
    
    '''''''''Add new unique phone nubers from array'''''''''''''
    For iCnt = 2 To uniquePhoneNoCount + 1
        Cells(iCnt, 1) = uniquePhoneNo(iCnt - 1)
    Next iCnt
    
    
    ' CLOSE THE SOURCE FILE.
    src.Close SaveChanges:=False            ' Tue - Ask to Save
    Set src = Nothing
    app.Quit
    Set app = Nothing
    
    Application.ScreenUpdating = True
    
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
