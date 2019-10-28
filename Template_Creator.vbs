Option Explicit

Private Sub OpenWorkbook()
    If MsgBox("Are you sure to continue?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    If Cells(1, 2) <> "" Then
        Call ReadDataFromCloseFile(Cells(1, 2))
    Else
        MsgBox "File is not given", vbExclamation
    End If
End Sub


Sub ReadDataFromCloseFile(mFileName As String)
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    Dim src As Workbook
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open(mFileName, True, True)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''SB_Schedule Upload'''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''Get the total rows from the source workbook.
    Dim iCnt As Long, i As Long, iTotalRows As Long, j As Long
    Dim tmpText As String
    
    
    '''''-----------Get All Peoples in Array-----------------''''
    Dim aFName() As String, aLName() As String, aEmail() As String
    Dim TotalPeople As Long
    
    TotalPeople = src.Worksheets("People").Cells(Rows.Count, 1).End(xlUp).Row
    TotalPeople = TotalPeople - 1
    
    ReDim aFName(1 To TotalPeople + 1)
    ReDim aLName(1 To TotalPeople + 1)
    ReDim aEmail(1 To TotalPeople + 1)
    
    For i = 1 To TotalPeople
        aFName(i) = src.Worksheets("People").Range("C" & i + 1).Value
        aLName(i) = src.Worksheets("People").Range("D" & i + 1).Value
        aEmail(i) = src.Worksheets("People").Range("E" & i + 1).Value
    Next i
    '''''''------------getting all peoples in array-------------------------'''''''''''
    
    
    
    '''''''''''-----------Getting all Providers-------------''''''''''''''''''''''''''
    Dim projName() As String, projID() As String
    Dim TotalProject As Long
    
    TotalProject = src.Worksheets("Provider_Project").Cells(Rows.Count, 1).End(xlUp).Row
    TotalProject = TotalProject - 1
    
    ReDim projName(1 To TotalProject + 1)
    ReDim projID(1 To TotalProject + 1)
    
    For i = 1 To TotalProject
        projID(i) = src.Worksheets("Provider_Project").Range("A" & i + 1).Value
        projName(i) = src.Worksheets("Provider_Project").Range("B" & i + 1).Value
    Next i
    '''''''''''-----------End of Getting all Providers-------------''''''''''''''''''''''''''
    
    
    iTotalRows = src.Worksheets("Base schedule").Cells(Rows.Count, 1).End(xlUp).Row
    iTotalRows = iTotalRows - 10
   
    '''''Copy data from source (close workgroup) to the destination workbook.
    For iCnt = 2 To iTotalRows + 1
        'Copy date
        src.Worksheets("SB_Schedule Upload").Range("A" & iCnt).Value = "'" & src.Worksheets("Base schedule").Range("B" & iCnt + 9).Value

        'Copy Start Time
        tmpText = src.Worksheets("Base schedule").Range("E" & iCnt + 9).Value

        src.Worksheets("SB_Schedule Upload").Range("B" & iCnt).Value = "'" & Trim(Left(tmpText, InStr(tmpText, "-") - 1))

        'Copy End Time
        src.Worksheets("SB_Schedule Upload").Range("C" & iCnt).Value = "'" & Trim(Right(tmpText, InStr(tmpText, "-") - 1))

        'Get Location for searching
        tmpText = src.Worksheets("Base schedule").Range("D4").Value

        'Copy Provider
        If src.Worksheets("Base schedule").Range("K" & iCnt + 9).Formula = "No Provider Assigned" Then
            '''Find Associated project ID
            For i = 1 To TotalProject + 1
                If projName(i) = tmpText Then
                    src.Worksheets("SB_Schedule Upload").Range("D" & iCnt).Value = "No Provider Assigned_" & projID(i)
                    Exit For
                End If
            Next i
        Else
            src.Worksheets("SB_Schedule Upload").Range("D" & iCnt).Value = src.Worksheets("Base schedule").Range("K" & iCnt + 9).Value
        End If

        'Set Quantity
        src.Worksheets("SB_Schedule Upload").Range("E" & iCnt).Value = "1"

        'Set Location
        src.Worksheets("SB_Schedule Upload").Range("G" & iCnt).Value = tmpText


        '''Get Email with Same Name in People Tab
        Dim fName As String, lName As String
        Dim ABC As Variant

        fName = lName = ""
        tmpText = src.Worksheets("Base schedule").Range("G" & iCnt + 9).Value
        ABC = Split(tmpText, ",")

        lName = Replace(ABC(0), ",", "")
        fName = Replace(ABC(1), ",", "")
        tmpText = Trim(fName) & " " & Trim(lName)
        
        
        For i = 1 To TotalPeople + 1
            If aFName(i) & " " & aLName(i) = tmpText Then
                src.Worksheets("SB_Schedule Upload").Range("F" & iCnt).Value = aEmail(i)
                Exit For
            End If
        Next i
    Next iCnt
    
        
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''SB_People Upload'''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim empID() As String, empProject() As String, empProjectID() As String
    Dim TotalEmployee As Long
    Dim tmpEmpID As String
    Dim allProjects As String
    
    
    TotalEmployee = src.Worksheets("Employee_Project").Cells(Rows.Count, 1).End(xlUp).Row
    TotalEmployee = TotalEmployee - 1

    ReDim empProjectID(1 To TotalEmployee + 1)
    ReDim empID(1 To TotalEmployee + 1)
    ReDim empProject(1 To TotalEmployee + 1)

    For i = 1 To TotalEmployee
        empProjectID(i) = src.Worksheets("Employee_Project").Range("A" & i + 1).Value
        empID(i) = src.Worksheets("Employee_Project").Range("B" & i + 1).Value
        empProject(i) = src.Worksheets("Employee_Project").Range("C" & i + 1).Value
    Next i
    
    iTotalRows = src.Worksheets("People").Cells(Rows.Count, 1).End(xlUp).Row
    iTotalRows = iTotalRows - 1

    For iCnt = 2 To iTotalRows + 1
        'Copy ID
        tmpEmpID = src.Worksheets("People").Range("A" & iCnt).Value
        src.Worksheets("SB_People Upload").Range("A" & iCnt).Value = tmpEmpID

        'Copy FirstName
        src.Worksheets("SB_People Upload").Range("B" & iCnt).Value = src.Worksheets("People").Range("C" & iCnt).Value

        'Copy Last Name
        src.Worksheets("SB_People Upload").Range("C" & iCnt).Value = src.Worksheets("People").Range("D" & iCnt).Value

        'Copy Email
        src.Worksheets("SB_People Upload").Range("E" & iCnt).Value = src.Worksheets("People").Range("E" & iCnt).Value

        '''Search Employee ID in Sheet
        allProjects = ""

        If tmpEmpID = "" Then Exit For

        For i = 1 To TotalEmployee + 1
            If empID(i) = tmpEmpID Then
                'Find Provider
                If empProject(i) = "No Provider Assigned" Then
                    allProjects = allProjects + "No Provider Assigned_" & empProjectID(i) & "|"
                Else
                    allProjects = allProjects + empProject(i) & "|"
                End If
            End If
        Next i

        If allProjects <> "" Then
            allProjects = Left(allProjects, Len(allProjects) - 1)
        End If

        src.Worksheets("SB_People Upload").Range("F" & iCnt).Value = allProjects
    Next iCnt
        
    
    
    ' CLOSE THE SOURCE FILE.
    src.Close True             ' Tue - Ask to Save
    Set src = Nothing
    
    Application.ScreenUpdating = True
    
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
