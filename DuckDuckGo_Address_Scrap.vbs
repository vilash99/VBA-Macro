Sub GoogleAutomatedSearch()
    
    If MsgBox("Are you sure to start scraping?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Application.DisplayStatusBar = True
    
    Dim I As Integer, URL As String

    Dim IE As Object
    Dim ieDoc As HTMLDocument
    Dim myElem As IHTMLElement, tmpElem As IHTMLElement
     

    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = False 'Prevent the IE window from showing up
        
        
    '''Loop through C1 to E1 given values
    For I = Sheets("DucDuckGo").Range("C1").Value To Sheets("DucDuckGo").Range("E1").Value
       'On Error Resume Next
    
        '''Remove special characters
        '''Ex url: https://duckduckgo.com/?q=Amilinda%2C+Milwaukee%2C+WI
        '''Working for MAP: &iaxm=maps
        With Application.WorksheetFunction
            URL = "https://duckduckgo.com/?q=" & .Substitute(.Substitute(.Substitute(Sheets("DucDuckGo").Range("A" & I).Value, " ", "+"), "&", "%26"), ",", "%2C") & "&iaxm=maps"
        End With
       
       
        '''Show Status
        Application.StatusBar = "Macro is running ... Now at row : " & I & " / " & Sheets("DucDuckGo").Range("E1").Value & "... Last search made at : " & Now
        
        
        IE.navigate URL
        Do While IE.readyState <> 4 Or IE.Busy = True
            DoEvents
        Loop
 
        Set ieDoc = IE.document
        Application.Wait (Now() + TimeValue("00:00:02"))


        Dim companyName As String, companyAddress As String, companyContact As String, companyWebsite As String
        
        ''''Initlize to blanks
        companyName = ""
        companyAddress = ""
        companyContact = ""
        companyWebsite = ""
        
        
        '''''Check if there are multiple results
        For Each tmpElem In ieDoc.getElementsByClassName("module__title place-detail__name")
           companyName = tmpElem.innerText
           Exit For
        Next tmpElem
        
        If companyName = "" Then
            '''Click on first result
            For Each myElem In ieDoc.getElementsByClassName("place-list-item__title")
                If LCase(myElem.tagName) = "h2" Then
                    
                    myElem.Click
                    
                    ''''Check if item is clicked
                    For Each tmpElem In ieDoc.getElementsByClassName("module__title place-detail__name")
                        companyName = tmpElem.innerText
                        Exit For
                    Next tmpElem
                    
                    If companyName <> "" Then
                        Exit For
                    End If
                End If
            Next myElem
        End If
        
        
        '''''Extract Company Name
        For Each myElem In ieDoc.getElementsByClassName("module__title place-detail__name")
           companyName = myElem.innerText
           Exit For
        Next myElem
        
        '''Extract Company Address
        For Each myElem In ieDoc.getElementsByClassName("place-detail__data")
            companyAddress = myElem.innerText
            Exit For
        Next myElem
        
        '''Extract Company Contact
        For Each myElem In ieDoc.getElementsByClassName("js-place-detail-phone")
            companyContact = myElem.innerText
            Exit For
        Next myElem
        
        '''Extract Company Website
        For Each myElem In ieDoc.getElementsByClassName("js-place-detail-website")
            companyWebsite = myElem.getAttribute("href")
            Exit For
        Next myElem
                
        
        '''''''''Seprate Address from HTML''''''''''
        Dim tmpData() As String
        
        If InStr(companyAddress, vbCrLf) > 0 Then
            tmpData = Split(companyAddress, vbCrLf)
            companyAddress = tmpData(0)
        End If
        
        If InStr(companyAddress, ": ") > 0 Then
            tmpData = Split(companyAddress, ": ")
            companyAddress = tmpData(1)
        End If
        
        Sheets("DucDuckGo").Range("B" & I).Value = companyName
        Sheets("DucDuckGo").Range("C" & I).Value = companyAddress
        Sheets("DucDuckGo").Range("I" & I).Value = companyContact
        Sheets("DucDuckGo").Range("J" & I).Value = companyWebsite
        
        '''Wait for Specific Time
        Application.Wait (Now() + TimeValue("00:00:" & Sheets("DucDuckGo").Range("G1").Value))
        
        'Save current scrape data
        ThisWorkbook.Save
    Next I
    
    IE.Quit
    Set myElem = Nothing
    Set ieDoc = Nothing
    Set IE = Nothing
    
    
    Application.StatusBar = ""
    Sheets("DucDuckGo").Activate
    MsgBox "Completed !", vbInformation, "Automated Google Search"
    
End Sub
