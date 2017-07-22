


Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwmilliseconds As Long)
Public packnum As Double
Public packlist As String
Public packcoll As Collection
Public pos As Collection
Public serial As Collection
Public qty As Collection
Public contqty As Collection
Public cont As Collection
Public ackpos As Collection
Public fullpos As Collection


Sub ASNGenerator()

'Input Numbers for ASN
    Set packcoll = New Collection
    x = CInt(InputBox("How many Packlists need an ASN?"))
    y = 0
    While y < x
        packcoll.Add InputBox("What is the Packlist Number?")
        y = y + 1
    Wend
    
    For Each it In packcoll
        messbox = messbox & " " & it
    Next it
    
    messagebox = "The ASN has been completed for the following packlist numbers: " & messbox
    
    packnum = 1
    
'Cross check pos for mistakes

'    Dim objsw As shdocvw.ShellWindows
'    Dim pocheckie As New InternetExplorer
'
'for each
'    pocheckie.navigate "http://rpm:8600/dropbox/rpm5/cvoShipping/packlistprintnew.cfm?packListNbr=" + CStr(packcoll(packnum))
'
'    Refocus RPM site - breaks on load
'    sleep (300)
'    Set objsw = New ShellWindows
'    sleep (200)
'    Set pocheckie = objsw.Item(objsw.Count - 1)
'
'    Set it = pocheckie.document.getElementsByTagName("h3")
        
    
'Declare Variables
    Dim objIE As InternetExplorer

nextasn:
    counter = 0
    
    Set ackpos = New Collection
    

'Select PO numbers VIA RPM website

    POGenerator (packnum)

'Initiate a new instance of IE and assign to objIE
    Set objIE = New InternetExplorer

'Make IE browser visible
    objIE.Visible = True

'Navigate IE to HDSN
    objIE.navigate "Https://www.h-dsn.com/hdrtns/Welcome.jsp"

'Wait for webpage to load
    Do While objIE.Busy = True Or objIE.readyState <> 4: DoEvents: Loop

'Input username into Username box - error reads when already logged in
    If objIE.LocationName <> "H-DSN Log-On" Then GoTo loggedin:
    objIE.document.getElementById("USER").value = "shipping@gcpaint.com"
    
'Input password
    objIE.document.getElementsByTagName("input")(2).value = "Gcpaint@9"

'Click 'go' button
    objIE.document.forms(0).all("login").Click

'Wait for browser to load
    Do While objIE.Busy = True Or objIE.readyState <> 4: DoEvents: Loop

'Navigate to Firm Order Visibility page
loggedin:
    objIE.navigate "https://www.h-dsn.com/svc/firm/firmMain.do?supplier=204983"

'Wait again for the browser
    Do While objIE.Busy = True Or objIE.readyState <> 4: DoEvents: Loop

'Select shipping plant, and open/partial line status
    For Each it In objIE.document.getElementsByTagName("option")
        If it.value = "Open" Then
            it.Selected = True
        End If
        If it.value = "Partial" Then
            it.Selected = True
        End If
        If it.value = "1001" Then
            it.Selected = True
            Exit For
        End If
    Next it

'Submit selections of issuer and line status
    objIE.document.getElementsByClassName("LinkBgrnd")(0).Click

'Wait again for the browser
    Do While objIE.Busy = True Or objIE.readyState <> 4: DoEvents: Loop

'Sort page by date
    objIE.document.getElementsByName("shipDateImg")(0).PreviousSibling.PreviousSibling.Click

'Wait again for the browser
    Do While objIE.Busy = True Or objIE.readyState <> 4: DoEvents: Loop
    
'Select PO numbers shipped based on pos variable
    For Each it In pos
        ackpos.Add it
    Next it
    x = 0
    counter = 0
    For Each it In objIE.document.getElementsByTagName("input")
        If it.Type = "checkbox" Then
            counter = 0
            For Each po In ackpos
                counter = counter + 1
                If CStr(po) = Mid(it.value, 6, 10) Then
                    If it.ParentNode.ParentNode.nextElementSibling.nextElementSibling.nextElementSibling.Children(0).innerText = "U " Then
                        it.Checked = True
                        ackpos.Remove (counter)
                        x = 1
                    End If
                End If
            Next po
        End If
    Next it

'Click Acknowledge button
    If x <> 1 Then GoTo selectpos
        
    
    
    
    For Each it In objIE.document.getElementsByTagName("img")
        If it.Name = ("ack") Then
            sleep (200)
            it.Click
            Exit For
        End If
    Next it
    
'Wait for browser
    Do While objIE.Busy = True Or objIE.readyState <> 4: DoEvents: Loop
    
'Sort page by date
    objIE.document.getElementsByName("shipDateImg")(0).PreviousSibling.PreviousSibling.Click

'Wait for browser
    Do While objIE.Busy = True Or objIE.readyState <> 4: DoEvents: Loop
    
'Select POS for creating ASN
selectpos:
Set ackpos = New Collection
    For Each it In pos
        ackpos.Add it
    Next it
    counter = 0
    For Each it In objIE.document.getElementsByTagName("input")
        If it.Type = "checkbox" Then
            counter = 0
            For Each po In ackpos
                counter = counter + 1
                If CStr(po) = Mid(it.value, 6, 10) Then
                    it.Checked = True
                    ackpos.Remove (counter)
                    Exit For
                End If
            Next po
        End If
        If ackpos.Count < 1 Then Exit For
    Next it
    
'Check for orders that are missing from HDSN - If so set variable to trigger Draft instead of Send
    If ackpos.Count > 0 Then
        For Each it In ackpos
            msg = msg & it & vbNewLine
            incompleteasn = 1
        Next it
        MsgBox (msg + " Order does not exist on H-DSN. Can not send current packlist:" + CStr(packcoll(packnum)) + ". Notify shipping coordinators at Harley and GCP before continuing.")
    End If
    
'Click Create ASN
    
    sleep (200)
    For Each it In objIE.document.getElementsByTagName("img")
        If it.Name = ("createAsn") Then
            it.Click
            Exit For
        End If
    Next it
    
'Wait for browser
    Do While objIE.Busy = True Or objIE.readyState <> 4: DoEvents: Loop
    
'Fill out ASN Details
    For Each it In objIE.document.getElementsByTagName("input")
        If it.Name = "asnNum" Then
            it.value = packcoll(packnum)
            Exit For
        End If
    Next it

'Populate all Packlist numbers
    For Each it In objIE.document.getElementsByTagName("img")
        If it.Name = "yes" Then
            it.Click
            Exit For
        End If
    Next it

'Enter BOL and Tracking number
    For Each it In objIE.document.getElementsByTagName("input")
        If it.Name = "BOL" Then it.value = Format(Date, "mmddyyyy")
        If it.Name = "trackNum[0]" Then it.value = Format(Date, "mmddyyyy")
    Next it
    
'Update SCAC code
    For Each it In objIE.document.getElementsByTagName("input")
        If it.Name = "scacCode" Then
            If Weekday(Date, vbSunday) = 4 Then
                it.value = "WENP"
            Else
                it.value = "SCNN"
            End If
            Exit For
        End If
    Next it
    
'Remove Auto Serial Generation
    For Each it In objIE.document.getElementsByTagName("input")
        If it.Name = ("autoSerial") Then
            it.Checked = True
            Exit For
        End If
    Next it
    
'Input the ASN QTY and Cont QTY data
    counter = 0
skip3:
    For Each it In objIE.document.getElementsByTagName("input")
        x = 0
        If it.Name = "lineItem[" & counter & "].poPickNum" Then
            For Each po In pos
                x = x + 1
                If it.value = CStr(po) Then
                    For Each it2 In objIE.document.getElementsByTagName("input")
                        If it2.Name = "lineItem[" & counter & "].qty" Then
                            If it2.value = CStr(qty(x)) Then
                                GoTo skip5
                            End If
                            uniquekey = Mid(it2.onblur, 76, 7)
                            it2.value = qty(x)
                            it2.FireEvent ("onblur")
                            Do While objIE.Busy = True Or objIE.readyState <> 4: DoEvents: Loop
                            sleep (200)
skip5:                      For Each it3 In objIE.document.getElementsByTagName("input")
                                If it3.Name = "lineItem[" & counter & "].container[0].containerQty" Then
                                    If it3.value = CStr(contqty(x)) Then
                                        counter = counter + 1
                                        GoTo skip3
                                    End If
                                    uniquekey = Mid(it3.onblur, 76, 7)
                                    it3.value = contqty(x)
                                    it3.FireEvent ("onblur")
                                    Do While objIE.Busy = True Or objIE.readyState <> 4: DoEvents: Loop
                                    sleep (200)
                                    counter = counter + 1
                                    GoTo skip3
                                End If
                            Next it3
                        End If
                    Next it2
                End If
            Next po
        End If
    Next it
    
'Input the MHU numbers
    counter = 0
    Dim objie2 As InternetExplorer
    Set objie2 = New InternetExplorer

    
While counter < pos.Count
resetpopup:
    sleep (200)
    For Each it In objIE.document.getElementsByTagName("img")
        If it.Name = "view_" + CStr(counter) Then
            it.Click
            Exit For
        End If
    Next it

'Target popup to input MHUS

    On Error GoTo resetpopup:
    Set objsw = New shdocvw.ShellWindows
    Set objie2 = Nothing
    For Each it In objsw
        If it.LocationName = "ASN Visibility" Then Set objie2 = it
        sleep (30)
    Next it
    On Error GoTo 0
    
'ASN Visibility
uniqueposet:
    it3 = 0
    it2 = 1
    If objie2 Is Nothing Then GoTo resetpopup
    For Each it In objie2.document.getElementsByTagName("input")
        If it.Name = "uniqueKey" Then
            sleep (10)
            uniquepo = Left(it.value, 10)
            For Each po In pos
                If uniquepo = po Then
                    it3 = 1
                    Exit For
                End If
                it2 = it2 + 1
            Next po
        End If
    Next it
    
'Enter the MHU's
    y = 1
    x = 0
    
    For Each po In fullpos
        If uniquepo = po Then Exit For
        y = y + 1
    Next po
    
    
    If it3 = 0 Then GoTo uniqueposet:
    While x < cont(it2)
        Set it = Nothing
        sleep (55)
        On Error GoTo skip3:
        Set it = objie2.document.getElementById("TMHU[" + CStr(x) + "]")
        On Error GoTo 0
        it.Focus
        it.value = serial(y + x)
        Set it = Nothing
        x = x + 1
    Wend
    
'Save MHU Numbers
    For Each it In objie2.document.getElementsByTagName("a")
        If it.href = "javascript:imgClick('save'); submitForm('save');" Then
            it.Click
            Exit For
        End If
    Next it
    sleep (500)
    counter = counter + 1
    
Wend

'Shipment Details
    Set totalcont = objIE.document.getElementById("totalCont")
    For Each it In objIE.document.getElementsByName("grossWt")
        If it.Name = "grossWt" Then
            it.Focus
            it.value = CVar(totalcont.innerText) * 100
        End If
    Next it
        For Each it In objIE.document.getElementsByName("pkg[0].total")
        If it.Name = "pkg[0].total" Then
            it.Focus
            it.value = CVar(totalcont.innerText)
        End If
    Next it
    
'Complete the ASN
    For Each it In objIE.document.getElementsByTagName("a")
        Select Case incompleteasn
            Case 0
                If it.href = "javascript:imgClick('send'); submitForm('send')" Then
                    it.Click
                    sleep (200)
                    Exit For
                End If
            Case 1
                If it.href = "javascript:imgClick('draft'); submitForm('draft');" Then
                    it.Click
                    sleep (200)
                    Exit For
                End If
        End Select
    Next it
      
    Do While objIE.Busy = True Or objIE.readyState <> 4: DoEvents: Loop
    
    If objIE.LocationName <> "H-DSN Order Management" Then MsgBox ("Double Check the ASN has sent properly")
    
'Loop to beginning to restart ASN based on packnum
    objIE.Quit
    sleep (1000)
    
    If packnum <> (packcoll.Count) Then
        packnum = packnum + 1
        GoTo nextasn
    End If
    
    
    MsgBox (messagebox)
    
End Sub

Sub POGenerator(i As Double)

    Dim myarray() As Variant
    Dim poholder As Object

    Set pos = New Collection
    Set serial = New Collection
    Set qty = New Collection
    Set contqty = New Collection
    Set cont = New Collection
    Set fullpos = New Collection

'Select PO numbers VIA RPM website
    Dim rpmIE As InternetExplorer
    Set rpmIE = New InternetExplorer
    rpmIE.Visible = True
    


'Navigate to appropriate Packslip page
    rpmIE.navigate "http://rpm:8600/dropbox/rpm5/cvoShipping/packlistprintnew.cfm?packListNbr=" + CStr(packcoll(packnum))
    
'Refocus RPM site - breaks on load
    sleep (300)
    Dim objsw As ShellWindows
    Set objsw = New ShellWindows
    sleep (200)
    Set rpmIE = objsw.Item(objsw.Count - 1)

'Wait for the browser
    Do While rpmIE.Busy = True Or rpmIE.readyState <> 4: DoEvents: Loop

'Generate PO's for packlist
    
    pos.Add "null", "null"
    
    For Each it In rpmIE.document.getElementsByName("partqty")
        Set poholder = it.ParentNode.PreviousSibling
        fullpos.Add (poholder.innerText)
        If poholder.innerText <> pos(pos.Count) Then
            pos.Add (poholder.innerText)
        End If
        Set poholder = it.ParentNode.ParentNode.NextSibling
        serial.Add (Left(Right(poholder.innerText, 9), 8))
    Next it
    
    pos.Remove "null"

    Set aligninfo = rpmIE.document.getElementsByName("partqty")
    

    For Each it In pos
        counter = 0
        Erase myarray
        For Each potentialpo In rpmIE.document.getElementsByName("partqty")
            Set poholder = potentialpo.ParentNode.PreviousSibling
            If it = poholder.innerText Then
                ReDim Preserve myarray(0 To counter)
                myarray(counter) = CInt(potentialpo.value)
                counter = counter + 1
            End If
        Next potentialpo
        qty.Add (WorksheetFunction.Sum(myarray))
        contqty.Add (myarray(0))
        cont.Add (counter)
    Next it

    
    
    On Error Resume Next
    
    rpmIE.Quit

End Sub




