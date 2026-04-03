'Written by Wyatt Sustaita, JITM2142

Public Session As Object
Public Screen As Object

Public Function InitializeSession()

    Dim Quick As Object
    Dim myString As String
    Dim result   As Boolean

    'Creates a Quick3270 object
    Set Quick = GetObject("redacted for privacy")

    'Retreives the Session object
    Set Session = Quick.ActiveSession


    'Retrieves the Screen Object
    Set Screen = Session.Screen

    'Make the Window Visible, default is invisible
    Quick.Visible = True


End Function

Public Function LudicrousMode(switch As Boolean)

    If switch Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.DisplayStatusBar = False
        Application.EnableEvents = False
        
    ElseIf Not switch Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayStatusBar = True
        Application.EnableEvents = True
        
    Else
        'Boy quit boolin
        Exit Function
        
    End If
    
End Function

Public Function CheckMainScreen(switch As Boolean)

    If Trim(Session.Screen.getstring(2, 33, 6)) <> "USERID" Or Trim(Session.Screen.getstring(1, 28, 25)) <> "J.B. HUNT TRANSPORT, INC." Then
        MsgBox "Go To The Main Screen And Try Again"
        Call LudicrousMode(False)
        switch = False
        Exit Function
        
    End If

End Function

Sub BillFill_Click()

    Dim NextRow As Integer
    Dim SealNum As String
    Dim POnum As String
    Dim StNumb As Integer
    Dim result As Boolean
    Dim ordNum As String

    InitializeSession

    'Check to see if host is on 1 screen
    Dim good As Boolean
    good = True
    Call CheckMainScreen(good)
    
    If good = False Then
        Exit Sub
    
    End If
    
    'Pull order number from clipboard
    Call Copy_Text_From_Clipboard(ordNum)
    
    result = Screen.moveTo(30, 49)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "1"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys ordNum
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    
    Sheet2.Range("m12").Value = Trim(Session.Screen.getstring(1, 6, 9))
    Sheet2.Range("d7").Value = Trim(Session.Screen.getstring(15, 10, 20))
    Sheet2.Range("m9").Value = Trim(Session.Screen.getstring(10, 17, 6))
    Sheet2.Range("d4").Value = Trim(Session.Screen.getstring(4, 29, 25))
    Sheet2.Range("d5").Value = Trim(Session.Screen.getstring(5, 29, 25))
    Sheet2.Range("d6").Value = Trim(Session.Screen.getstring(8, 29, 25))
    Sheet2.Range("d9").Value = Trim(Session.Screen.getstring(4, 55, 25))
    Sheet2.Range("d10").Value = Trim(Session.Screen.getstring(5, 55, 25))
    Sheet2.Range("d11").Value = Trim(Session.Screen.getstring(8, 55, 25))
    Sheet2.Range("i27").Value = Trim(Session.Screen.getstring(13, 10, 5))
    Sheet2.Range("h32").Value = Trim(Session.Screen.getstring(2, 62, 18))
    Sheet2.Range("J27").Value = Trim(Session.Screen.getstring(13, 23, 5))
    Sheet2.Range("d14").Value = Trim(Session.Screen.getstring(9, 29, 25))
    Sheet2.Range("d15").Value = Trim(Session.Screen.getstring(9, 55, 25))
    Sheet2.Range("d16").Value = Trim(Session.Screen.getstring(10, 55, 25))
   
    Screen.SendKeys "<PF10>"
    result = Screen.WaitForKbdUnlock()
        
    Sheet2.Range("n3").Value = Trim(Session.Screen.getstring(18, 32, 19))
    Sheet2.Range("m51").Value = Trim(Session.Screen.getstring(18, 32, 19))

    Screen.SendKeys "<PF9>"
    result = Screen.WaitForKbdUnlock()
    
    Sheet2.Range("m10").Value = Trim(Session.Screen.getstring(8, 59, 11))

    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()
    
    Do While (Trim(Session.Screen.getstring(1, 8, 2)) <> "99")
        Screen.SendKeys "<PF8>"
        result = Screen.WaitForKbdUnlock()

    Loop

    Sheet2.Range("B57").Value = Trim(Session.Screen.getstring(22, 12, 10))
    Sheet2.Range("B58").Value = Trim(Session.Screen.getstring(22, 23, 10))
    Sheet2.Range("B59").Value = Trim(Session.Screen.getstring(22, 34, 10))
    Sheet2.Range("B60").Value = Trim(Session.Screen.getstring(22, 45, 10))
    Sheet2.Range("B61").Value = Trim(Session.Screen.getstring(22, 56, 10))
    Sheet2.Range("B62").Value = Trim(Session.Screen.getstring(22, 67, 10))
    Sheet2.Range("B63").Value = Trim(Session.Screen.getstring(23, 12, 10))
    Sheet2.Range("B64").Value = Trim(Session.Screen.getstring(23, 23, 10))
    Sheet2.Range("B65").Value = Trim(Session.Screen.getstring(23, 34, 10))
    Sheet2.Range("B66").Value = Trim(Session.Screen.getstring(23, 45, 10))
    Sheet2.Range("D12").Value = Trim(Session.Screen.getstring(19, 63, 10))

    Do While (Trim(Session.Screen.getstring(2, 33, 1)) <> "U")
        Screen.SendKeys "<PF12>"
        result = Screen.WaitForKbdUnlock()
    
    Loop
    
End Sub

Sub PrinterSel()

    Application.Dialogs(xlDialogPrinterSetup).Show

End Sub

Sub PrintMeh()

    ActiveSheet.PrintOut

End Sub


Sub ChassisBeam_Click()

    Dim Marker As New Collection
    Dim Cell As Variant
    Dim result As Boolean

    InitializeSession

    'Check to see if host on main screen
    Dim good As Boolean
    good = True
    Call CheckMainScreen(good)
    
    If good = False Then
        Exit Sub
    
    End If

    Call LudicrousMode(True)

    'Select all chassis'
    Range("A65000").Select
    Selection.End(xlUp).Select
    Range(Selection, "A2").Select
    
    'Add chassis' to collection
    For Each Cell In Selection
        Marker.Add Cell.Value
    Next Cell
    
    result = Screen.moveTo(30, 49)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "405"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    
    
    'Use each chassis in host screen
    For Each Item In Marker
        
        i = 0
        
        If ActiveCell.Value = "" Or Range("I7").Value = "" Or Range("I9").Value = "" Then
            Exit Sub
        
        Else
            result = Screen.moveTo(3, 26)
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys Item
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys "<ENTER>"
            result = Screen.WaitForKbdUnlock()
            
            If Trim(Session.Screen.getstring(29, 2, 17)) = "EQUIP IS ON ORDER" Then
                ActiveCell.Offset(0, 1).Value = Trim(Session.Screen.getstring(29, 2, 48))
                i = 1
            
                ElseIf Trim(Session.Screen.getstring(29, 2, 31)) = "*** INVALID  TRAILER NUMBER ***" Then
                ActiveCell.Offset(0, 1).Value = "Invalid trailer number"
                i = 1
            
                ElseIf Trim(Session.Screen.getstring(29, 2, 31)) = "EQUIP IS DISPATCHED-LOCATION CA" Then
                ActiveCell.Offset(0, 1).Value = "Equipment is dispatched, location cannot be changed"
                i = 1
            
                ElseIf Trim(Session.Screen.getstring(29, 2, 31)) = "EQUIPMENT ASSIGNED TO CONTAINER" Then
                ActiveCell.Offset(0, 1).Value = Trim(Session.Screen.getstring(29, 2, 39))
                i = 1
            
                ElseIf Trim(Session.Screen.getstring(5, 5, 1)) = "" Then
                ActiveCell.Offset(0, 1).Value = "Error updating trailer"
                i = 1
            
                Else
                i = 0
            
            End If
            
            If i = "0" Then
            result = Screen.moveTo(6, 69)
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys Range("I7")
            result = Screen.WaitForKbdUnlock()
            result = Screen.moveTo(7, 73)
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys Range("I9")
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys "<PF6>"
            result = Screen.WaitForKbdUnlock()
            
                If Trim(Session.Screen.getstring(29, 2, 23)) = "*** UPDATE COMPLETE ***" Then
                ActiveCell.Offset(0, 1).Value = "Success"
            
                Else
                ActiveCell.Offset(0, 1).Value = "Error updating trailer"
                End If
                
            End If
            
        End If
        
        ActiveCell.Offset(1, 0).Activate
        
    Next Item
    
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()
    
End Sub

Sub ChassisBreak_Click()

Dim Marker As New Collection
Dim Cell As Variant
Dim result As Boolean

InitializeSession

    
    'Check to see if host on 405 screen
    If Trim(Session.Screen.getstring(1, 44, 15)) <> "ATTACH / DETACH" Or Trim(Session.Screen.getstring(7, 3, 1)) <> "" Then
    
        If Trim(Session.Screen.getstring(1, 44, 15)) = "ATTACH / DETACH" And Trim(Session.Screen.getstring(7, 3, 9)) = "CONTAINER" Then
            Screen.SendKeys "<PF12>"
            result = Screen.WaitForKbdUnlock()
            result = Screen.moveTo(30, 49)
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys "217"
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys "<ENTER>"
            result = Screen.WaitForKbdUnlock()
            
        Else
            MsgBox "Go To The 217 Screen And Try Again"
            Exit Sub
            
        End If
        
    End If
    
    'Select all chassis'
    Range("A65000").Select
    Selection.End(xlUp).Select
    Range(Selection, "A2").Select
    
    'Add chassis' to collection
    For Each Cell In Selection
        Marker.Add Cell.Value
    Next Cell
    
    'Use each chassis in host screen
    For Each Item In Marker
        
        i = 0
        
        If ActiveCell.Value = "" Then
            Exit Sub
        
        Else
            result = Screen.moveTo(5, 27)
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys Item
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys "<ENTER>"
            result = Screen.WaitForKbdUnlock()
            
            If Trim(Session.Screen.getstring(30, 2, 30)) = "INVALID VALUE ENTERED IN FIELD" Then
                ActiveCell.Offset(0, 1).Value = Trim(Session.Screen.getstring(30, 2, 30))
                i = 1
                
            End If
            
            If i = "0" Then
            Screen.SendKeys "<PF6>"
            result = Screen.WaitForKbdUnlock()
            
                If Trim(Session.Screen.getstring(30, 2, 17)) = "UPDATE SUCCESSFUL" Then
                ActiveCell.Offset(0, 1).Value = "Success"
            
                Else
                ActiveCell.Offset(0, 1).Value = "Error breaking chassis from container"
                End If
                
            End If
            
        End If
        
        ActiveCell.Offset(1, 0).Activate
        
    Next Item
    
End Sub

Sub Clear_btn_Click()

    'specify sheets
    If ActiveSheet.name = "redacted for privacy" Then
        Exit Sub
    
    ElseIf ActiveSheet.name = "redacted for privacy" Or ActiveSheet.name = "redacted for privacy" Then
        Range("A2:B65000").ClearContents
        Range("I7").ClearContents
        Range("I9").ClearContents
        Range("A2:B65000").Interior.Color = RGB(191, 191, 191)
        Range("I7").Interior.Color = RGB(191, 191, 191)
        Range("I9").Interior.Color = RGB(191, 191, 191)
        Range("A2").Select
    
    ElseIf ActiveSheet.name = "redacted for privacy" Then
        Range("A:H").ClearContents
        Range("A:H").ClearFormats
        Range("A:H").Interior.Color = RGB(191, 191, 191)
        Range("A1").Select

    ElseIf ActiveSheet.name = "redacted for privacy" Then
        Range("A1").Value = "Activity Name"
        Range("A7").Select
    
        Do While (ActiveCell.Value <> "")
            ActiveCell.MergeArea.ClearContents
            ActiveCell.Offset(0, 1).MergeArea.ClearContents
            ActiveCell.Offset(0, 2).MergeArea.ClearContents
            ActiveCell.Offset(0, 3).MergeArea.ClearContents
            ActiveCell.Offset(1, 0).Activate
        
        Loop
        Range("A7").Select
    
    ElseIf ActiveSheet.name = "redacted for privacy" Then
        Range("A2:D65000").ClearContents
        Range("A2").Select
        
    ElseIf ActiveSheet.name = "redacted for privacy" Then
        Range("B2:C65000").ClearContents
        Range("B2").Select
        
    ElseIf ActiveSheet.name = "redacted for privacy" Then
        Range("A3:H65000").ClearContents
        Range("B3").Value = "Nothing to report at this time"
        Range("D3").Value = "Nothing to report at this time"
        Range("F3").Value = "Nothing to report at this time"
        Range("H3").Value = "Nothing to report at this time"
        Range("A3").Select
    
    ElseIf ActiveSheet.name = "redacted for privacy" Then
        Range("A3:B65000").ClearContents
        Range("B3").Value = "Nothing to report at this time"
        Range("A3").Select
        
    ElseIf ActiveSheet.name = "redacted for privacy" Then
        Range("A3:B65000").ClearContents
        Range("B3").Value = "Nothing to report at this time"
        Range("A3").Select
        
    ElseIf ActiveSheet.name = "redacted for privacy" Then
        Range("A3:B65000").ClearContents
        Range("B3").Value = "Nothing to report at this time"
        Range("A3").Select
    
    Else
        Exit Sub
    
    End If

End Sub

Sub OrgOrderSumm_Click()

Dim Marker As New Collection
Dim Cell As Variant
Dim refDate As Date

    'Establish refDate
    Dim wkd As Integer
    Dim X, y, z As Integer
    
    wkd = Weekday(Date, vbSunday)
    w = -7
    
    If wkd < 4 Then
        refDate = DateAdd("d", w, Date:=Date)
        
    End If
    
    refDate = DateAdd("d", -wkd, refDate)

    'Select all dates
    Range("C1000").Select
    Selection.End(xlUp).Select
    Range(Selection, "C1").Select
    
    'Add dates to collection
    For Each Cell In Selection
        Marker.Add Cell.Value
    Next Cell
    

    For Each Item In Marker

        If ActiveCell.Value = "" Then
        Exit Sub
        
        Else
            
            If Item = refDate Or Item = DateAdd("d", 1, refDate) And ActiveCell.Offset(0, 2).Value < 0.5 Then
                ActiveCell.Offset(0, 5).Value = "Old"
                With Range(ActiveCell.Offset(0, -2), ActiveCell.Offset(0, 5)).Interior
                    .ColorIndex = 4
                    
                End With
            
            Else
                ActiveCell.Offset(0, 5).Value = "Discount Double Check"
                'Alternative ways to set the cell background color
                With Range(ActiveCell.Offset(0, -2), ActiveCell.Offset(0, 5)).Interior
                    .ColorIndex = 22
                    
                End With
                
            End If
        
        End If
        
        ActiveCell.Offset(1, 0).Activate
        
        Next Item
        
        Range("A1").Select
        
End Sub

Sub boardDriverList_Click()

    Call Clear_btn_Click
    
    InitializeSession

    'check on 333 screen
    If Trim(Session.Screen.getstring(1, 26, 31)) <> "Drivers On/Off Duty Maintenance" Then
        MsgBox "Go To The 333 Screen And Try Again"
        Exit Sub
    End If

    Dim an As String
    Dim bc As String
    Dim i As Integer
    i = 3
    
    an = InputBox("Please input the name of the activity" & vbCrLf & vbCrLf & "This will become the header.", "Activity Name")
        If an = "" Then
            Exit Sub
        End If
    bc = InputBox("Please input the board code you wish to fill", "Board Code")
        If bc = "" Then
            Exit Sub
        End If
    Range("A1").Value = StrConv(an, 3)
    
    Do While i > 0
        If Len(bc) <> 3 Then
            bc = InputBox("Error, incorrect board code length. " & i & " attempts remaining." & vbCrLf & vbCrLf & _
            "Please input the board code you wish to fill", "Board Code")
            
            If i > 1 Then
                i = i - 1
            Else
                Exit Sub
            End If
            
        ElseIf Not IsNumeric(Mid(bc, 2, 1)) Or _
            (Not (Asc(Mid(bc, 1, 1)) > 64 And Asc(Mid(bc, 1, 1)) < 91) And Not (Asc(Mid(bc, 1, 1)) > 96 And Asc(Mid(bc, 1, 1)) < 123)) Or _
            (Not (Asc(Mid(bc, 3, 1)) > 64 And Asc(Mid(bc, 3, 1)) < 91) And Not (Asc(Mid(bc, 3, 1)) > 96 And Asc(Mid(bc, 3, 1)) < 123)) Then
            
            bc = InputBox("Error, invalid board code. " & i & " attempts remaining." & vbCrLf & vbCrLf & _
            "Please input the board code you wish to fill", "Board Code")
            
            If i > 1 Then
                i = i - 1
            Else
                Exit Sub
            End If
            
        Else
            i = 0
        End If
        
    Loop
    
    'bc ready to go here, or nah

    result = Screen.moveTo(2, 7)
    Screen.SendKeys bc
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    
    Dim y, j As Integer
    y = 5
    j = 1
    
    Dim str, str2 As String
    
    Range("A7").Select
    
    
    Do While y < 21
        If Trim(Session.Screen.getstring(y, 27, 1)) = "" Then
            Exit Do
        Else
            ActiveCell.Value = Trim(Session.Screen.getstring(y, 27, 7))
            str = Trim(Session.Screen.getstring(y, 2, 21))
            j = 1
            
            Do
            j = j + 1
            Loop Until Mid(str, j, 1) = ","
            
            str = Left(str, j) & " " & Mid(str, j + 1, Len(str))
            ActiveCell.Offset(0, 1).Value = StrConv(str, 3)
            ActiveCell.Offset(0, 2).Value = "____________"
            ActiveCell.Offset(0, 3).Value = "___________________________________________"
            ActiveCell.Offset(1, 0).Activate
            y = y + 1
        End If

        If y = 21 And Trim(Session.Screen.getstring(21, 28, 11)) <> "END OF FILE" Then
            Screen.SendKeys "<PF8>"
            result = Screen.WaitForKbdUnlock()
            y = 5
        End If
    Loop
    
    Range("A7").Select
    
    Screen.SendKeys "<PF5>"
    result = Screen.WaitForKbdUnlock()
    If Trim(Session.Screen.getstring(21, 21, 1)) = "U" Then
        Screen.SendKeys "<PF5>"
        result = Screen.WaitForKbdUnlock()
    End If
    
    MsgBox "Driver list ready!"
    
End Sub

Sub Update_Status_Codes_Local()
    
    Call Clear_btn_Click
    Call LudicrousMode(True)
    
    InitializeSession

    'check on main screen
    If Trim(Session.Screen.getstring(2, 33, 6)) <> "USERID" Or Trim(Session.Screen.getstring(1, 28, 25)) <> "redacted for privacy" Then
        MsgBox "Go To The Main Screen And Try Again"
        Call LudicrousMode(False)
        Exit Sub
        
    End If

    result = Screen.moveTo(30, 49)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "198"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "redacted for privacy"
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    result = Screen.moveTo(19, 9)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF4>"
    result = Screen.WaitForKbdUnlock()


    Dim a As Integer
    Dim st As String
    Dim stO As String
    Dim act As String
    Dim trk As String
    Dim alp As String
    Dim ntf As Boolean
    
    Range("A2").Select
    
    'add eof con
    
    Do Until (Trim(Session.Screen.getstring(31, 1, 1)) = "T" And Trim(Session.Screen.getstring(19, 9, 1)) = "_") Or _
            Trim(Session.Screen.getstring(31, 1, 1)) = "*"
        For a = 19 To 30
            st = ""
            stO = ""
            act = ""
            trk = ""
            alp = ""

            If Trim(Session.Screen.getstring(a, 9, 1)) = "" Then
                GoTo Ender2
            
            Else
                trk = Trim(Session.Screen.getstring(a, 2, 6))
                alp = Trim(Session.Screen.getstring(a, 9, 6))
                stO = Session.Screen.getstring(a, 38, 2)
                
                result = Screen.moveTo(a, 9)
                result = Screen.WaitForKbdUnlock()
                Screen.SendKeys "<PF10>"
                result = Screen.WaitForKbdUnlock()
                
                If Trim(Session.Screen.getstring(1, 29, 15)) = "MESSAGE DISPLAY" Then
                    Screen.SendKeys "<PF3>"
                    result = Screen.WaitForKbdUnlock()
                
                End If
                
                'Check msg and error additions
                If Trim(Session.Screen.getstring(19, 2, 18)) = "OBC MESSAGES EXIST" Then
                    act = ", check messages"
                    
                    If Trim(Session.Screen.getstring(19, 23, 10)) = "OBC ERRORS" Then
                        st = "??"
                        act = act & ", check errors"
                    
                    End If
                
                ElseIf Trim(Session.Screen.getstring(19, 3, 10)) = "OBC ERRORS" Then
                    st = "??"
                    act = ", check errors"
                    
                End If
                    
                'st checker
                If st = "??" Then
                    'do nothing, update manually
                    
                ElseIf stO = "1" Or stO = "2" Or stO = "3" Or stO = "4" Or stO = "5" Or stO = "6" Or _
                    stO = "7" Or stO = "8" Or stO = "9" Or stO = "??" Or stO = "10" Or stO = "11" Then
                    st = stO
                
                ElseIf Trim(Session.Screen.getstring(6, 62, 1)) = "" And Trim(Session.Screen.getstring(30, 63, 1)) = "3" Then
                    st = "1"
                    If st = stO Then
                        act = ", preplan needs dispatch"
                        
                    ElseIf stO = "MT" Then '*change this one for MT status
                        st = "2"
                        act = ", preplan needs dispatch"
                        
                    End If
                        
                ElseIf Trim(Session.Screen.getstring(30, 63, 2)) = "1" Then
                    st = "??"
                    act = ", ld " & Trim(Session.Screen.getstring(30, 73, 7)) & " still on train"
                        
                ElseIf Trim(Session.Screen.getstring(30, 63, 1)) = "J" And Trim(Session.Screen.getstring(6, 62, 1)) = "" Then
                    st = "2"
                        
                ElseIf Trim(Session.Screen.getstring(30, 49, 1)) = "3" Or (Trim(Session.Screen.getstring(30, 63, 1)) = "4" And _
                    Trim(Session.Screen.getstring(6, 62, 1)) <> "") Then
                    st = "5"
                    
                ElseIf Trim(Session.Screen.getstring(29, 31, 1)) = "L" Then
                    st = "6"
                        
                ElseIf Trim(Session.Screen.getstring(6, 62, 5)) = Trim(Session.Screen.getstring(7, 7, 5)) Or _
                        (Trim(Session.Screen.getstring(6, 62, 1)) <> "" And Trim(Session.Screen.getstring(6, 64, 1)) = "") Then
                    st = "7"
                        
                ElseIf Trim(Session.Screen.getstring(6, 62, 1)) = "" Then
                    st = "8"
                            
                Else
                    st = "9"
                        
                End If
                
                result = Screen.moveTo(8, 22)
                result = Screen.WaitForKbdUnlock()
                Screen.SendKeys st
                result = Screen.WaitForKbdUnlock()
                Screen.SendKeys "<ENTER>"
                result = Screen.WaitForKbdUnlock()
                Screen.SendKeys "<PF12>"
                result = Screen.WaitForKbdUnlock()
                
                If Trim(Session.Screen.getstring(1, 36, 1)) = "C" Then
                    Screen.SendKeys "<PF12>"
                    result = Screen.WaitForKbdUnlock()
                    Screen.SendKeys "<PF12>"
                    result = Screen.WaitForKbdUnlock()
                
                End If
                
                If stO = "  " Then
                    act = "Status updated from 'blank' to " & st & act
                
                ElseIf stO = st Then
                    act = "Keep status as " & stO & act
                
                Else
                    act = "Status updated from " & stO & " to " & st & act
                
                End If
                
            End If
            
Ender:

        ActiveCell.Value = trk
        ActiveCell.Offset(0, 1).Value = alp
        ActiveCell.Offset(0, 2).Value = act
        ActiveCell.Offset(1, 0).Activate

Ender2:
        
        If Trim(Session.Screen.getstring(31, 8, 1)) = "E" And a = 30 Then
            Exit Do
        End If
        
        Next a
        
        Screen.SendKeys "<PF8>"
        result = Screen.WaitForKbdUnlock()

    Loop
                
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()
    Range("A2").Select
    
    'Call Check_Notify
    Call LudicrousMode(False)
End Sub

Sub Check_Notify()

    InitializeSession
    
    If Trim(Session.Screen.getstring(2, 33, 6)) <> "USERID" Or Trim(Session.Screen.getstring(1, 28, 25)) <> "redacted for privacy" Then
        MsgBox "Go To The Main Screen And Try Again"
        Call LudicrousMode(False)
        Exit Sub
        
    End If

    Call LudicrousMode(True)
    
    Range("A2").Select
    
    'check for #notify on future load
    result = Screen.moveTo(30, 49)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "366"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    
    Dim moNow, daNow, tiNow, moApt, daApt, tiApt, moDif, daDif, tiDif, dif As Double
    Dim moA, moB, daA, daB, tiXA, tiXB, tiYA, tiYB As String
    Dim ld, notify As String
    
    Do Until ActiveCell.Value = ""
        result = Screen.moveTo(4, 14)
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys ActiveCell.Value
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
        
        If Trim(Session.Screen.getstring(4, 57, 1)) = "" Then
            GoTo EnderX
        
        End If
        
        result = Screen.moveTo(4, 57)
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<PF4>"
        result = Screen.WaitForKbdUnlock()
        
        ld = Trim(Session.Screen.getstring(1, 8, 6))
        
        If Trim(Session.Screen.getstring(30, 33, 1)) = "" Then ' if ld is pu then
            Screen.SendKeys "<PF10>"
            result = Screen.WaitForKbdUnlock()
            
            moA = Trim(Session.Screen.getstring(11, 14, 1))
            moB = Trim(Session.Screen.getstring(11, 15, 1))
            daA = Trim(Session.Screen.getstring(11, 16, 1))
            daB = Trim(Session.Screen.getstring(11, 17, 1))
            yeA = Trim(Session.Screen.getstring(11, 18, 1))
            yeB = Trim(Session.Screen.getstring(11, 19, 1))
            tiXA = Trim(Session.Screen.getstring(11, 37, 1))
            tiXB = Trim(Session.Screen.getstring(11, 38, 1))
            tiYA = Trim(Session.Screen.getstring(11, 39, 1))
            tiYB = Trim(Session.Screen.getstring(11, 40, 1))
            
            Screen.SendKeys "<PF12>"
            result = Screen.WaitForKbdUnlock()
            
        Else ' if ld is del then
            moA = Trim(Session.Screen.getstring(11, 49, 1))
            moB = Trim(Session.Screen.getstring(11, 50, 1))
            daA = Trim(Session.Screen.getstring(11, 51, 1))
            daB = Trim(Session.Screen.getstring(11, 52, 1))
            yeA = Trim(Session.Screen.getstring(11, 53, 1))
            yeB = Trim(Session.Screen.getstring(11, 54, 1))
            tiXA = Trim(Session.Screen.getstring(11, 61, 1))
            tiXB = Trim(Session.Screen.getstring(11, 62, 1))
            tiYA = Trim(Session.Screen.getstring(11, 63, 1))
            tiYB = Trim(Session.Screen.getstring(11, 64, 1))
            
        End If
        
        Screen.SendKeys "<PF12>"
        result = Screen.WaitForKbdUnlock()
        
        'check first #notify
        moApt = (Asc(moA) - 48) * 10 + (Asc(moB) - 48)
        daApt = (Asc(daA) - 48) * 10 + (Asc(daB) - 48)
        yeApt = 2000 + (Asc(yeA) - 48) * 10 + (Asc(yeB) - 48)
        dtApt = DateSerial(yeApt, moApt, daApt)
        dtNow = Date
        dtDif = DateDiff("d", dtNow, dtApt)
        
        tiApt = (((Asc(tiXA) - 48) * 10 + (Asc(tiXB) - 48)) + ((Asc(tiYA) - 48) * 10 + Asc(tiYB) - 48) / 60) / 24
        tiNow = Time
        tiDif = tiApt - tiNow + dtDif
        
        If tiDif < (2 / 24) And tiApt < (23 + (54 / 60)) And dtDif >= 0 Then
            notify = "Check notify on " & ld
        
        Else
            notify = ""
            
        End If
       
        For a = 10 To 28
            If Trim(Session.Screen.getstring(a, 49, 1)) = "" Then
                Exit For
                
            End If
                
            result = Screen.moveTo(a, 7)
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys "<PF4>"
            result = Screen.WaitForKbdUnlock()
            
            If Trim(Session.Screen.getstring(1, 26, 1)) = "T" Then
                Screen.SendKeys "<PF12>"
                result = Screen.WaitForKbdUnlock()
                GoTo Skip
            
            End If
                
            ld = Trim(Session.Screen.getstring(1, 8, 6))
        
            If Trim(Session.Screen.getstring(30, 33, 1)) = "" Then ' if ld is pu then
                Screen.SendKeys "<PF10>"
                result = Screen.WaitForKbdUnlock()
            
                moA = Trim(Session.Screen.getstring(11, 14, 1))
                moB = Trim(Session.Screen.getstring(11, 15, 1))
                daA = Trim(Session.Screen.getstring(11, 16, 1))
                daB = Trim(Session.Screen.getstring(11, 17, 1))
                yeA = Trim(Session.Screen.getstring(11, 18, 1))
                yeB = Trim(Session.Screen.getstring(11, 19, 1))
                tiXA = Trim(Session.Screen.getstring(11, 37, 1))
                tiXB = Trim(Session.Screen.getstring(11, 38, 1))
                tiYA = Trim(Session.Screen.getstring(11, 39, 1))
                tiYB = Trim(Session.Screen.getstring(11, 40, 1))
            
                Screen.SendKeys "<PF12>"
                result = Screen.WaitForKbdUnlock()
                
            Else ' if ld is del then
                moA = Trim(Session.Screen.getstring(11, 49, 1))
                moB = Trim(Session.Screen.getstring(11, 50, 1))
                daA = Trim(Session.Screen.getstring(11, 51, 1))
                daB = Trim(Session.Screen.getstring(11, 52, 1))
                yeA = Trim(Session.Screen.getstring(11, 53, 1))
                yeB = Trim(Session.Screen.getstring(11, 54, 1))
                tiXA = Trim(Session.Screen.getstring(11, 61, 1))
                tiXB = Trim(Session.Screen.getstring(11, 62, 1))
                tiYA = Trim(Session.Screen.getstring(11, 63, 1))
                tiYB = Trim(Session.Screen.getstring(11, 64, 1))
            
            End If
        
            Screen.SendKeys "<PF12>"
            result = Screen.WaitForKbdUnlock()
        
            'check following #notify
            moApt = (Asc(moA) - 48) * 10 + (Asc(moB) - 48)
            daApt = (Asc(daA) - 48) * 10 + (Asc(daB) - 48)
            yeApt = 2000 + (Asc(yeA) - 48) * 10 + (Asc(yeB) - 48)
            dtApt = DateSerial(yeApt, moApt, daApt)
            dtNow = Date
            dtDif = DateDiff("d", dtNow, dtApt)
        
            tiApt = dtDif + (((Asc(tiXA) - 48) * 10 + (Asc(tiXB) - 48)) + ((Asc(tiYA) - 48) * 10 + (Asc(tiYB) - 48)) / 60) / 24
            tiNow = Time
            tiDif = tiApt - tiNow
        
            If tiDif < (2 / 24) And tiApt < (23 + 54 / 60) And dtDif >= 0 Then
                If notify = "" Then
                    notify = "Check notify on " & ld
                    
                Else
                    notify = notify & ", and " & ld
                        
                End If
            
            End If
Skip:

        Next a
        
EnderX:
        
        If notify = "" Then
            notify = "All clear!"
            
        End If
        
        ActiveCell.Offset(0, 3).Value = notify
        ActiveCell.Offset(1, 0).Activate
        
        notify = ""
        
    Loop
    
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()
    Range("A2").Select
    
    Call LudicrousMode(False)
    
End Sub

Sub Add_Acomm_Click()

    'subs

End Sub

Sub EoD()

Exit Sub
    
    Call Clear_btn_Click
    Call LudicrousMode(True)
    
    InitializeSession

    'check on 198 screen
    If Trim(Session.Screen.getstring(2, 33, 6)) <> "USERID" Or Trim(Session.Screen.getstring(1, 28, 25)) <> "J.B. HUNT TRANSPORT, INC." Then
        MsgBox "Go To The Main Screen And Try Again"
        Call LudicrousMode(False)
        Exit Sub
        
    End If

    result = Screen.moveTo(30, 49)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "198"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "L3A"
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    result = Screen.moveTo(19, 9)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF4>"
    result = Screen.WaitForKbdUnlock()


    Dim a As Integer
    Dim st As String
    Dim stO As String
    Dim act As String
    Dim alp As String
    Dim ntf As Boolean
    
    Range("A2").Select
    
    Do
        For a = 19 To 30
            st = ""
            stO = ""
            act = ""
            trk = ""
            alp = ""

            If Trim(Session.Screen.getstring(a, 9, 1)) = "" Then
                GoTo Ender2
            
            Else
                alp = Trim(Session.Screen.getstring(a, 9, 6))
                stO = Trim(Session.Screen.getstring(a, 38, 2))
                
                result = Screen.moveTo(a, 9)
                result = Screen.WaitForKbdUnlock()
                Screen.SendKeys "<PF10>"
                result = Screen.WaitForKbdUnlock()
                
                If Trim(Session.Screen.getstring(1, 29, 15)) = "MESSAGE DISPLAY" Then
                    Screen.SendKeys "<PF3>"
                    result = Screen.WaitForKbdUnlock()
                
                End If
                
                'st checker
                If Trim(Session.Screen.getstring(30, 49, 10)) = "MULTIPPLAN" Or (Trim(Session.Screen.getstring(30, 63, 1)) = "P" And _
                    Trim(Session.Screen.getstring(6, 62, 1)) <> "") Then
                    st = "OK"
                        
                ElseIf Trim(Session.Screen.getstring(6, 62, 1)) = "" Then
                    st = "HM"
                            
                ElseIf Trim(Session.Screen.getstring(29, 31, 1)) = "L" Then
                    st = "PU"
                        
                ElseIf Trim(Session.Screen.getstring(6, 62, 5)) = Trim(Session.Screen.getstring(7, 7, 5)) Or _
                        Trim(Session.Screen.getstring(6, 64, 1)) = "" Then
                    st = "."
                            
                Else
                    st = "PU"
                        
                End If
                
            End If
                
            result = Screen.moveTo(8, 22)
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys st
            result = Screen.WaitForKbdUnlock()
                
                If st = "." Then
                    Screen.SendKeys "<DELETE>"
                    result = Screen.WaitForKbdUnlock()
                        
                End If
                    
            Screen.SendKeys "<ENTER>"
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys "<PF12>"
            result = Screen.WaitForKbdUnlock()
            
Ender:

        ActiveCell.Value = trk
        ActiveCell.Offset(0, 1).Value = alp
        ActiveCell.Offset(0, 2).Value = act
        ActiveCell.Offset(1, 0).Activate

Ender2:
        
        If Trim(Session.Screen.getstring(31, 8, 1)) = "E" And a = 30 Then
            Exit Do
        End If
        
        Next a
        
        Screen.SendKeys "<PF8>"
        result = Screen.WaitForKbdUnlock()

    Loop Until Trim(Session.Screen.getstring(31, 1, 31)) = "TEQUIP    DUPLICATE RECORD  KEY" And _
        Trim(Session.Screen.getstring(19, 9, 6)) = "______"
                
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()
    Range("A2").Select
    
    Call LudicrousMode(False)
    Exit Sub

    'Email Section
    Dim Email As Outlook.Application
    Set Email = New Outlook.Application
    Dim alp As String
    Dim st As String
    Dim NewMail As Outlook.MailItem
    Set NewMail = Email.CreateItem(olMailItem)

    NewMail.To = "redacted for privacy"
    NewMail.cc = ""
    NewMail.subject = "redacted for privacy"
    
    NewMail.HTMLBody = "Hi," & vbNewLine & "This is a test email from Excel" & _
    vbNewLine & vbNewLine & _
    "Regards," & vbNewLine & _
    "VBA Coder"
    
    'newmail.send | not yet my young one

    Call LudicrousMode(False)
    
End Sub

Sub EnderEmail()

    Range("E3:G21").Value = Range("A3:C21").Value

End Sub

Sub MissingBOL()
Exit Sub

    Call LudicrousMode(True)
    
    InitializeSession

    'check main screen
    If Trim(Session.Screen.getstring(2, 33, 6)) <> "USERID" Or Trim(Session.Screen.getstring(1, 28, 25)) <> "redacted for privacy" Then
        MsgBox "Go To The Main Screen And Try Again"
        Call LudicrousMode(False)
        Exit Sub
        
    Else
        'do it
        result = Screen.moveTo(30, 49)
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "405"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
        
        Dim arr(1, 1) As Variant
        Dim alpha As String '2,0
        Dim quant As Integer 'counted
        Dim load As String  '
        Dim board As String '
        Dim cust As String  '
        
        
        Range("C1").Select
        Do
            'offsets
        Loop Until ActiveCell.Value <> "redacted for privacy"
    
    End If
    
    Call LudicrousMode(False)
    Call Clear_btn_Click
    
End Sub

Sub Change_Hazmat()

    Dim ordNum As String
    Call Copy_Text_From_Clipboard(ordNum)
    
    If Len(ordNum) <> "7" Then
        MsgBox "Invalid Order Number!" & vbNewLine & vbNewLine & ordNum & " is invalid!"
        Exit Sub
    
    End If
    
    Dim Response As Integer
    Response = MsgBox("Add or Remove hazmat from load '" & ordNum & "'?" & vbNewLine & vbNewLine & "Yes = Add, No = Remove, Cancel = Cancel", vbQuestion + vbYesNoCancel, "Add/Remove Hazmat?")
    
    Dim toggle As String
    
    If Response = vbYes Then
        toggle = "Y"
    
    ElseIf Response = vbNo Then
        toggle = "N"
        
    Else
        MsgBox " Did not change hazmat on load '" & ordNum & "'."
        Exit Sub
        
    End If
        InitializeSession
        Dim good As Boolean
        good = True
        Call CheckMainScreen(good)
    
        If good = False Then
            Exit Sub
    
        End If
    
        result = Screen.moveTo(30, 49)
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "1"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys ordNum
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<PF10>"
        result = Screen.WaitForKbdUnlock()
        result = Screen.moveTo(19, 52)
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys toggle
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
    
        If Trim(Session.Screen.getstring(30, 2, 1)) = "P" Then
            Screen.SendKeys "<ENTER>"
            result = Screen.WaitForKbdUnlock()
            MsgBox "Hazmat removed from load " & ordNum
        
        ElseIf Trim(Session.Screen.getstring(30, 4, 1)) <> "U" Then
            MsgBox "Failed to remove hazmat from load " & ordNum
        
        Else
            MsgBox "Hazmat removed from load " & ordNum
        
        End If
    
        Screen.SendKeys "<PF12>"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<PF12>"
        result = Screen.WaitForKbdUnlock()
    
End Sub

Sub Change_Trl_Type()
    
    Dim ordNum As String
    Call Copy_Text_From_Clipboard(ordNum)
    
    If Len(ordNum) <> "7" Then
        MsgBox "Invalid Order Number!" & vbNewLine & vbNewLine & ordNum & " is invalid!"
        Exit Sub
    
    End If
    
    InitializeSession
    
    Dim good As Boolean
    good = True
    Call CheckMainScreen(good)
    
    If good = False Then
        Exit Sub
    
    End If

    result = Screen.moveTo(30, 49)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "1"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys ordNum
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    result = Screen.moveTo(14, 61)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "53"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()
    
    MsgBox "Trailer type changed to '53' for load '" & ordNum & "'."

End Sub

Function Copy_Text_From_Clipboard(Ace As String)

    Dim strPaste As String
    Dim DataObj As MSForms.DataObject
    Set DataObj = New MSForms.DataObject
    DataObj.GetFromClipboard
    strPaste = DataObj.GetText(1)
    Ace = strPaste

End Function

Sub CheckClipboard()

    Dim Val As String
    Call Copy_Text_From_Clipboard(Val)
    
    If Val = "" Then
        MsgBox "Clipboard is empty!"
    
    Else
        MsgBox "Clipboard contains value: '" & Val & "'"
    
    End If

End Sub

Sub CheckMTBilling()

    Dim MTNum As String
    Call Copy_Text_From_Clipboard(MTNum)
    
    If Len(MTNum) <> "6" Then
        MsgBox "Invalid MT Number!" & vbNewLine & vbNewLine & "'" & MTNum & "' is invalid!"
        Exit Sub
        
    End If
    
    Dim obApp As Object
    Dim NewMail As MailItem
    Dim Rail As String

    Rail = InputBox("Please confirm rail code", "Examples: redacted for privacy")
    Rail = UCase(Rail)
    
    If Len(Rail) <> "2" Then
        MsgBox "Invalid rail code!" & vbNewLine & vbNewLine & "'" & Rail & "' is invalid!"
        
    End If

    Set obApp = Outlook.Application
    Set NewMail = obApp.CreateItem(olMailItem)
 
    'You can change the concrete info as per your needs
    With NewMail
         .subject = "Please confirm billing for mt " & MTNum & " to " & Rail & ", thank you!"
         .To = 'redacted for privacy
         .cc = 'redacted for privacy
         .display
         
    End With
    
    Dim Response As String
    Dim xRail As String
    
    If Rail = "1" Then
        xRail = 'redacted for privacy
    
    ElseIf Rail = "2" Then
        xRail = 'redacted for privacy
    
    Else
        Exit Sub
    
    End If
    
    Response = MTNum & " billed for " & xRail & ", thank you!"
    Copy_Text_To_Clipboard (Response)
    
End Sub

Function Copy_Text_To_Clipboard(Ace As String)
    
    Dim obj As New DataObject
    obj.SetText Ace
    obj.PutInClipboard

End Function

Sub GetCustomerCode()
    Dim ordNum As String
    Call Copy_Text_From_Clipboard(ordNum)
    
    If Len(ordNum) <> "7" Then
        MsgBox "Invalid order Number!" & vbNewLine & ordNum & "'" & MTNum & "' is invalid!"
        Exit Sub
        
    End If
    
    InitializeSession
    
    Dim good As Boolean
    good = True
    Call CheckMainScreen(good)
    
    If good = False Then
        Exit Sub
    
    End If
    
    Dim CustCode As String
    
    result = Screen.moveTo(30, 49)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "1"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys ordNum
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    
    If Trim(Session.Screen.getstring(8, 45, 2)) <> "TX" And Trim(Session.Screen.getstring(8, 71, 2)) <> "TX" Then
        MsgBox "Load ain't from Texas!"
        
    ElseIf Trim(Session.Screen.getstring(8, 45, 2)) = "TX" Then
        If Trim(Session.Screen.getstring(8, 71, 2)) = "TX" Then
            MsgBox "Load is point to point!"
            
        Else
            CustCode = Trim(Session.Screen.getstring(3, 33, 6))
            Copy_Text_To_Clipboard (CustCode)
            MsgBox "Customer code: '" & CustCode & "' copied to clipboard, ready for paste"
        
        End If
    
    Else
        CustCode = Trim(Session.Screen.getstring(3, 59, 6))
        Copy_Text_To_Clipboard (CustCode)
        MsgBox "Customer code: '" & CustCode & "' copied to clipboard, ready for paste"
    
    End If
    
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()
    
End Sub

Sub Change_PPNote()

    Dim ordNum As String
    Call Copy_Text_From_Clipboard(ordNum)
    
    If Len(ordNum) <> "7" Then
        MsgBox "Invalid Order Number!" & vbNewLine & vbNewLine & ordNum & " is invalid!"
        Exit Sub
    
    End If
    
    Dim Response As Integer
    Response = MsgBox("Change preplan note to '/ACOMM' OR '/RESET' for load '" & ordNum & "'?" & vbNewLine & vbNewLine & "Yes = /ACOMM, No = /RESET, Cancel = Cancel", vbQuestion + vbYesNoCancel, "ACOMM/RESET?")
    
    Dim toggle As String
    
    If Response = vbYes Then
        toggle = "/ACOMM"
    
    ElseIf Response = vbNo Then
        toggle = "/RESET"
        
    Else
        MsgBox " Did not change preplan note on load '" & ordNum & "'."
        Exit Sub
        
    End If
        InitializeSession
        Dim good As Boolean
        good = True
        Call CheckMainScreen(good)
    
        If good = False Then
            Exit Sub
    
        End If
    
        result = Screen.moveTo(30, 49)
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "1"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys ordNum
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<PF10>"
        result = Screen.WaitForKbdUnlock()
        result = Screen.moveTo(12, 75)
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys toggle
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
        
        If Trim(Session.Screen.getstring(30, 2, 1)) = "P" Then
            result = Screen.moveTo(10, 51)
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys "X"
            result = Screen.WaitForKbdUnlock()
            result = Screen.moveTo(12, 75)
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys toggle
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys "<ENTER>"
            
        End If
        
        Screen.SendKeys "<PF12>"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<PF12>"
        result = Screen.WaitForKbdUnlock()
    
End Sub

Sub Change_PPNote2()

    Dim ordNum As String
    Call Copy_Text_From_Clipboard(ordNum)
    
    If Len(ordNum) <> "7" Then
        MsgBox "Invalid Order Number!" & vbNewLine & vbNewLine & ordNum & " is invalid!"
        Exit Sub
    
    End If
    
    Dim Response As Integer
    Response = MsgBox("Change preplan note to '/TRAIN' OR '/NTRDY' for load '" & ordNum & "'?" & vbNewLine & vbNewLine & "Yes = /TRAIN, No = /NTRDY, Cancel = Cancel", vbQuestion + vbYesNoCancel, "TRAIN/NTRDY?")
    
    Dim toggle As String
    
    If Response = vbYes Then
        toggle = "/TRAIN"
    
    ElseIf Response = vbNo Then
        toggle = "/NTRDY"
        
    Else
        MsgBox " Did not change preplan note on load '" & ordNum & "'."
        Exit Sub
        
    End If
    
        InitializeSession
        Dim good As Boolean
        good = True
        Call CheckMainScreen(good)
    
        If good = False Then
            Exit Sub
    
        End If
    
        result = Screen.moveTo(30, 49)
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "1"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys ordNum
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<PF10>"
        result = Screen.WaitForKbdUnlock()
        result = Screen.moveTo(12, 75)
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys toggle
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        
        If Trim(Session.Screen.getstring(30, 2, 1)) = "P" Then
            result = Screen.moveTo(10, 51)
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys "X"
            result = Screen.WaitForKbdUnlock()
            result = Screen.moveTo(12, 75)
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys toggle
            result = Screen.WaitForKbdUnlock()
            Screen.SendKeys "<ENTER>"
            
        End If
        
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<PF12>"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<PF12>"
        result = Screen.WaitForKbdUnlock()
    
End Sub

Sub TenFourTFTHU()

    Call Copy_Text_To_Clipboard("10-4, thanks for the heads up!")

End Sub

Sub TenFourKUPP()

    Call Copy_Text_To_Clipboard("10-4, keep us posted please!")

End Sub

Sub TenFourTYSM()

    Dim Gender As String
    
    Gender = MsgBox("Sir or Ma'am?" & vbNewLine & vbNewLine & "Yes = Sir, No = Ma'am, Cancel = Cancel", vbQuestion + vbYesNoCancel, "Sir/Ma'am?")
    
    If Gender = vbYes Then
        Call Copy_Text_To_Clipboard("10-4, thank you sir!")
    
    ElseIf Gender = vbNo Then
        Call Copy_Text_To_Clipboard("10-4, thank you ma'am!")
        
    Else
        MsgBox "Did not copy text to clipboard."
    
    End If

End Sub

Sub Did_You_Grab_An_MT()

    Dim truck As String
    
    Call Copy_Text_From_Clipboard(truck)
    InitializeSession
    
    If Len(truck) <> "6" And Left(truck, 1) <> "3" Then
        MsgBox "Invalid truck number!" & vbNewLine & vbNewLine & "'" & truck & "' is invalid!"
        Exit Sub
    
    End If
    
    Dim good As Boolean
    good = True
    Call CheckMainScreen(good)
    
    If good = False Then
        Exit Sub
    
    End If
    
    result = Screen.moveTo(30, 49)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "15"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys truck
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF3>"
    result = Screen.WaitForKbdUnlock()
    result = Screen.moveTo(8, 22)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "??"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF3>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF3>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "Did you grab an MT?"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF6>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()

End Sub

Sub LoadUnload()
    'work in progress////////////////////////////////////////////////////////////////
    
    Dim truck As String
    Dim toggle As String
    Dim init As Integer
    Dim name As String
    
    Call Copy_Text_From_Clipboard(truck)
    InitializeSession
    
    If Len(truck) <> "6" Then
        MsgBox "Invalid truck number!" & vbNewLine & vbNewLine & "'" & truck & "' is invalid!"
        Exit Sub
    
    End If
    
    Dim Response As Integer
    Response = MsgBox("Change preplan not to '/ACOMM' OR '/RESET' for load '" & ordNum & "'?" & vbNewLine & vbNewLine & "Yes = /ACOMM, No = /RESET, Cancel = Cancel", vbQuestion + vbYesNoCancel, "ACOMM/RESET?")
    
    If Response = vbYes Then
        toggle = "LOADING"
    
    ElseIf Response = vbNo Then
        toggle = "UNLOADING"
        
    Else
        MsgBox "No message will be sent to truck '" & truck & "'."
        Exit Sub
        
    End If
    
        Dim good As Boolean
        good = True
        Call CheckMainScreen(good)
    
        If good = False Then
            Exit Sub
    
        End If
    
        result = Screen.moveTo(30, 49)
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "15"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys truck
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
        
        init = 15

        Do While init < 49
            If Trim(Session.Screen.getstring(4, init, 1)) <> "," Then
                MsgBox "Not a ','. init = '" & init & "'."
                Exit Sub
        
            End If
        
        Loop
        
        init = init + 2
        name = Trim(Session.String.getstring(2, init, 49 - init))
        
        MsgBox "Hey, " & name
        
        Screen.SendKeys "<PF5>"
        result = Screen.WaitForKbdUnlock()
        
        Exit Sub
        
        result = Screen.moveTo(12, 75)
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys toggle
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<ENTER>"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<PF12>"
        result = Screen.WaitForKbdUnlock()
        Screen.SendKeys "<PF12>"
        result = Screen.WaitForKbdUnlock()

End Sub

Sub KCSFax()
    
    Dim ordNum As String
    
    Call Copy_Text_From_Clipboard(ordNum)
    
    If Len(ordNum) <> "7" Then
        MsgBox "Invalid Order Number!" & vbNewLine & vbNewLine & ordNum & " is invalid!"
        Exit Sub
    
    End If
    
    ordNum = "BOL for load " & ordNum
    
    Call PrinterSel
    
    Dim objOutApp As Object, objOutMail As Object
    Dim strBody As String, strSig As String
 
    Set objOutApp = CreateObject("Outlook.Application")
    Set objOutMail = objOutApp.CreateItem(0)
 
    On Error Resume Next
    With objOutMail
        .To = 'email redacted for privacy
        .cc = 'email redacted for privacy
        .subject = ordNum
        .display

    End With
 
    On Error GoTo 0
    Set objOutMail = Nothing
    Set objOutApp = Nothing
    
    Call Copy_Text_To_Clipboard(ordNum)
    
End Sub

Sub Whos_That_Chassis_Number()

    Dim truck As String
    Dim trl As String
    
    Call Copy_Text_From_Clipboard(truck)
    InitializeSession
    
    If Len(truck) <> "6" Or Left(truck, 1) <> "3" Then
        MsgBox "Invalid truck number!" & vbNewLine & vbNewLine & "'" & truck & "' is invalid!"
        Exit Sub
    
    End If
    
    Dim good As Boolean
    good = True
    Call CheckMainScreen(good)
    
    If good = False Then
        Exit Sub
    
    End If
    
    result = Screen.moveTo(30, 49)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "15"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys truck
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF3>"
    result = Screen.WaitForKbdUnlock()
    
    If Trim(Session.Screen.getstring(2, 27, 1)) = "" Then
        MsgBox "No trailer number assigned to truck, exiting"
        Exit Sub
        
    End If
    
    trl = Trim(Session.Screen.getstring(2, 27, 6))
    
    result = Screen.moveTo(8, 22)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "??"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF3>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF3>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "Hey, what's the chassis number for " & trl & "? Trying to process your calls."
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF6>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()

End Sub

Sub Whos_That_State_Code()

    Dim truck As String
    
    Call Copy_Text_From_Clipboard(truck)
    InitializeSession
    
    If Len(truck) <> "6" And Left(truck, 1) <> "3" Then
        MsgBox "Invalid truck number!" & vbNewLine & vbNewLine & "'" & truck & "' is invalid!"
        Exit Sub
    
    End If
    
    Dim good As Boolean
    good = True
    Call CheckMainScreen(good)
    
    If good = False Then
        Exit Sub
    
    End If
    
    result = Screen.moveTo(30, 49)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "15"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys truck
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF3>"
    result = Screen.WaitForKbdUnlock()
    result = Screen.moveTo(8, 22)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "??"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF3>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF3>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "What's the final destination for this load? Getting an error on this end"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF6>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()

End Sub

Sub EoD_Fill()

    Dim Res As Integer
    
    Res = InputBox("1 = Still Running" & vbCrLf & "2 = Run, End" & vbCrLf & "3 = In, End" & vbCrLf & "4 = MT, End" & vbCrLf & "5 = Del since" & vbCrLf & "6 = Pick since", "Comment >> Result")
    
    Select Case Res
    
    Case 1
        ActiveCell.Value = "Still running"
    Case 2
        ActiveCell.Value = "Running last load, end of day after"
    Case 3
        ActiveCell.Value = "Ingating last load, end of day after"
    Case 4
        ActiveCell.Value = "Ingating MT, end of day after"
    Case 5
        ActiveCell.Value = ActiveCell.Value + ", unloading since"
    Case 6
        ActiveCell.Value = ActiveCell.Value + ", loading since"
    Case Else
        MsgBox "Bad input"
        Exit Sub
    End Select

End Sub

Sub HOD_EoD()

    ' SET Outlook APPLICATION OBJECT.
    Dim num As Integer
    Dim rnga As String
    Dim rngb As String
    Dim coverage As String
    Dim cc As String
    Dim subj As String
    
    If ActiveSheet.name = "1" Then
        coverage = 'redacted for privacy
        cc = 'redacted for privacy
        subj = "PASSDOWN 1"
        
        num = 3
        rnga = "B" & num
    
        Do While IsEmpty(Range(rnga)) = False
            num = num + 1
            rnga = "B" & num
    
        Loop
    
        rngb = "A1:B" & num - 1

        Range(rngb).Copy

    ElseIf ActiveSheet.name = "2" Then
        coverage = 'redacted for privacy
        cc = 'redacted for privacy
        subj = "PASSDOWN 2"
        
        num = 3
        rnga = "B" & num
    
        Do While IsEmpty(Range(rnga)) = False
            num = num + 1
            rnga = "B" & num
    
        Loop
    
        rngb = "A1:B" & num - 1

        Range(rngb).Copy
        
    ElseIf ActiveSheet.name = "3" Then
        coverage = 'redacted for privacy
        cc = 'redacted for privacy
        subj = "PASSDOWN 3"
        
        num = 3
        rnga = "B" & num
    
        Do While IsEmpty(Range(rnga)) = False
            num = num + 1
            rnga = "B" & num
    
        Loop
    
        rngb = "A1:B" & num - 1

        Range(rngb).Copy
        
    ElseIf ActiveSheet.name = "Regional" Then
        coverage = 'redacted for privacy
        cc = 'redacted for privacy
        subj = "PASSDOWN 4"
        
        num = 3
        rnga = "B" & num
    
        Do While IsEmpty(Range(rnga)) = False
            num = num + 1
            rnga = "B" & num
    
        Loop
    
        rngb = "A1:B" & num - 1

        Range(rngb).Copy
    
    End If
    
    num = 3
    rnga = "B" & num
    
    Do While IsEmpty(Range(rnga)) = False
        num = num + 1
        rnga = "B" & num
    
    Loop
    
    rngb = "A1:B" & num - 1

    Range(rngb).Copy
    
    Dim objOutlook As Object
    Set objOutlook = Outlook.Application

    ' CREATE EMAIL OBJECT.
    Dim objEmail As MailItem
    Set objEmail = objOutlook.CreateItem(olMailItem)

    With objEmail
        .To = coverage
        .cc = cc
        .subject = subj
        .display    ' DISPLAY MESSAGE.
        
    End With
    
    If ActiveSheet.name = "1" Then
        MsgBox "Waiting paste 1"
        
        num = 3
        rnga = "D" & num
    
        Do While IsEmpty(Range(rnga)) = False
            num = num + 1
            rnga = "D" & num
    
        Loop
    
        rngb = "C1:D" & num - 1

        Range(rngb).Copy
        
        MsgBox "Waiting paste 2"
        
        num = 3
        rnga = "F" & num
    
        Do While IsEmpty(Range(rnga)) = False
            num = num + 1
            rnga = "F" & num
    
        Loop
    
        rngb = "E1:F" & num - 1

        Range(rngb).Copy
        
        MsgBox "Waiting paste 3"
        
        num = 3
        rnga = "H" & num
    
        Do While IsEmpty(Range(rnga)) = False
            num = num + 1
            rnga = "H" & num
    
        Loop
    
        rngb = "G1:H" & num - 1

        Range(rngb).Copy
        
        MsgBox "Waiting paste 4"
        
    End If

End Sub

Sub Hash_Notify()

    Dim ordNum As String
    Call Copy_Text_From_Clipboard(ordNum)
    
    If Len(ordNum) <> "7" Then
        MsgBox "Invalid Order Number!" & vbNewLine & vbNewLine & ordNum & " is invalid!"
        Exit Sub
    
    End If
    
    InitializeSession
    
    Dim good As Boolean
    good = True
    Call CheckMainScreen(good)
    
    If good = False Then
        Exit Sub
    
    End If

    result = Screen.moveTo(30, 49)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "1"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys ordNum
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    
    Do While Trim(Session.Screen.getstring(22, 17, 1)) <> ""
        Screen.SendKeys "<PF8>"
        result = Screen.WaitForKbdUnlock()
    
    Loop
    
    result = Screen.moveTo(22, 6)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "a"
    result = Screen.WaitForKbdUnlock()
    result = Screen.moveTo(22, 17)
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "#notify we will be late"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<ENTER>"
    result = Screen.WaitForKbdUnlock()
    Screen.SendKeys "<PF12>"
    result = Screen.WaitForKbdUnlock()

End Sub
