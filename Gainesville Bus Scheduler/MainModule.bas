Attribute VB_Name = "MainModule"
Type bus
    location As String
    code As Integer
End Type
Const offsetCycle = 44
Const offsetDay = 9
Const offsetBus = 1
Const offset = 6
'constants are the constants to offset for a new day or bus.

Public bus(1 To 6) As bus
Public prevCode(1 To 6) As Variant

Public currentCycle As Integer
Public currentDay As Integer
Sub refresh()
'
' refresh Macro for Route Information refresh button
    activeworkbook.RefreshAll
End Sub

Sub update(ByVal Target As Range)


    If Not Intersect(Target, Range("updateTime")) Is Nothing Then
    'if statement determines if the update time has changed.
        Dim time As String
        Dim add As Boolean
        Dim activewb As String
        
        
        activewb = activeworkbook.Name
        activews = ActiveSheet.Name
        Workbooks("p4_vannuge.xlsm").Activate
        Worksheets("Data").Activate
        'this will allow me to work on other excel books but still update this book in the background.
        
    
            Dim busnum As Integer
            Dim limit As Integer
            
        For i = 1 To UBound(bus())
            prevCode(i) = Worksheets("Data").Range("lastB" & i).Value
            Worksheets("Data").Range("lastB" & i).Value = 0
        Next
         
    
    'I'm recording all the last stop codes and then setting values to 0 to clear them.
         
        For Each C In Target.Cells
    
            'for each cell in the all of the worksheet.
                If InStr(1, C.Value, "[") <> 0 And InStr(1, C.Value, "bus number 1 will arrive") = 0 Then
                    Debug.Print C.Value
                    'will detect if a bus is is at a stop with using "["
    
                    
                    If InStr(1, C.Value, ",") <> 0 Then
                    'this will see if theres a comma in the cell, which indicates there are more than one bus in one place.
                        Dim x As Variant
                        Dim temp As Variant
                        
                        
                        x = Split(Mid(C.Value, 2, InStr(1, C.Value, "]") - 2), ", ")
                    'Used in case more than one bus is in one place, will split up the numbers
                    
                        For Each temp In x
                        'for each bus number
                        
                            bus(CInt(temp)).location = Mid(C.Value, InStr(1, C.Value, "]") + 2)
                            Debug.Print temp
                            Debug.Print bus(CInt(temp)).location
                      
                            bus(CInt(temp)).code = Mid((Worksheets("Route Information").Range(C.Address).offset(1, 0)), 12)
                            Debug.Print "code" & bus(CInt(temp)).code
                        
                            If bus(CInt(temp)).code <> prevCode(temp) Then
                                add = True
                                Worksheets("Data").Range("lastB" & temp).Value = bus(CInt(temp)).code
                            End If
                        Next
                        Call addCollision(x)
                    Else
                        busnum = Mid(C.Value, 2, 1)
                        Debug.Print busnum
                        
                        bus(busnum).location = Mid(C.Value, InStr(1, C.Value, "]") + 2)
                        Debug.Print bus(busnum).location
                        bus(busnum).code = Mid((Worksheets("Route Information").Range(C.Address).offset(1, 0)), 12)
                        
                        Debug.Print bus(busnum).code
                        If bus(busnum).code <> prevCode(busnum) Then
                            add = True
                        End If
                        Worksheets("Data").Range("lastB" & busnum).Value = bus(busnum).code
                    End If
        
                End If
            Next
            
        If add = True Then
        'if add is true, then will add new data to the table
            If Worksheets("Data").Range("currentDate").Value <> Date Then
                Call MainModule.addDay
            End If
                    
            time = Mid(Range("updatetime").Value, 18)
        
            Worksheets("Data").Range("currentTime").Value = time
            Worksheets("Data").Range("currentDate").Value = Date
            
            Call MainModule.addData
        End If
        
        Workbooks(activewb).Activate
        Worksheets(activews).Activate
       Application.ScreenUpdating = True

    End If

End Sub

Public Sub addData()
    Dim currentStart, entry As Range
    Dim count As Integer
    Dim col, col2 As String
    Dim activewb As String
    
    activewb = activeworkbook.Name
    
    Workbooks("p4_vannuge.xlsm").Activate
    Application.ScreenUpdating = False
    Worksheets("Data").Activate
    
    currentDay = Range("daysLogged").Value - 1
    Set currentStart = Range("start").offset(0, offsetDay * currentDay + offset)
        col = Mid(Range(currentStart.Address).offset(0, 1).Address, 2, Len(Range(currentStart.Address).offset(0, 1).Address) - 3)
    'col is to extract the letter of the column with the times
    count = WorksheetFunction.CountA(Range(col & ":" & col))
    'count is to find out where to add a new entry
    Set entry = Range(currentStart.Address).offset(count, 1)
        entry.Value = Format(Range("currentTime").Value, "Long Time")
    
    For i = 1 To UBound(bus())
        entry.offset(0, i).Value = CInt(bus(i).code)
    Next
    
    Call addAverage
    
    
End Sub
    
Public Sub addDay()
    'add a new day of data
    activewb = activeworkbook.Name
    Workbooks("p4_vannuge.xlsm").Activate
    Application.ScreenUpdating = False
    Worksheets("Data").Activate
    currentDay = Worksheets("Data").Range("daysLogged").Value
    
    
    'copy first table as a template
    Range("dayTemplate").Select
    Selection.Copy
    Range("start").offset(-1, offsetDay * currentDay + offset).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("start").offset(0, offsetDay * currentDay + offset).Value = Date
    
    'clear the contents of the copied table
    Dim rngRange As Range
    With Sheets("Data")
      Set rngRange = .Range _
        (.Cells(7, offsetDay * currentDay + offset), .Cells(.Rows.count, .Columns.count))
    End With
    
    rngRange.ClearContents
    
    Dim x As Integer
    'autofit just the new table, won't mess with some pre-existing formatting)
    For x = offsetDay * currentDay To Worksheets("Data").UsedRange.Columns.count
         Columns(x).EntireColumn.AutoFit
    Next x
    
    
    Range("daysLogged").Value = currentDay + 1
End Sub

Sub addAverage()


    Dim cindex, cindex2, minutes As Integer
    
    'For Each c In Worksheets("Data").Range("stopRange")
    'will insert new last stop time data into last stop time table
    For i = 1 To UBound(bus())
        If bus(i).code <> 0 Then
                
            cindex2 = WorksheetFunction.Match(bus(i).code & "*", Worksheets("Data").Range("stopRange"), 0)
           ' If prevCode(i) <> 0 Then
               ' cindex = WorksheetFunction.Match(prevCode(i) & "*", Worksheets("Data").Range("stopRange"), 0)
                
          '  Else
            Dim prevtime, lasttime As Variant
          
            lasttime = Worksheets("Data").Range("currentTime").Value
            'previous stop time and time difference calculation
            If prevCode(i) <> bus(i).code Then
            
                Worksheets("Data").Range("prevStop").Cells(cindex2).Value = Worksheets("Data").Range("lastStop").Cells(cindex2).Value
                prevtime = Worksheets("Data").Range("prevStop").Cells(cindex2).Value
                Range("Data!stopDif").Cells(cindex2).Value = 1440 * Abs(lasttime - prevtime)
                Worksheets("Data").Range("numRec").Cells(cindex2).Value = CInt(Worksheets("Data").Range("numRec").Cells(cindex2).Value) + 1
                
                If Worksheets("Data").Range("numRec").Cells(cindex2).Value <= 1 Then
                    Worksheets("Data").Range("average").Cells(cindex2).Value = Worksheets("Data").Range("stopDif").Cells(cindex2).Value
                Else
                    Worksheets("Data").Range("average").Cells(cindex2).Value = Worksheets("Data").Range("average").Cells(cindex2).Value * CInt((Range("Data!numRec").Cells(cindex2).Value - 1) / (Range("Data!numRec").Cells(cindex2).Value)) + (1 / Range("Data!numRec").Cells(cindex2).Value) * CInt(Range("Data!stopDif").Cells(cindex2).Value)
                
                End If
                Debug.Print Worksheets("Data").Range("average").Cells(cindex2).Value
                Debug.Print CInt(Worksheets("Data").Range("average").Cells(cindex2).Value) * CInt((Range("Data!numRec").Cells(cindex2).Value - 1) / (Range("Data!numRec").Cells(cindex2).Value))
                Debug.Print (1 / Range("Data!numRec").Cells(cindex2).Value) * CInt(Range("Data!stopDif").Cells(cindex2).Value)
                
                
            End If
            
            Worksheets("Data").Range("lastStop").Cells(cindex2).Value = Worksheets("Data").Range("currentTime").Value
        
            'End If
            
            
           ' If (Abs(cindex - cindex2) <= 1) Then
    
            '    Worksheets("Data").Range("stopRange").offset(cindex2, 1).Value = Worksheets("Data").Range("currentTime").Value
          '  Else
            
               ' For j = 0 To Abs(cindex - cindex2)
                'will fill in every bus stop between the previously coded stop for that bus and the current stop.
                 '  Worksheets("Data").Range("stopRange").offset(cindex2 - j, 1).Value = Worksheets("Data").Range("currentTime").Value
               ' Next
           ' End If
            
        End If
    Next
    'Next
End Sub

Sub addCollision(x As Variant)
Dim collisionStart As Range
Dim count As Integer
Dim temp As Variant


Set collisionStart = Worksheets("Real-time + Analysis").Range("collisionStart")
count = WorksheetFunction.CountA(Range("I:I"))
Debug.Print count
collisionStart.offset(count, 0) = Range("Data!currentDate").Text
collisionStart.offset(count, 1) = Range("Data!currentTime").Text

For i = LBound(x) To UBound(x)
    If i = LBound(x) Then
        collisionStart.offset(count, 2) = x(i)
    Else
        collisionStart.offset(count, 2) = collisionStart.offset(count, 2) & "," & x(i)
    End If
Next

collisionStart.offset(count, 3) = bus(x(LBound(x))).location


End Sub
