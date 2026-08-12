Option Explicit

'=====================================================================
' IO ASSIGNMENT GENERATOR
' Parses a Prosafe-RS "IO Parameter Builder" CSV export and renders
' an IO Assignment sheet (all nodes, one combined sheet, v1).
'=====================================================================

Public Const COMM_MODULE As String = "SSB401-E3"        ' ESB Bus communication module
Public Const POWER_MODULE As String = "SPW484-E1"       ' Power supply module
Public Const CONTROLLER_MODULE As String = "S2CP471-11" ' Controller/CPU module (Node 1 only)
Public Const TOTAL_CHANNELS As Long = 16

Public Const CLR_REDUNDANT As Long = 13434879   ' light amber - RGB(255,242,204)
Public Const CLR_HEADER As Long = 14277081      ' light grey  - RGB(217,217,217)

'====================== ENTRY POINT ======================
Sub GenerateIOAssignment()

    Dim csvPath As Variant
    csvPath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , _
        "Select IO Parameter Builder CSV export")
    If csvPath = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "Parsing CSV..."

    Dim dictNodes As Object
    Set dictNodes = ParseCSV(CStr(csvPath))

    Application.StatusBar = "Validating redundancy..."
    Dim issues As Collection
    Set issues = ValidateRedundancy(dictNodes)

    Application.StatusBar = "Rendering IO Assignment sheet..."
    Dim ws As Worksheet
    Set ws = PrepareSheet("IO Assignment")
    RenderIOAssignment ws, dictNodes

    If issues.Count > 0 Then
        ReportIssues issues
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "IO Assignment generated for " & dictNodes.Count & " node(s)." & vbCrLf & _
           issues.Count & " redundancy flag mismatch(es) found" & _
           IIf(issues.Count > 0, " - see 'Redundancy Review' sheet.", "."), vbInformation

End Sub

'====================== SHEET PREP ======================
Private Function PrepareSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 8
    Set PrepareSheet = ws
End Function

'====================== CSV PARSER ======================
' Reads the hierarchical @SHEET,Node / @SHEET,Module / @SHEET,Channel
' block structure into: dictNodes(NodeNo) -> dictSlots(SlotNo) -> {
'   "Device": String, "Redundant": "0"/"1",
'   "Channels": dictChannels(ChannelNo) -> Tag String
' }
Private Function ParseCSV(path As String) As Object
    Dim dictNodes As Object
    Set dictNodes = CreateObject("Scripting.Dictionary")

    Dim fnum As Integer
    fnum = FreeFile
    Open path For Input As #fnum

    Dim ln As String
    Dim section As String: section = ""
    Dim curNode As String, curSlot As String, curDevice As String

    Do While Not EOF(fnum)
        Line Input #fnum, ln
        ln = Trim$(ln)
        If Len(ln) = 0 Then GoTo NextLine

        Dim parts() As String
        parts = Split(ln, ",")

        If Left$(ln, 7) = "@SHEET," Then
            section = Trim$(parts(1))
            GoTo NextLine
        End If

        Select Case section

            Case "Node"
                If parts(0) = "@Node Number" Then
                    Dim nKey As String: nKey = Trim$(parts(1))
                    If Not dictNodes.Exists(nKey) Then
                        dictNodes.Add nKey, CreateObject("Scripting.Dictionary")
                    End If
                End If

            Case "Module"
                If parts(0) = "@Node Number" Then
                    curNode = Trim$(parts(1))
                    If Not dictNodes.Exists(curNode) Then
                        dictNodes.Add curNode, CreateObject("Scripting.Dictionary")
                    End If
                ElseIf parts(0) = "@Slot Number" Then
                    curSlot = Trim$(parts(1))
                ElseIf parts(0) = "@Device" Then
                    curDevice = Trim$(parts(1))
                ElseIf parts(0) = "@Dual-Redundant" Then
                    ' by this point curNode, curSlot, curDevice are all set
                    ' (CSV field order: Node Number, Slot Number, Device, Dual-Redundant, ...)
                    Dim dictSlots As Object
                    Set dictSlots = dictNodes(curNode)
                    Dim slotInfo As Object
                    Set slotInfo = CreateObject("Scripting.Dictionary")
                    slotInfo.Add "Device", curDevice
                    slotInfo.Add "Redundant", Trim$(parts(1))
                    slotInfo.Add "Channels", CreateObject("Scripting.Dictionary")
                    If dictSlots.Exists(curSlot) Then dictSlots.Remove curSlot
                    dictSlots.Add curSlot, slotInfo
                End If

            Case "Channel"
                If parts(0) = "@Channel Number" Then
                    ' header row - column order is fixed by the Prosafe-RS export spec:
                    ' Channel Number, Wiring Position, I/O Variable Name, Direction, Comment, ...
                    ' nothing to do here, just skip
                Else
                    Dim chNum As String
                    chNum = Trim$(parts(0))
                    If Len(chNum) > 0 And IsNumeric(chNum) Then
                        Dim tagName As String: tagName = ""
                        If UBound(parts) >= 2 Then tagName = Trim$(parts(2))
                        Dim chans As Object
                        Set chans = dictNodes(curNode)(curSlot)("Channels")
                        If chans.Exists(chNum) Then chans.Remove chNum
                        chans.Add chNum, tagName
                    End If
                End If

        End Select

NextLine:
    Loop
    Close #fnum

    Set ParseCSV = dictNodes
End Function

'====================== REDUNDANCY VALIDATION ======================
' Cross-checks the Dual-Redundant flag against slot occupancy for
' every odd-numbered slot. Flags disagreements rather than guessing.
Private Function ValidateRedundancy(dictNodes As Object) As Collection
    Dim issues As New Collection
    Dim nodeKey As Variant, slotKey As Variant

    For Each nodeKey In dictNodes.Keys
        Dim slots As Object: Set slots = dictNodes(nodeKey)
        For Each slotKey In slots.Keys
            Dim slotNum As Long: slotNum = CLng(slotKey)
            If slotNum Mod 2 = 1 Then
                Dim pairKey As String: pairKey = CStr(slotNum + 1)
                Dim pairExists As Boolean: pairExists = slots.Exists(pairKey)
                Dim flagStr As String: flagStr = slots(slotKey)("Redundant")
                Dim flagRedundant As Boolean: flagRedundant = (flagStr = "1")

                ' Consistent cases: pair absent & flag=1 (redundant pair)
                '                   pair present & flag=0 (two non-redundant modules)
                ' Anything else is a mismatch worth a human look.
                If (pairExists And flagRedundant) Or (Not pairExists And Not flagRedundant) Then
                    issues.Add "Node " & nodeKey & " Slot " & slotKey & _
                        ": Dual-Redundant flag = " & flagStr & _
                        ", but paired slot " & pairKey & _
                        IIf(pairExists, " IS occupied", " is EMPTY") & _
                        " in the CSV. Please verify against site/CSV."
                End If
            End If
        Next slotKey
    Next nodeKey

    Set ValidateRedundancy = issues
End Function

'====================== RENDER: IO ASSIGNMENT ======================
Private Sub RenderIOAssignment(ws As Worksheet, dictNodes As Object)

    Dim sortedNodes() As Long
    Dim cnt As Long: cnt = dictNodes.Count
    ReDim sortedNodes(cnt - 1)
    Dim i As Long: i = 0
    Dim k As Variant
    For Each k In dictNodes.Keys
        sortedNodes(i) = CLng(k)
        i = i + 1
    Next k

    Dim a As Long, b As Long, tmp As Long
    For a = LBound(sortedNodes) To UBound(sortedNodes) - 1
        For b = a + 1 To UBound(sortedNodes)
            If sortedNodes(b) < sortedNodes(a) Then
                tmp = sortedNodes(a): sortedNodes(a) = sortedNodes(b): sortedNodes(b) = tmp
            End If
        Next b
    Next a

    Dim r As Long: r = 1
    For i = LBound(sortedNodes) To UBound(sortedNodes)
        Dim nodeNum As Long: nodeNum = sortedNodes(i)
        Dim slots As Object: Set slots = dictNodes(CStr(nodeNum))
        r = RenderNodeBlock(ws, r, nodeNum, slots)
        r = r + 2
    Next i

    ws.Columns.AutoFit

End Sub

' Renders one node's block starting at startRow. Returns the next free row.
Private Function RenderNodeBlock(ws As Worksheet, startRow As Long, nodeNum As Long, slots As Object) As Long

    Dim r As Long: r = startRow
    Dim c As Long

    ws.Cells(r, 1).Value = "Node : " & nodeNum
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    ' ---- Build the column plan for this node's template ----
    ' Each entry: Array(ColumnLabel, "csv"/"fixed", SlotNumber-or-FixedModuleName)
    Dim colPlan As Collection
    Set colPlan = New Collection
    Dim s As Long

    If nodeNum = 1 Then
        For s = 1 To 6
            colPlan.Add Array("S" & s, "csv", s)
        Next s
        colPlan.Add Array("S7", "fixed", COMM_MODULE)
        colPlan.Add Array("S8", "fixed", COMM_MODULE)
        colPlan.Add Array("Controller", "fixed", CONTROLLER_MODULE)
        colPlan.Add Array("Controller", "fixed", CONTROLLER_MODULE)
        colPlan.Add Array("Power", "fixed", POWER_MODULE)
        colPlan.Add Array("Power", "fixed", POWER_MODULE)
    Else
        For s = 1 To 8
            colPlan.Add Array("S" & s, "csv", s)
        Next s
        colPlan.Add Array("Comm", "fixed", COMM_MODULE)
        colPlan.Add Array("Comm", "fixed", COMM_MODULE)
        colPlan.Add Array("Power", "fixed", POWER_MODULE)
        colPlan.Add Array("Power", "fixed", POWER_MODULE)
    End If

    ' ---- Slot label header row ----
    ws.Cells(r, 1).Value = "Ch"
    ws.Cells(r, 1).Font.Bold = True
    c = 2
    Dim item As Variant
    For Each item In colPlan
        ws.Cells(r, c).Value = item(0)
        ws.Cells(r, c).Font.Bold = True
        ws.Cells(r, c).Interior.Color = CLR_HEADER
        ws.Cells(r, c + 1).Interior.Color = CLR_HEADER
        c = c + 2
    Next item
    r = r + 1

    ' ---- Module type row ----
    c = 2
    For Each item In colPlan
        Dim kind As String: kind = item(1)
        Dim moduleName As String
        Dim isRedundant As Boolean: isRedundant = False

        If kind = "csv" Then
            Dim slotKey As String: slotKey = CStr(item(2))
            If slots.Exists(slotKey) Then
                moduleName = slots(slotKey)("Device")
                isRedundant = (slots(slotKey)("Redundant") = "1")
            Else
                moduleName = ""
            End If
        Else
            moduleName = item(2)
        End If

        ws.Cells(r, c).Value = moduleName & IIf(isRedundant, " Redundant", "")
        ws.Cells(r, c).Font.Bold = True
        ws.Range(ws.Cells(r, c), ws.Cells(r, c + 1)).Borders.LineStyle = xlContinuous
        If isRedundant Then
            ws.Range(ws.Cells(r, c), ws.Cells(r, c + 1)).Interior.Color = CLR_REDUNDANT
        End If
        c = c + 2
    Next item
    r = r + 1

    ' ---- 16 channel rows ----
    Dim ch As Long
    For ch = 1 To TOTAL_CHANNELS
        ws.Cells(r, 1).Value = ch
        c = 2
        For Each item In colPlan
            Dim kind2 As String: kind2 = item(1)
            If kind2 = "csv" Then
                Dim slotKey2 As String: slotKey2 = CStr(item(2))
                If slots.Exists(slotKey2) Then
                    Dim chans As Object: Set chans = slots(slotKey2)("Channels")
                    Dim chKey As String: chKey = CStr(ch)
                    If chans.Exists(chKey) Then
                        ws.Cells(r, c + 1).Value = chans(chKey)
                    End If
                End If
            End If
            ws.Range(ws.Cells(r, c), ws.Cells(r, c + 1)).Borders.LineStyle = xlContinuous
            c = c + 2
        Next item
        r = r + 1
    Next ch

    RenderNodeBlock = r

End Function

'====================== REDUNDANCY ISSUE REPORT ======================
Private Sub ReportIssues(issues As Collection)
    Dim ws As Worksheet
    Set ws = PrepareSheet("Redundancy Review")
    ws.Cells(1, 1).Value = "Dual-Redundant flag / slot-occupancy mismatches - please verify against the CSV or site data"
    ws.Cells(1, 1).Font.Bold = True

    Dim i As Long
    For i = 1 To issues.Count
        ws.Cells(i + 2, 1).Value = issues(i)
    Next i
    ws.Columns(1).ColumnWidth = 110
    ws.Columns(1).WrapText = False
End Sub


