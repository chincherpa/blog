Title: EXCEL zeugs
Date: 2018-11-30 12:09
Author: Lulef
Category: Sammlung
Slug: excel-zeugs
Status: published
EXCEL Funktionen
<!--more-->
    'Folderpicker
    Function GetFolder(strPath As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Pfad zu den Bildern"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
    NextCode:
    GetFolder = sItem
    Set fldr = Nothing
    End Function
    ----------------------------------------------------------
    http://www.cpearson.com/excel/FOLDERTREEVIEW.ASPX
    --------------------------------------------------------------------------------
    'Mehrere Spalten zu einer
    =INDEX(MyData;1+GANZZAHL((ZEILE(A1)-1)/SPALTEN(MyData));REST(ZEILE(A1)-1+SPALTEN(MyData);SPALTEN(MyData))+1)
    'runterziehen bis Fehler #BEZUG
    ------------------------------------------------
    Sub Dictionary()
    ' Select Tools->References from the Visual Basic menu.
    ' Check box beside "Microsoft Scripting Runtime" in the list.
    Dim sName As String
    Dim dict As New Scripting.Dictionary
    dict.CompareMode = TextCompare 'GROSS/klein egal
    ' Add to fruit to Dictionary
    dict.Add "Lutz", 1
    dict.Add "Hannah", 2
    dict.Add Range("A1").Value, Range("B1").Value
    dict.Add Range("A2").Value, Range("B2").Value
    'sName = InputBox("Name neuer Eintrag")
    'sAlter = InputBox("Wert neuer Eintrag")
    'dict.Add sName, sAlter
    sFrage = InputBox("Name abfragen")
    If dict.Exists(sFrage) Then
    MsgBox sFrage & " existiert mit dem Wert: " & dict(sFrage)
    Else
    MsgBox sFrage & " existiert nicht."
    End If
    Set dict = Nothing
    End Sub
    --------------------------------------------------------------------------------
    Sub JedeNteZeile()
    Dim rng As Range
    Dim InputRng As Range
    Dim OutRng As Range
    Dim xInterval As Integer
    Set InputRng = Application.Selection
    Set InputRng = Application.InputBox("Range :", InputRng.Address, Type:=8)
    xInterval = Application.InputBox("Jede wievielte Zeile selektieren?", Type:=1)
    For i = 1 To InputRng.Rows.Count Step xInterval ' + 1
    Set rng = InputRng.Cells(i, 1)
    If OutRng Is Nothing Then
    Set OutRng = rng
    Else
    Set OutRng = Application.Union(OutRng, rng)
    End If
    Next
    OutRng.EntireRow.Select
    End Sub
    --------------------------------------------------------------------------------
    Sub JedeNteSpalte()
    Dim rng As Range
    Dim InputRng As Range
    Dim OutRng As Range
    Dim xInterval As Integer
    Set InputRng = Application.Selection
    Set InputRng = Application.InputBox("Range :", InputRng.Address, Type:=8)
    xInterval = Application.InputBox("Jede wievielte Spalte selektieren", Type:=1)
    For i = 1 To InputRng.Columns.Count Step xInterval ' + 1
    Set rng = InputRng.Cells(1, i)
    If OutRng Is Nothing Then
    Set OutRng = rng
    Else
    Set OutRng = Application.Union(OutRng, rng)
    End If
    Next
    OutRng.EntireColumn.Select
    End Sub
    --------------------------------------------------------------------------------
    'Prüfen, ob eine Datei existiert
    Public Function MyFileExists(MyFilePath As String) As Boolean
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    MyFileExists = objFSO.FileExists(MyFilePath)
    Set objFSO = Nothing
    End Function
    Sub xusername()
    MsgBox Application.UserName & Chr(13) & Len(Application.UserName)
    MsgBox ", Bertrandt" & Chr(13) & Len(", Bertrandt")
    blaenge = Len(", Bertrandt")
    neuername = Left(Application.UserName, Len(Application.UserName) - blaenge)
    MsgBox neuername & Chr(13) & Len(neuername)
    End Sub
    --------------------------------------------------------------------------------
    benutzer = Application.UserName
    --------------------------------------------------------------------------------
    'Worksheet_SelectionChange
    'In diesem Beispiel wird der Inhalt des Arbeitsmappenfenster verschoben, bis sich die Markierung in der oberen linken Ecke des Fensters befindet.
    Private Sub Worksheet_SelectionChange(ByVal Target As Range) 
    With ActiveWindow 
    .ScrollRow = Target.Row 
    .ScrollColumn = Target.Column 
    End With 
    End Sub
    --------------------------------------------------------------------------------
    Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    On Error Resume Next
    ActiveCell.Value = DateClicked
    Unload Me
    End Sub
    --------------------------------------------------------------------------------
    'Doppelklick in Zelle-Ereignis
    Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, Range("A3:A25, A29:A40, F3:F37, K3:K40, P3:P40")) Is Nothing Then
    UserFormKalender.Show
    End If
    If Target.Column = 9 Or Target.Column = 10 Then
    UserFormKalender.Show
    End If
    End Sub
    ------------------------------
    Private Sub Worksheet_Activate()
    End Sub
    ------------------------------
    Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
    End Sub
    --------------------------------------------------------------------------------
    'letzte Zeile finden
    letztezeile = Sheets("Tabelle1").Cells(Rows.Count, 1).End(xlUp).Row
    --------------------------------------------------------------------------------
    'Anzahl
    Dim iVal As Integer
    iVal = Application.WorksheetFunction.COUNTIF(Range("A1:A10"),"Suchbegriff")
    --------------------------------------------------------------------------------
    'Hyperlinkmenü öffnen
    Application.Dialogs(xlDialogInsertHyperlink).Show
    --------------------------------------------------------------------------------
    'PasteSpecial
    ws.Range("C5:L19").Copy
    wb.Sheets(1).Range("A1").PasteSpecial Paste:=xlPasteValues
    wb.Sheets(1).Range("A1").PasteSpecial Paste:=xlPasteFormats
    --------------------------------------------------------------------------------
    '"Close"-Schaltfläche disablen
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
    Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Private Const SC_CLOSE As Long = &HF060
    Private Sub UserForm_Initialize()
    'Dim hWndForm As Long
    'Dim hMenu As Long
    'hWndForm = FindWindow("ThunderDFrame", Me.Caption) 'XL2000
    'hMenu = GetSystemMenu(hWndForm, 0)
    'DeleteMenu hMenu, SC_CLOSE, 0&
    End Sub
    --------------------------------------------------------------------------------
    'Werte eines Monats addieren
    'Zeile (99) = Anzahl der Werte
    {=SUMME(WENN(MONAT($B$2:$B$99)=MONAT($L2);$E$2:$E$99))}
    {=SUMME(WENN((MONAT($A$1:A6)=MONAT(D1))*(JAHR($A$1:A6)=JAHR(D1));$B$1:B6))}
    --------------------------------------------------------------------------------
    'Alle definierten Namen mit Bezügen auflisten
    'Dieses kleine Makro listet alle definierten Namen und der dazugehörigen Bezüge in einem neuen Arbeitsblatt auf.
    'Der Vollständigkeit halber - es geht auch schneller ohne Makro: Einfach F3 drücken und die Option "Liste einfügen" wählen.
    Sub MWExcelNamenAdresseAuflisten()
    Dim oName As Object
    Dim oAusgabe As Object
    Dim z As Long
    Set oAusgabe = Sheets.Add
    z = 2
    oAusgabe.Cells(1, 1) = "Name"
    oAusgabe.Cells(1, 2) = "Bereich"
    For Each oName In ActiveWorkbook.Names
    oAusgabe.Cells(z, 1) = oName.Name
    oAusgabe.Cells(z, 2) = ActiveWorkbook.Names.Item(oName.Name)
    z = z + 1
    Next oName
    Set oName = Nothing
    Set oAusgabe = Nothing
    End Sub
    --------------------------------------------------------------------------------
    Sub suchen() 
    Dim wb As Workbook 
    Dim ws As Worksheet 
    Dim material As String 
    Dim c As Range 
    material = InputBox("Bitte geben Sie eine Materialnummer ein: ") 
    Set wb = Workbooks.Open(Filename:="D:\Documents\Documents\uni\master-thesis\Prognose\probe\Materialverbrauch_probe") 
    Set ws = wb.Worksheets(1) 
    Set c = ws.Range("A:A").Find(material, LookIn:=xlValues, LookAt:=xlWhole) 
    If Not c Is Nothing Then 
    MsgBox "Wert ist vorhanden in der Zeile " & c.Row 
    Else 
    MsgBox "Wert ist nicht vorhanden!" 
    End If 
    wb.Close savechanges:=False 
    End Sub
    --------------------------------------------------------------------------------
    'Letzte Zeile, letzte Spalte und letzte Zelle per VBA ermitteln 
    'Version 1a Ermittlung der letzten Zeile:
    Public Sub letzte_zeile_1()
    'Hier wird die letzte Zeile ermittelt
    'Egal in welcher Spalte sich die letzte Zeile befindet
    'Es werden alle Spalten geprüft und die letzte Zeile ausgegeben
    letztezeile = Sheets(1).UsedRange.SpecialCells(xlCellTypeLastCell).Row
    MsgBox letztezeile
    End Sub
    Version 1b: Ermittlung letzte Zeile in Spalte A
    Public Sub letzte_zeile_2()
    'Hier wir die letzte Zeile der Spalte A ermittelt
    'letztezeile = ActiveSheet.Cells(65536, 1).End(xlUp).Row 'Bis Excel 2003
    letztezeile = ActiveSheet.Cells(1048576, 1).End(xlUp).Row 'Ab Excel 2007
    MsgBox letztezeile
    End Sub
    'Version 1c: Ermittlung der letzten Zeile in Spalte A (Wird bei Excel-Inside als Standardversion verwendet, da damit eine Versionsunabhängige Funktionsweise garantiert ist)
    Public Sub letzte_zeile_3()
    'Hier wir die letzte Zeile der Spalte A ermittelt
    letztezeile = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox letztezeile
    End Sub
    'Version 1d: Ermittlung der letzten Zeile im benutzten Bereich
    Public Sub letzte_spalte_1()
    'Hier wird die letzte Zeile ermittelt
    'Egal in welcher Spalte sich die letzte Zeile befindet
    'Es werden alle Spalten geprüft und die letzte Zeile ausgegeben
    letztespalte = Sheets(1).UsedRange.SpecialCells(xlCellTypeLastCell).Column
    MsgBox letztespalte
    End Sub
    --------------------------------------------------------------------------------
    Sub letztespaltefinden()
    letztespalte = ActiveSheet.Cells(1, 16384).End(xlToLeft).Column
    Cells(1, letztespalte).Select
    MsgBox letztespalte
    End Sub
    --------------------------------------------------------------------------------
    Sub find_next_in_col()
    Set nexter = Columns(ActiveCell.Column).Find(ActiveCell.Value, After:=Range(ActiveCell.Address))
    nexter.Select
    End Sub
    Public Function status_bar(s As String, t As String)
    Application.StatusBar = s
    ' Wait for 10 seconds.
    Application.Wait (Now + TimeValue("0:00:" + t))
    Application.StatusBar = False
    End Function
    Sub Uppercase()
    'GROSS SCHREIBEN
    Dim rng As Range, cell As Range
    Set rng = Selection
    For Each cell In rng
    If Not cell.HasFormula Then
    cell.Value = UCase(cell.Value)
    End If
    Next cell
    End Sub
    Sub Lowercase()
    'klein schreiben
    Dim rng As Range, cell As Range
    Set rng = Selection
    For Each cell In rng
    If Not cell.HasFormula Then
    cell.Value = LCase(cell.Value)
    End If
    Next cell
    End Sub
    Sub Propercase()
    Dim rng As Range, cell As Range
    Set rng = Selection
    For Each cell In rng
    If Not cell.HasFormula Then
    cell.Value = WorksheetFunction.Proper(cell.Value)
    End If
    Next cell
    End Sub
    Sub berechnung_umstellen()
    aktuell = Application.Calculation
    If aktuell = -4105 Then
    Application.Calculation = xlManual '-4135
    Else
    Application.Calculation = xlAutomatic '-4105
    End If
    End Sub
    Sub sort2lists()
    Application.ScreenUpdating = False
    Dim InputRng1 As Range
    Dim InputRng2 As Range
    'Dim InputStartRow As Integer
    On Error GoTo Endealles
    Set InputRng1 = Application.Selection
    Set InputRng1 = Application.InputBox("Range 1:", InputRng1.Address, Type:=8)
    Set InputRng2 = Application.Selection
    Set InputRng2 = Application.InputBox("Range 2:", InputRng2.Address, Type:=8)
    Dim InputStartRow As Variant
    InputStartRow = InputBox("Ab Zeile:", "Eingabe")
    With ActiveSheet
    lastrow1 = .Cells(.Rows.Count, InputRng1.Column).End(xlUp).Row
    lastrow2 = .Cells(.Rows.Count, InputRng2.Column).End(xlUp).Row
    End With
    abstandspalten = InputRng2.Column - InputRng1.Column
    If lastrow1 > lastrow2 Then
    lastrow = lastrow1
    Else
    lastrow = lastrow2
    End If
    Set rng1 = Range(Cells(1, InputRng1.Column), Cells(lastrow1, InputRng1.Column))
    Set rng2 = Range(Cells(1, InputRng2.Column), Cells(lastrow2, InputRng2.Column))
    For Each cell In rng1
    If Not cell.Value = "" Then
    Set rgFound = rng2.Find(cell.Value)
    If rgFound Is Nothing Then
    cell.Resize(, InputRng2.Columns.Count).Offset(0, abstandspalten).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Else
    abstand = rgFound.Row - cell.Row
    If abstand > 0 Then
    cell.Resize(abstand, InputRng1.Columns.Count).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    End If
    End If
    Next
    Endealles:
    Application.ScreenUpdating = True
    End Sub
    Sub Swap2Cells()
    Dim beide As String
    Dim AddressArray() As String
    beide = Selection.Address
    AddressArray() = Split(beide, ",")
    part1 = Range(AddressArray(0)).Value
    part2 = Range(AddressArray(1)).Value
    Range(AddressArray(0)) = part2
    Range(AddressArray(1)) = part1
    End Sub
    Sub SuchenErsetzen()
    Application.ScreenUpdating = False
    Dim sTxtSuchen As String
    Dim sTxtNeu As String
    sTxtSuchen = InputBox("Zu suchenden Text eingeben:")
    'If sTxtSuchen = "" Then Exit Sub
    sTxtNeu = InputBox("Neuen Text eingeben:")
    'If sTxtNeu = "" Then Exit Sub
    Set InputRng = Application.Selection
    Set InputRng = Application.InputBox("Range :", InputRng.Address, Type:=8)
    InputRng.Select
    Selection.Replace what:=sTxtSuchen, Replacement:=sTxtNeu, Lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Application.ScreenUpdating = True
    End Sub
    Sub JedeNteZeile()
    Dim rng As Range
    Dim InputRng As Range
    Dim OutRng As Range
    Dim yInterval As Integer
    Set InputRng = Application.Selection
    Set InputRng = Application.InputBox("Range :", InputRng.Address, Type:=8)
    yInterval = Application.InputBox("Jede wievielte Zeile selektieren?", Type:=1)
    For i = yInterval To InputRng.Rows.Count Step yInterval ' + 1
    Set rng = InputRng.Cells(i, 1)
    If OutRng Is Nothing Then
    Set OutRng = rng
    Else
    Set OutRng = Application.Union(OutRng, rng)
    End If
    Next
    OutRng.EntireRow.Select
    End Sub
    Sub JedeNteSpalte()
    Dim rng As Range
    Dim InputRng As Range
    Dim OutRng As Range
    Dim xInterval As Integer
    Set InputRng = Application.Selection
    Set InputRng = Application.InputBox("Range :", InputRng.Address, Type:=8)
    xInterval = Application.InputBox("Jede wievielte Spalte selektieren", Type:=1)
    For i = xInterval To InputRng.Columns.Count Step xInterval ' + 1
    Set rng = InputRng.Cells(1, i)
    If OutRng Is Nothing Then
    Set OutRng = rng
    Else
    Set OutRng = Application.Union(OutRng, rng)
    End If
    Next
    OutRng.EntireColumn.Select
    End Sub
    Sub SortiereSheets1()
    Application.ScreenUpdating = False
    Dim WS As Worksheet
    Dim x As Integer
    Dim y As Integer
    Set WS = ActiveSheet
    For x = 1 To ActiveWorkbook.Worksheets.Count
    For y = x To ActiveWorkbook.Worksheets.Count
    If UCase(Worksheets(y).Name) < UCase(Worksheets(x).Name) Then
    Worksheets(y).Move Before:=Worksheets(x)
    End If
    Next y
    Next x
    WS.Activate
    Set WS = Nothing
    Application.ScreenUpdating = True
    End Sub
    Sub SortiereSheets2()
    Application.ScreenUpdating = False
    Dim WS As Worksheet
    Dim x As Integer
    Dim y As Integer
    Set WS = ActiveSheet
    For x = 1 To ActiveWorkbook.Worksheets.Count
    For y = x To ActiveWorkbook.Worksheets.Count
    If UCase(Worksheets(y).Name) > UCase(Worksheets(x).Name) Then
    Worksheets(y).Move Before:=Worksheets(x)
    End If
    Next y
    Next x
    WS.Activate
    Set WS = Nothing
    Application.ScreenUpdating = True
    End Sub
    Sub ListFilesoeffnen()
    ufListFiles.Show
    End Sub
    Sub zelle_nach_links_verschieben()
    Selection.Cut
    ActiveCell.Offset(0, -1).Select
    ActiveSheet.Paste
    End Sub
    Sub zelle_nach_rechts_verschieben()
    Selection.Cut
    ActiveCell.Offset(0, 1).Select
    ActiveSheet.Paste
    End Sub
    Sub Zelle_nach_unten_verschieben()
    Selection.Cut
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    End Sub
    Sub Zelle_nach_oben_verschieben()
    Selection.Cut
    ActiveCell.Offset(-1, 0).Select
    ActiveSheet.Paste
    End Sub
    Sub zellen_einfügen()
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End Sub
    Sub zellen_loschen()
    Selection.Delete Shift:=xlUp
    End Sub
    Sub streifenmachen()
    Application.ScreenUpdating = False
    Dim rng As Range
    Set rng = Selection
    ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "tempTabelle"
    Range("tempTabelle[#All]").Select
    ActiveSheet.ListObjects("tempTabelle").TableStyle = "TableStyleLight2"
    ActiveSheet.ListObjects("tempTabelle").Unlist
    With Selection.Font
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Bold = False
    End With
    Selection.Offset(2, 0).Select
    numRows = Selection.Rows.Count
    numColumns = Selection.Columns.Count
    Selection.Resize(2, numColumns).Select
    Selection.Copy
    Selection.Offset(-2, 0).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone
    Application.CutCopyMode = False
    Selection.Resize(numRows, numColumns).Select
    With Selection
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
    End With
    Application.ScreenUpdating = True
    End Sub
    Sub leere_Zeilen_loeschen()
    rng = Selection.Rows.Count
    ActiveCell.Offset(0, 0).Select
    Application.ScreenUpdating = False
    For i = 1 To rng
    If ActiveCell.Value = "" Then
    Selection.EntireRow.Delete
    Else
    ActiveCell.Offset(1, 0).Select
    End If
    Next i
    Application.ScreenUpdating = True
    End Sub
    Sub JoinText()
    myCol = Selection.Columns.Count
    For i = 1 To myCol
    ActiveCell = ActiveCell.Offset(0, 0) & ActiveCell.Offset(0, i)
    ActiveCell.Offset(0, i) = ""
    Next i
    End Sub
    Sub SelectUsedRange()
    ActiveSheet.UsedRange.Select
    End Sub
    Sub JaNein()
    YesNo = MsgBox("Ja oder Nein?", vbYesNo + vbCritical, "Ja oder doch lieber Nein?")
    Select Case YesNo
    Case vbYes
    MsgBox "Jap" 'Insert your code here if Yes is clicked
    Case vbNo
    MsgBox "Nope" 'Insert your code here if No is clicked
    End Select
    End Sub
    Sub Multi_Ersetzen()
    Dim rngWo As Range
    Dim rngWas As Range
    ''Wenn Bereiche immer gleich:
    'Set rngWo = Range("C2")
    'Set rngWas = Range("A2:A13")
    Set rngWo = Application.Selection
    Set rngWo = Application.InputBox("Wo suchen?:", rngWo.Address, Type:=8)
    Set rngWas = Application.Selection
    Set rngWas = Application.InputBox("Was ersetzen?:", rngWas.Address, Type:=8)
    For Each cell In rngWas
    rngWo.Replace what:=cell.Value, Replacement:=cell.Offset(0, 1).Value, Lookat:= _
    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Next cell
    End Sub
    Sub txtFilesToTabs()
    Dim FilesToOpen
    Dim x As Integer
    Dim wkbAll As Workbook
    Dim wkbTemp As Workbook
    Dim sDelimiter As String
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    sDelimiter = "|"
    FilesToOpen = Application.GetOpenFilename _
    (FileFilter:="Text Files (*.txt), *.txt", _
    MultiSelect:=True, Title:="Text Files to Open")
    If TypeName(FilesToOpen) = "Boolean" Then
    MsgBox "No Files were selected"
    GoTo ExitHandler
    End If
    x = 1
    Set wkbTemp = Workbooks.Open(FileName:=FilesToOpen(x))
    wkbTemp.Sheets(1).Copy
    Set wkbAll = ActiveWorkbook
    wkbTemp.Close (False)
    wkbAll.Worksheets(x).Columns("A:A").TextToColumns _
    Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, _
    ConsecutiveDelimiter:=False, _
    Tab:=False, Semicolon:=False, _
    Comma:=False, Space:=False, _
    Other:=True, OtherChar:="|"
    x = x + 1
    While x <= UBound(FilesToOpen)
    Set wkbTemp = Workbooks.Open(FileName:=FilesToOpen(x))
    With wkbAll
    wkbTemp.Sheets(1).Move After:=.Sheets(.Sheets.Count)
    .Worksheets(x).Columns("A:A").TextToColumns _
    Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, _
    ConsecutiveDelimiter:=False, _
    Tab:=False, Semicolon:=False, _
    Comma:=False, Space:=False, _
    Other:=True, OtherChar:=sDelimiter
    End With
    x = x + 1
    Wend
    ExitHandler:
    Application.ScreenUpdating = True
    Set wkbAll = Nothing
    Set wkbTemp = Nothing
    Exit Sub
    ErrHandler:
    MsgBox Err.Description
    Resume ExitHandler
    End Sub
    Sub Looptest()
    Application.EnableCancelKey = xlErrorHandler
    Application.StatusBar = "Looptest läuft..."
    DoEvents
    On Error GoTo ErrHandler
    Dim x As Long
    Dim y As Long
    Dim lContinue As Long
    y = 1000000000
    For x = 1 To y Step 1
    Next
    Application.EnableCancelKey = xlInterrupt
    Exit Sub
    ErrHandler:
    If Err.Number = 18 Then
    Application.StatusBar = "ESC..."
    DoEvents
    lContinue = MsgBox(prompt:=Format(x / y, "0.0%") & " complete" & vbCrLf & _
    "Weiter?", Buttons:=vbYesNo)
    If lContinue = vbYes Then
    Resume
    Else
    Application.EnableCancelKey = xlInterrupt
    MsgBox ("Programm vorzeitig beendet")
    'Exit Sub
    End If
    End If
    Application.EnableCancelKey = xlInterrupt
    Application.StatusBar = ""
    End Sub
    Public Function UmlauteKlein(Anything As Variant) As Variant
    Dim i As Long
    Dim Ch As String * 1
    Dim Res As String
    If IsNull(Anything) Then Umlaut = Null: Exit Function
    Anything = LCase(Anything)
    For i = 1 To Len(Anything)
    Ch = Mid$(Anything, i, 1)
    Select Case Asc(Ch)
    Case Asc("ä"): Res = Res & "ae"
    Case Asc("ö"): Res = Res & "oe"
    Case Asc("ü"): Res = Res & "ue"
    Case Asc("ß"): Res = Res & "ss"
    Case Else: Res = Res & Ch
    End Select
    Next
    UmlauteKlein = Res
    End Function
    Public Function conc(ByRef rng As Range, Optional ByVal del As String = " ") As String
    Dim cell As Range
    For Each cell In rng
    If cell <> "" Then
    conc = conc & cell & del
    End If
    Next
    If Len(conc) > 0 Then _
    conc = left(conc, Len(conc) - Len(del))
    End Function
    '######################################################################'2 Spalten vergleichen
    '######################################################################
    Public Sub StringsVergleich()
    Application.ScreenUpdating = False
    Dim InputRng1 As Range
    Dim InputRng2 As Range
    Dim InputRng3 As Range
    Dim InputRng4 As Range
    Set InputRng1 = Application.Selection
    Set InputRng1 = Application.InputBox("Range 1:", InputRng1.Address, Type:=8)
    Set InputRng2 = Application.Selection
    Set InputRng2 = Application.InputBox("Range 2:", InputRng2.Address, Type:=8)
    Set InputRng3 = Application.Selection
    Set InputRng3 = Application.InputBox("Range 3:", InputRng3.Address, Type:=8)
    Set InputRng4 = Application.Selection
    Set InputRng4 = Application.InputBox("Range 4:", InputRng4.Address, Type:=8)
    rng1 = InputRng1.Column
    rng2 = InputRng2.Column
    rng3 = InputRng3.Column
    rng4 = InputRng4.Column
    letztezeile = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Dim a() As Byte, b() As Byte, a_$, b_$, i&, j&, d&, u&, l&, x&, y&, f&()
    Const GAP = -1
    Const PAD = "_"
    Columns(rng3).Clear '[3:4].Clear
    Columns(rng4).Clear
    Columns(rng3).Font.Name = "Courier New" '.[c:d].Font.Name = "Courier New"
    Columns(rng4).Font.Name = "Courier New"
    For zeile = 2 To letztezeile
    a_ = ""
    b_ = ""
    a = Cells(zeile, rng1).Text '[a + zeile].Text
    b = Cells(zeile, rng2).Text '[b + zeile].Text
    ReDim f(0 To UBound(b) \ 2 + 1, 0 To UBound(a) \ 2 + 1)
    For i = 1 To UBound(f, 1)
    For j = 1 To UBound(f, 2)
    x = j - 1: Debug.Print x
    y = i - 1: Debug.Print y
    If a(x * 2) = b(y * 2) Then
    d = 1 + f(y, x)
    u = 0 + f(y, j)
    l = 0 + f(i, x)
    Else
    d = -1 + f(y, x)
    u = GAP + f(y, j)
    l = GAP + f(i, x)
    End If
    f(i, j) = Max(d, u, l)
    Next
    Next
    i = UBound(f, 1)
    j = UBound(f, 2)
    On Error Resume Next
    Do
    x = j - 1
    y = i - 1
    d = f(y, x)
    u = f(y, j)
    l = f(i, x)
    Select Case True
    Case Err
    If y < 0 Then GoTo left Else GoTo up
    Case d >= u And d >= l Or Mid$(a, j, 1) = Mid$(b, i, 1)
    diag:
    a_ = Mid$(a, j, 1) & a_
    b_ = Mid$(b, i, 1) & b_
    i = i - 1: j = j - 1
    Case u > l
    up:
    a_ = PAD & a_
    b_ = Mid$(b, i, 1) & b_
    i = i - 1
    Case l > u
    left:
    a_ = Mid$(a, j, 1) & a_
    b_ = PAD & b_
    j = j - 1
    End Select
    Loop Until i < 1 And j < 1
    DecorateStrings a_, b_, Cells(zeile, rng3), Cells(zeile, rng4), PAD 'Range("C" & zeile), Range("D" & zeile), PAD ' '[c2], [d2], PAD
    Next zeile
    Application.ScreenUpdating = True
    End Sub
    Private Function Max(a&, b&, c&) As Long
    Max = a
    If b > a Then Max = b
    If c > b Then Max = c
    End Function
    Private Sub DecorateStrings(a$, b$, rOutA As Range, rOutB As Range, PAD$)
    Dim i&, j&
    FloatArtifacts a, b, PAD
    FloatArtifacts b, a, PAD
    rOutA = a
    rOutB = b
    For i = 1 To Len(a)
    If Mid$(a, i, 1) <> Mid$(b, i, 1) Then
    If Mid$(a, i, 1) <> PAD Then
    rOutA.Characters(i, 1).Font.Color = vbRed
    End If
    End If
    Next
    For i = 1 To Len(b)
    If Mid$(a, i, 1) <> Mid$(b, i, 1) Then
    If Mid$(b, i, 1) <> PAD Then
    rOutB.Characters(i, 1).Font.Color = vbRed
    End If
    End If
    Next
    End Sub
    Private Sub FloatArtifacts(s1$, s2$, PAD$)
    Dim c&, k&, i&, p&
    For i = 1 To Len(s1)
    c = InStr(i, s1, PAD)
    If c Then
    k = 0
    Do
    k = k + 1
    If Mid$(s1, c + k, 1) <> PAD Then
    If Mid$(s2, c, 1) = Mid$(s1, c + k, 1) Then
    p = InStr(c + k, s1, PAD)
    If p < (c + k + 6) And p > 0 Then
    Mid$(s1, c, 1) = Mid$(s1, c + k, 1)
    Mid$(s1, c + k, 1) = PAD
    i = c
    Exit Do
    Else
    i = c + k
    Exit Do
    End If
    Else
    i = c + k
    Exit Do
    End If
    End If
    If c + k > Len(s1) Then Exit Do
    Loop
    Else
    Exit For
    End If
    Next
    End Sub
    '######################################################################'2 Spalten vergleichen
    '######################################################################
    '######################################################################'Dateien_auflisten
    '######################################################################
    'Option Explicit
    Public Ordner As String
    Public erweiterung As String
    Public NextRow As Long
    Sub Dateien_auflisten()
    Ordner = GetFolder(Ordner)
    If Len(Ordner) > 0 Then
    erweiterung = Application.InputBox("Dateien mit welcher Erweiterung?")
    ListFiles (Ordner)
    End If
    End Sub
    Sub ListFiles(Ordner As String)
    Application.ScreenUpdating = False
    Dim objFSO As Scripting.FileSystemObject
    Dim objTopFolder As Scripting.Folder
    Dim strTopFolderName As String
    Columns("A:E").ClearContents
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    Range("A1").Value = "Dateiname"
    Range("B1").Value = "Größe [kb]"
    Range("C1").Value = "Dateityp"
    Range("D1").Value = "Erstelldatum"
    Range("E1").Value = "ShortPath"
    strTopFolderName = Ordner
    If Len(strTopFolderName) > 0 Then
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTopFolder = objFSO.GetFolder(strTopFolderName)
    Call RecursiveFolder(objTopFolder, True)
    End If
    Range("A1").Select
    'Unload UserForm2
    MsgBox "Count: " + CStr(NextRow - 2)
    Application.ScreenUpdating = True
    End Sub
    Sub RecursiveFolder(objFolder As Scripting.Folder, IncludeSubFolders As Boolean)
    Dim objFile As Scripting.File
    Dim objSubFolder As Scripting.Folder
    'Dim NextRow As Long
    NextRow = Cells(Rows.Count, "A").End(xlUp).Row + 1
    'Loop through each file in the folder
    For Each objFile In objFolder.Files
    If Right(objFile, Len(erweiterung)) = erweiterung Then ' Or Right(objFile, 3) = "odx" Then
    ActiveSheet.Hyperlinks.Add Anchor:=Cells(NextRow, "A"), Address:=objFile.ParentFolder, TextToDisplay:=objFile.Name
    Cells(NextRow, "B").Value = objFile.Size / 1000
    Cells(NextRow, "C").Value = objFile.Type
    Cells(NextRow, "D").Value = objFile.DateCreated
    Cells(NextRow, "E").Value = objFile.ShortPath
    NextRow = NextRow + 1
    End If
    Next objFile
    'Loop through files in the subfolders
    If IncludeSubFolders Then
    For Each objSubFolder In objFolder.SubFolders
    Call RecursiveFolder(objSubFolder, True)
    Next objSubFolder
    End If
    End Sub
    Function GetFolder(strPath As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
    .Title = "Ordner auswählen"
    .AllowMultiSelect = False
    ' .InitialFileName = "\\cp091227\IBT\EE_Datenmanagement\EE-Daten\Flashfiles\"
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
    End With
    NextCode:
    GetFolder = sItem
    Set fldr = Nothing
    End Function
    '######################################################################
    'Dateien_auflisten
    '######################################################################
    VBA Code to Unlock a Locked Excel Sheet
    Sub PasswordBreaker()
    'Breaks worksheet password protection.
    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    On Error Resume Next
    For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
    For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
    For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
    ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
    Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
    Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    If ActiveSheet.ProtectContents = False Then
    MsgBox "One usable password is " & Chr(i) & Chr(j) & _
    Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
    Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    Exit Sub
    End If
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next
    End Sub
