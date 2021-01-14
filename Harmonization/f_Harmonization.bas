Attribute VB_Name = "f_Harmonization"
Option Explicit

Public issueCount As Integer

Sub buildUserForm()

Dim uf As UserForm
Dim f As Frame
Dim t As Object
Set uf = HarmonizationHelper

Set f = uf.Frame1
Dim c As CheckBox
Set c = uf.CheckBox1

End Sub

Sub showUserForm()
  HarmonizationHelper.Show
End Sub

Sub HarmonizeContainers(repair As Boolean)
  Dim prop As HMIProperty
  Dim obj As HMIObject
  Dim sobj As HMIObject
  Dim ocont As HMIRectangle
  Dim oconts As Collection
  Dim ounit As hmiIOField
  Dim ounits As New Collection
  Dim otitle As HMIStaticText
  Dim otitles As New Collection
  Dim include, notinclude As String
  Dim foundContainers As New Collection
  Dim foundTitles As New Collection
  Dim c As Variant
  Dim t As New Collection
  Dim ti As HMIRectangle
  Dim i As Integer
  Dim col As HMIObjects
  Dim startTime, endTime
  Dim searchedAreas As New Collection
  Dim area As Variant
  Dim areaLeft, areaTop, areaWidth, areaHeight
  Dim uf As UserForm
  Dim secocont As HMIRectangle
  Dim o As HMIRectangle
  Dim contHighlights As New Collection
  Dim titleHighlights As New Collection
  Dim unitHighlights As New Collection
  Dim offset As Integer
  
  issueCount = 0
  
  startTime = Timer()
  
  Rem restrict searches to relevant objects
  Dim objs As HMIObjects
  Set objs = ActiveDocument.HMIObjects
  Dim oSMS As New Collection 'customized objects
  Set oSMS = GetHMIObjectsByType("HMICustomizedObject", objs, "@V3")
  Dim otitleSet As New Collection
  Set otitleSet = GetHMIObjectsByType("HMIStaticText", objs)
  Dim ounitSet1 As New Collection
  Dim ounitSet2 As New Collection
  Dim ounitSet As New Collection
  Set ounitSet1 = GetHMIObjectsByType("HMIIOField", objs, "@V3_SMS_Unit")
  Set ounitSet2 = GetHMIObjectsByType("HMIIOField", objs, "Unit_Met_Imp")
  Set ounitSet = joinCollections(ounitSet1, ounitSet2)
  Dim potentialContainers As New Collection
  Set potentialContainers = GetHMIObjectsByType("HMIRectangle", objs, , 5000, -2147483557)
  
  For Each obj In oSMS 'looks to standard object, then at possibly existing container from potential set
    ActiveDocument.Selection.DeselectAll 'clear selection for next steps
    'obj.Selected = True
    
    Set oconts = getContainers(obj, objs, 0)
    For Each o In oconts
      'o.Selected = True
      On Error Resume Next
      foundContainers.Add o, o.ObjectName
      searchedAreas.Add o
    Next
  Next

  For Each o In foundContainers
    'tryotitleAgain:
      For Each sobj In otitleSet
        If sobj.Left < (o.Left + o.Width) And sobj.Left > (o.Left) And sobj.Top < (o.Top + o.Height) And sobj.Top < (o.Top) Then
          If (o.Top - sobj.Top < (10 + 5)) _
          And (Abs(o.Left - sobj.Left) < 30) Then
            offset = 0
            If sobj.FONTBOLD = True And sobj.FONTSIZE = 14 Then
              offset = 5
            End If
              Set otitle = sobj
              'otitle.Selected = True
              Rem check if not ok to add
              If (otitle.Left <> o.Left + 10 - offset) Or Left(otitle.Text, 1) <> " " Or Right(otitle.Text, 1) <> " " Or CInt(otitle.Top) <> CInt(o.Top - otitle.Height / 2) Then
                otitles.Add otitle
              End If
              Exit For
          End If
        End If
      Next sobj
      If otitle Is Nothing Then
        'MsgBox "trying search for title object again/there's no title for this container"
        'GoTo noTitle
        'GoTo tryotitleAgain
      Else
        If repair = True Then Call ArrangeContainerTitle(otitle, o, offset)
        For Each sobj In ounitSet
          Set ounit = Nothing
          If (Abs(otitle.Top - sobj.Top) < 5) _
          And (Abs(otitle.Left - sobj.Left) < o.Width - sobj.Width) And sobj.Left > otitle.Left Then
            Set ounit = sobj
            'ounit.Selected = True
            Rem check if not ok to add
            If CInt(ounit.Left) <> CInt(otitle.Left + otitle.Width - 2) Or Right(ounit.InputValue, 1) <> " " Or ounit.Top <> otitle.Top Or ounit.AlignmentLeft <> 0 Or ounit.AdaptBorder <> True Then
              ounits.Add ounit
            Else
              Set ounit = Nothing
            End If
          End If
          If Not ounit Is Nothing And repair = True Then
            ActiveDocument.Selection.DeselectAll
            'ounit.Selected = True
            'otitle.Selected = True
            'o.Selected = True
            Call ArrangeContainerTitleUnit(ounit, otitle, o)
          ElseIf ounit Is Nothing Then
            'MsgBox "trying search for ounit object again"
          End If
        Next sobj
      End If
noTitle:
  Next o
  endTime = Timer()
  
  Rem Set contHighlights = CreateHighlights(searchedAreas)
  Set titleHighlights = CreateHighlights(otitles)
  Set unitHighlights = CreateHighlights(ounits)
  
  issueCount = unitHighlights.Count + titleHighlights.Count
  
  If repair = True Then
    Call DeleteObjectsFoundInCol(contHighlights)
    Call DeleteObjectsFoundInCol(titleHighlights)
    Call DeleteObjectsFoundInCol(unitHighlights)
  End If
  
  Set uf = HarmonizationHelper
  ActiveDocument.Selection.DeselectAll
  On Error Resume Next
  uf.Hide
  
  Dim what As String
  If repair = True Then
    what = "repair of container objects"
  Else
    what = "check of container objects"
  End If
  
  uf.StatusMsg.Caption = "Ran " & what & ", " & issueCount & " issues in " & FormatNumber(endTime - startTime, 3) & " seconds"
End Sub

Sub HarmonizeButtons(repair As Boolean)
  Dim buttons As New Collection
  Dim versionInfos As New Collection
  Dim sorted As New Collection
  Dim grp As New Collection
  Dim elem As Variant
  Dim elemSec As Variant
  Dim bt As HMIButton
  Dim startTime, endTime
  Dim tops As New Collection
  Dim col As New Collection
  Dim indexLeft
  Dim colMain As New Collection
  Dim colSpaces As New Collection
  Dim avgSpce
  Dim spce
  Dim curVersionInfo As HMIStaticText
  
  Dim i As Integer
  Dim j As Long
  Dim l As Integer
  
  Dim highlightMain As New Collection
  Dim highlightSec As New Collection
  Dim difLeft As New Collection
  Dim minLeft As Integer
  
  Dim lessThanTen As New Collection
  Dim difThanAvg As New Collection
  Dim intAvg
  
  Dim uf As UserForm
  Set uf = HarmonizationHelper
  
  Dim minimumSpacing: Dim averageSpacing
  minimumSpacing = CInt(uf.TextBox1.value)
  averageSpacing = CInt(uf.TextBox2.value)
  
  issueCount = 0
  
  startTime = Timer()
  
  Set buttons = GetHMIObjectsByType("HMIButton", ActiveDocument.HMIObjects)
  Set versionInfos = GetHMIObjectsByType("HMIStaticText", ActiveDocument.HMIObjects, "VersionInfo")
    
redoSpacing:
  Set col = getButtonGroups(buttons)
  Set colMain = col(1)
  Set colSpaces = col(2)
  Set col = Nothing
    
  Rem spacing issues find and repair
  For i = 1 To colSpaces.Count
    For j = 1 To colSpaces(i).Count
    If colSpaces(i)(j).Count > 0 Then
      avgSpce = 0
      For Each elem In colSpaces(i)(j)
        avgSpce = avgSpce + elem
      Next
      avgSpce = avgSpce / colSpaces(i)(j).Count
      If avgSpce = Int(avgSpce) Then intAvg = avgSpce
      If intAvg = 0 Then intAvg = 14
    End If
      l = 1
      For Each elem In colSpaces(i)(j)
        Rem avg spacing issues
        Dim fixed As Boolean
        If elem <> averageSpacing Then
          On Error Resume Next
          difThanAvg.Add colMain(i)(j)(l + 1), colMain(i)(j)(l + 1).ObjectName 'above this, dif than avg spacing
          colMain(i)(j)(l + 1).Selected = True
          If repair = True Then 'repair top here
          Dim term
          term = 1
            If l = 1 Then
              l = 0: term = -1
            End If
            Set curVersionInfo = GetObjectFromCollectionByCoordinates(versionInfos, colMain(i)(j)(l + 1).Left, colMain(i)(j)(l + 1).Top, 2, 2)
            colMain(i)(j)(l + 1).Top = colMain(i)(j)(l + 1).Top - term * (elem - averageSpacing)
            curVersionInfo.Top = colMain(i)(j)(l + 1).Top
            GoTo redoSpacing
            fixed = True
          End If
        End If
        Rem less than 10 px spacing issues
        If elem < minimumSpacing Then
          On Error Resume Next
          lessThanTen.Add colMain(i)(j)(l + 1), colMain(i)(j)(l + 1).ObjectName 'above this, less than 10
          colMain(i)(j)(l + 1).Selected = True
          If repair = True And fixed = False Then 'repair top here
            Set curVersionInfo = GetObjectFromCollectionByCoordinates(versionInfos, colMain(i)(j)(l + 1).Left, colMain(i)(j)(l + 1).Top, 2, 2)
            colMain(i)(j)(l + 1).Top = colMain(i)(j)(l + 1).Top - (elem - minimumSpacing)
            curVersionInfo.Top = colMain(i)(j)(l + 1).Top
            If l < colSpaces(i)(j).Count Then colSpaces(i)(j)(l + 1) = colSpaces(i)(j)(l + 1) - (elem - minimumSpacing)
            elem = minimumSpacing
          End If
        End If
        l = l + 1
      Next
    Next
  Next
    
  issueCount = issueCount + ActiveDocument.Selection.Count
  ActiveDocument.Selection.DeselectAll
  
  Rem left alignment issues (wrt leftmost coordinate)
  For i = 1 To colMain.Count
    For l = 1 To colMain(i).Count
      minLeft = 3000
      Set difLeft = Nothing
      For Each elem In colMain(i)(l)
        If elem.Left < minLeft Then minLeft = elem.Left
      Next
      j = 1
      For Each elem In colMain(i)(l)
        If elem.Left <> minLeft Then
          ActiveDocument.Selection.DeselectAll
          colMain(i)(l)(j).Selected = True
          On Error Resume Next
          difLeft.Add colMain(i)(l)(j), colMain(i)(l)(j).ObjectName.value
          If repair = True Then 'repair left here
            Set curVersionInfo = GetObjectFromCollectionByCoordinates(versionInfos, colMain(i)(l)(j).Left, colMain(i)(l)(j).Top, 2, 2)
            colMain(i)(l)(j).Left = minLeft
            curVersionInfo.Left = colMain(i)(l)(j).Left
          End If
        End If
        j = j + 1
      Next
      Set highlightSec = CreateLeftHighlights(difLeft, minLeft) 'highlighting left alignment issues compared to min left
      issueCount = issueCount + highlightSec.Count
    Next
  Next
  
  ActiveDocument.Selection.DeselectAll
  
  Set highlightMain = CreateHighlights(difThanAvg) 'will only highlight what was moved already
  Set highlightMain = CreateHighlights(lessThanTen) 'will only highlight what was moved already
  
  ActiveDocument.Selection.DeselectAll
  
  endTime = Timer()
  
  Dim what As String
  If repair = True Then
    what = "repair of button groups"
  Else
    what = "check of button groups"
  End If
  
  uf.StatusMsg.Caption = "Ran " & what & ", " & issueCount & " issues in " & FormatNumber(endTime - startTime, 3) & " seconds"
End Sub

Sub HarmonizeUnits(repair As Boolean)
  Dim ioFields As New Collection
  Dim io As hmiIOField
  Dim startTime, endTime
  
  issueCount = 0
  
  Rem restrict searches to relevant objects
  Dim objs As HMIObjects
  Set objs = ActiveDocument.HMIObjects
  
  startTime = Timer()
  
  Dim oActObjLeft As HMIObject
  Dim oActObjRight As HMIObject
  
  Dim oActRef As New Collection 'customized objects
  Set oActRef = joinCollections(GetHMIObjectsByType("HMICustomizedObject", objs, "ActV"), GetHMIObjectsByType("HMICustomizedObject", objs, "RefV"))
  Set oActRef = joinCollections(oActRef, GetHMIObjectsByType("HMICustomizedObject", objs, "AnaMsr"))
  Set oActRef = joinCollections(oActRef, GetHMIObjectsByType("HMICustomizedObject", objs, "_Pid"))
  
  Dim ounitSet As New Collection
  Set ounitSet = joinCollections(GetHMIObjectsByType("HMIIOField", objs, "@V3_SMS_Unit"), GetHMIObjectsByType("HMIIOField", objs, "Unit_Met"))
  
  Dim objsInContainer As New Collection
  Dim oconts As New Collection
      
  Dim ioLeft
  Dim ioTop
  
  Dim isPolygonContainer
  
  Dim issuesIo As New Collection
    
  Dim potentialContainers As New Collection
  Set potentialContainers = GetHMIObjectsByType("HMIRectangle", objs, , 5000, -2147483557)
      
  Set ioFields = GetHMIObjectsByType("HMIIOField", ActiveDocument.HMIObjects)
  
  For Each io In ounitSet
  ActiveDocument.Selection.DeselectAll
    'check if object is near actV
    'check if object has parantheses or not for appropriate cases
    
    'If io.FONTBOLD = True Then GoTo nextio
    'io.Selected = True
    
    'find if in container !!!!!!!!!!!!!!!!!!!!1111
    Set oconts = Nothing
    Set oconts = getContainers(io, objs, 3)
    If oconts.Count = 1 Then
      Set objsInContainer = ObjectsInContainer(oconts(1), objs, 5)
      If oconts(1).Type = "HMIPolygon" Then
        isPolygonContainer = 1
      ElseIf oconts(1).Type = "HMIRectangle" Then
        isPolygonContainer = 0
      End If
      'Call selectCollection(objsInContainer)
    ElseIf oconts.Count > 1 Then
      If oconts(1).Left = oconts(2).Left And oconts(1).Top = oconts(2).Top And oconts(1).Width = oconts(2).Width And oconts(1).Height = oconts(2).Height Then
        objs(oconts(2).ObjectName).delete 'delete duplicate
        oconts.Remove (2)
      Else
        'MsgBox "more than 1 container for this io field"
      End If
    'Else: MsgBox "page has no container": Exit For
    End If
    'Call selectCollection(oconts)
    
    'find object
    Set oActObjLeft = GetObjectFromCollectionByCoordinates(oActRef, io.Left - 70, io.Top - 5, 15, 10)
    Set oActObjRight = GetObjectFromCollectionByCoordinates(oActRef, io.Left + io.Width, io.Top - 5, 10, 10)
    
    'check for both
    If Not oActObjRight Is Nothing And Not oActObjLeft Is Nothing Then
      MsgBox "getting both right and left act ref here"
    End If
    
    If Not oActObjLeft Is Nothing And isPolygonContainer = 0 Then
      'oActObjLeft.Selected = True
        If oActObjLeft.Width = 78 Then 'has fph handling
        ioLeft = CInt(oActObjLeft.Left + oActObjLeft.Width + 1) 'account if width is different?
      Else
        ioLeft = CInt(oActObjLeft.Left + oActObjLeft.Width + 1 + 3) 'account if width is different?
      End If
      ioTop = oActObjLeft.Top + RoundUp((oActObjLeft.Height - io.Height) / 2) - 1
'      If repair = False Then
        If io.Left <> ioLeft Or io.Top <> ioTop Or io.AdaptBorder <> True Or io.AlignmentLeft <> 0 Or io.AlignmentTop <> 1 _
        Or InStr(1, io.InputValue, " ", vbBinaryCompare) _
        Or InStr(1, io.InputValue, "[", vbBinaryCompare) _
        Or InStr(1, io.InputValue, "]", vbBinaryCompare) Then
          issuesIo.Add io, io.ObjectName
        End If
'      End If
      If repair = True Then
        io.AdaptBorder = True
        io.Left = ioLeft
        io.Top = ioTop
        io.AlignmentLeft = 0
        io.AlignmentTop = 1
        io.FONTBOLD = False
        If InStr(1, io.InputValue, " ", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, " ", "", , , vbBinaryCompare)
        If InStr(1, io.InputValue, "[", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, "[", "", , , vbBinaryCompare)
        If InStr(1, io.InputValue, "]", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, "]", "", , , vbBinaryCompare)
      End If
    ElseIf Not oActObjRight Is Nothing And isPolygonContainer = 0 Then
      'oActObjRight.Selected = True
      If oActObjRight.Width = 78 Then 'has fph handling
        ioLeft = CInt(oActObjRight.Left - io.Width) 'account if width is different?
      Else
        ioLeft = CInt(oActObjRight.Left - io.Width - 3) 'account if width is different?
      End If
      ioTop = oActObjRight.Top + RoundUp((oActObjRight.Height - io.Height) / 2) - 1
      'ActiveDocument.selection.DeselectAll
      'io.Selected = True
      'oActObjRight.Selected = True
'      If repair = False Then
        If io.Left <> ioLeft Or io.Top <> ioTop Or io.AlignmentLeft <> 2 Or io.AlignmentTop <> 1 _
        Or InStr(1, io.InputValue, " ", vbBinaryCompare) _
        Or InStr(1, io.InputValue, "[", vbBinaryCompare) = 0 _
        Or InStr(1, io.InputValue, "]", vbBinaryCompare) = 0 Then
          issuesIo.Add io, io.ObjectName
        End If
'      End If
      If repair = True Then
        io.Left = ioLeft
        io.Top = ioTop
        io.AlignmentLeft = 2
        io.AlignmentTop = 1
        io.FONTBOLD = False
        If InStr(1, io.InputValue, " ", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, " ", "", , , vbBinaryCompare)
        If InStr(1, io.InputValue, "[", vbBinaryCompare) = 0 Then io.InputValue = "[" & io.InputValue
        If InStr(1, io.InputValue, "]", vbBinaryCompare) = 0 Then io.InputValue = io.InputValue & "]"
      End If
    ElseIf isPolygonContainer = 1 Then
      'do different actions ...should be bolded, et al
      If Not oActObjLeft Is Nothing Then
        If oActObjLeft.Width = 78 Then 'has fph handling
          ioLeft = CInt(oActObjLeft.Left + oActObjLeft.Width - 3)
        Else
          ioLeft = CInt(oActObjLeft.Left + oActObjLeft.Width)
        End If
        ioTop = oActObjLeft.Top + RoundUp((oActObjLeft.Height - io.Height) / 2)
          If io.Left <> ioLeft Or io.Top <> ioTop Or io.AdaptBorder <> True Or io.AlignmentLeft <> 0 Or io.AlignmentTop <> 1 _
          Or InStr(1, io.InputValue, " ", vbBinaryCompare) _
          Or InStr(1, io.InputValue, "[", vbBinaryCompare) _
          Or InStr(1, io.InputValue, "]", vbBinaryCompare) Then
            issuesIo.Add io, io.ObjectName
          End If
        If repair = True Then
          io.AdaptBorder = True
          io.Left = ioLeft
          io.Top = ioTop
          io.AlignmentLeft = 0
          io.AlignmentTop = 1
          io.FONTBOLD = True
          If InStr(1, io.InputValue, " ", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, " ", "", , , vbBinaryCompare)
          If InStr(1, io.InputValue, "[", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, "[", "", , , vbBinaryCompare)
          If InStr(1, io.InputValue, "]", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, "]", "", , , vbBinaryCompare)
        End If
      ElseIf Not oActObjRight Is Nothing Then
        If oActObjRight.Width = 78 Then 'has fph handling
          ioLeft = CInt(oActObjRight.Left - io.Width) 'account if width is different?
        Else
          ioLeft = CInt(oActObjRight.Left - io.Width - 3) 'account if width is different?
        End If
        ioTop = oActObjRight.Top + RoundUp((oActObjRight.Height - io.Height) / 2) - 1
          If io.Left <> ioLeft Or io.Top <> ioTop Or io.AlignmentLeft <> 2 Or io.AlignmentTop <> 1 _
          Or InStr(1, io.InputValue, " ", vbBinaryCompare) _
          Or InStr(1, io.InputValue, "[", vbBinaryCompare) = 0 _
          Or InStr(1, io.InputValue, "]", vbBinaryCompare) = 0 Then
            issuesIo.Add io, io.ObjectName
          End If
        If repair = True Then
          io.Left = ioLeft
          io.Top = ioTop
          io.AlignmentLeft = 2
          io.AlignmentTop = 1
          io.FONTBOLD = False
          If InStr(1, io.InputValue, " ", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, " ", "", , , vbBinaryCompare)
          If InStr(1, io.InputValue, "[", vbBinaryCompare) = 0 Then io.InputValue = "[" & io.InputValue
          If InStr(1, io.InputValue, "]", vbBinaryCompare) = 0 Then io.InputValue = io.InputValue & "]"
        End If
      End If
    Else
      'MsgBox "no ActObj !"
    End If
nextio:
  Next
  
  Set oconts = Nothing
  If repair = False Then Set oconts = CreateHighlights(issuesIo)  'reuse coll for useless highlights...
  
  Dim what As String
  If repair = True Then
    what = "repair of unit io fields"
  Else
    what = "check of unit io fields"
  End If
  
  Dim uf As UserForm
  Set uf = HarmonizationHelper
  uf.StatusMsg.Caption = "Ran " & what & ", " & issuesIo.Count & " issues in " & FormatNumber(Timer() - startTime, 3) & " seconds"
End Sub

Rem /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Rem /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Rem //////////////////////////// FUNCTIONS //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Rem /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Rem /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Rem *************************** Buttons *****************************************************************************************************************************************************

Public Function GetObjectFromCollectionByCoordinates(col As Collection, l As Variant, t As Variant, toleranceLeft As Integer, toleranceTop As Integer) As HMIObject
  Dim elem As HMIObject
  Dim o As HMIObject
  
  For Each elem In col
    If Abs(elem.Left - l) < toleranceLeft And Abs(elem.Top - t) < toleranceTop Then
      Set o = elem
      Exit For
    End If
  Next
  
  Set GetObjectFromCollectionByCoordinates = o
End Function

Public Function GetNearbyVerticalButtonsGroup(c As Variant, buttons As Collection, heightMargin As Integer, widthMargin As Integer, leftMargin As Integer) As Collection
  Dim b As HMIButton
  Dim cl1 As New Collection
  Dim cl2 As New Collection
  Dim elem As HMIButton
  Dim col As New Collection

  Rem go on a search from the found button, upwards until interval gets too large then do not include
  Rem same goes for search to bottom, and remove from main button collection set from which to get related cols

  c.Selected = True
  On Error Resume Next
  buttons.Remove c.ObjectName
  For Each b In buttons
    Rem less than 21 deviation from left/right
    Rem spacing margin tolerance
    Rem buttons will be grouped as long as there's less than 5 pixels difference in width
    b.Selected = True
    If Abs(b.Left - c.Left) < leftMargin And (Abs(c.Top - b.Top - b.Height) < heightMargin Or Abs(c.Top + c.Height - b.Top) < heightMargin) And Abs(b.Width - c.Width) < widthMargin Then
      ActiveDocument.Selection.DeselectAll
      b.Selected = True
      buttons.Remove b.ObjectName
      Set cl2 = GetNearbyVerticalButtonsGroup(b, buttons, heightMargin, widthMargin, leftMargin)
    End If
  Next

  For Each elem In ActiveDocument.Selection
    col.Add elem, elem.ObjectName
  Next

  Set GetNearbyVerticalButtonsGroup = col
End Function

Public Function IndexOf(ByVal coll As Collection, ByVal Item As Variant) As Long
    Dim i As Long
    For i = 1 To coll.Count
        If coll(i).ObjectName.value = Item.ObjectName.value Then
            IndexOf = i
            Exit Function
        End If
    Next
End Function

Public Function getButtonGroups(col As Collection) As Collection
  Dim i As Integer
  Dim j As Long
  Dim l As Integer
  
  Dim indexLeft
  Dim spce
  
  Dim elem As HMIObject
  
  Dim colSpaces As New Collection
  Dim colMain As New Collection
  Dim colVert As New Collection
  Dim spcSec As New Collection
  Dim emptyCol As New Collection
  Dim emptyColSpc As New Collection
  Dim sorted As New Collection
  
  Dim colSec As New Collection
  Dim highlightMain As New Collection
  Dim highlightSec As New Collection
  Dim difLeft As New Collection
  Dim minLeft As Integer
  
  Set sorted = SortCollectionByProperty(col, "Left")
  
  indexLeft = 0
  i = 1
  colVert.Add emptyCol, CStr(i)
  For Each elem In sorted
    If indexLeft = 0 Then indexLeft = elem.Left
    If elem.Left > indexLeft And Abs(elem.Left - indexLeft) > 10 Then  '10 is the precision for error along vertical group, even 50 should be fine
      i = i + 1
      Set emptyCol = Nothing
      colVert.Add emptyCol
      indexLeft = elem.Left
    End If
    colVert(i).Add elem
  Next
  
  For j = 1 To colVert.Count
    i = 1: l = 1
    Set colSec = Nothing: Set spcSec = Nothing
    Set emptyCol = Nothing:  Set emptyColSpc = Nothing
    Set sorted = SortCollectionByProperty(colVert(j), "Top")
    colSec.Add emptyCol
    spcSec.Add emptyColSpc
    For Each elem In sorted
      If i > 0 And i < sorted.Count Then
        spce = Abs(sorted(i).Top + sorted(i).Height - sorted(i + 1).Top)
        If spce > 20 Then
          colSec(l).Add sorted(i), sorted(i).ObjectName.value
          l = l + 1
          Set emptyCol = Nothing: colSec.Add emptyCol
          Set emptyColSpc = Nothing: spcSec.Add emptyColSpc
        Else
          colSec(l).Add sorted(i), sorted(i).ObjectName.value
          spcSec(l).Add spce, CStr(i)
        End If
        i = i + 1
        If i = sorted.Count Then colSec(l).Add sorted(i), sorted(i).ObjectName.value
      End If
    Next
  colMain.Add colSec
  colSpaces.Add spcSec
  Next
  
  Set colSec = Nothing
  colSec.Add colMain
  colSec.Add colSpaces
    
Set getButtonGroups = colSec
End Function

Rem *************************** Containers **************************************************************************************************************************************************

Public Function getContainers(o As Variant, os As HMIObjects, maxDeviationOutside As Integer) As Collection
  Dim sobj As Variant
  Dim prop As HMIProperty
  Dim conts As New Collection
  Dim potentialContainers As Collection
  
  Set potentialContainers = GetHMIObjectsByType("HMIRectangle", os, , 5000, -2147483557)
  Set potentialContainers = joinCollections(potentialContainers, GetHMIObjectsByType("HMIPolygon", os, , 10, -2147483557))
  Set potentialContainers = joinCollections(potentialContainers, GetHMIObjectsByType("HMIPolygon", os, , 10, -2147483556))
  Set potentialContainers = joinCollections(potentialContainers, GetHMIObjectsByType("HMIRectangle", os, , 5000, -2147483556))
  
  'Call selectCollection(potentialContainers)
  
  For Each sobj In potentialContainers
  Set prop = sobj.BackColor
    If ((o.Left + o.Width) - (sobj.Left + sobj.Width)) <= maxDeviationOutside And (o.Left - sobj.Left) >= -maxDeviationOutside And _
    ((o.Top + o.Height) - (sobj.Top + sobj.Height)) <= maxDeviationOutside And (o.Top - sobj.Top) >= -maxDeviationOutside And _
    sobj.Type <> "HMIGroup" Then
      If InStr(1, sobj.ObjectName, "@", vbTextCompare) = 0 Then 'layer < 29 ?
        conts.Add sobj
      End If
    End If
  Next
  
  'Call selectCollection(conts)
  
  Set getContainers = conts
End Function

Private Sub ArrangeContainerTitle(t As HMIStaticText, c As HMIRectangle, off As Integer)
  If Left(t.Text, 1) <> " " Then t.Text = " " & t.Text
  If Right(t.Text, 1) <> " " Then t.Text = t.Text & " "
  t.Left = c.Left + 10 - off
  t.Top = c.Top - t.Height / 2
End Sub

Private Sub ArrangeContainerTitleUnit(u As hmiIOField, t As HMIStaticText, c As HMIRectangle)
  'If CInt(ounit.left) <> CInt(otitle.left + otitle.Width - 1) Or Right(ounit.InputValue, 1) <> " " Or ounit.top <> otitle.top Or ounit.AlignmentLeft <> 0 Then
  u.InputValue = Replace(u.InputValue, " ", "", , , vbBinaryCompare)
  
  If Left(u.InputValue, 1) = " " Then u.InputValue = Right(u.InputValue, Len(u.InputValue - 1))
  If Right(u.InputValue, 1) <> " " Then u.InputValue = u.InputValue & " "
  
  u.Left = t.Left + t.Width - 2
  u.AdaptBorder = True
  u.AlignmentLeft = 0
  u.Top = t.Top
  Dim vb As HMIScriptInfo
  Set vb = u.OutputValue.Dynamic
    With vb 'will have to check sourcecode when selecting units to change...
      Dim src As String
      src = .sourceCode
      If InStr(1, src, """]""", vbBinaryCompare) Then
        .sourceCode = Replace(.sourceCode, """]""", """] """, , , vbBinaryCompare)
      End If
    End With
End Sub

Private Function ObjectsInContainer(c As HMIObject, os As HMIObjects, tol As Integer) As Collection
  Dim o As HMIObject
  Dim col As New Collection
  
  For Each o In os
    If (o.Left > c.Left Or Abs(o.Left - c.Left) <= tol) _
    And ((o.Left + o.Width) <= (c.Left + c.Width)) _
    And ((o.Top < c.Top) Or Abs(o.Top - c.Top) <= tol) _
    And Abs((o.Top + o.Height) - (c.Top + o.Height)) <= tol Then
      col.Add o, o.ObjectName
    End If
  Next
  col.Remove (c.ObjectName)
  Set ObjectsInContainer = col
End Function

Rem *************************** Common *****************************************************************************************************************************************************

Public Function SortCollectionByProperty(colInput As Collection, prop As String) As Collection
  Dim iCounter As Integer
  Dim iCounter2 As Integer
  Dim temp As Variant
  
  For iCounter = 1 To colInput.Count - 1
      For iCounter2 = iCounter + 1 To colInput.Count
          If colInput(iCounter).Properties(prop) > colInput(iCounter2).Properties(prop) Then
             Set temp = colInput(iCounter2)
             colInput.Remove iCounter2
             colInput.Add temp, temp.ObjectName, iCounter
          End If
      Next iCounter2
  Next iCounter
  Set SortCollectionByProperty = colInput
End Function

Public Function joinCollections(col1 As Collection, col2 As Collection) As Collection
  Dim mainCol As New Collection
  Dim elem As Variant
  
  For Each elem In col1
    On Error Resume Next
    mainCol.Add elem, elem.ObjectName
  Next
  For Each elem In col2
    On Error Resume Next
    mainCol.Add elem, elem.ObjectName
  Next
  
  Set joinCollections = mainCol
End Function

Sub selectCollection(col As Collection)
  Dim elem As HMIObject
  
  For Each elem In col
    elem.Selected = True
    Debug.Print elem.ObjectName
  Next
End Sub

Private Function CreateLeftHighlights(col As Collection, leftValue As Integer) As Collection
  Dim a As HMIObject
  Dim newCol As New Collection
  Dim o As HMILine
  Dim l, t, w, h

  For Each a In col
    l = a.Left 'CInt(s(0))
    t = a.Top 'CInt(s(1))
    w = a.Width 'CInt(s(2))
    h = a.Height 'CInt(s(3))
    Set o = ActiveDocument.HMIObjects.AddHMIObject("HarmonizationAuxiliaries", "HMILine")
      With o
        .Left = leftValue
        .Top = a.Top
        .Height = a.Height
        .Width = 0
        .BorderColor = RGB(255, 201, 14) 'gold
        .Transparency = 50
        .BorderStyle = 512
        .BorderEndStyle = 514
        .BorderBackColor = RGB(255, 201, 14) 'gold
        .BorderWidth = 6
        .GlobalColorScheme = False
        .GlobalShadow = False
        .Left = .Left - (.BorderWidth / 2) - 1
      End With
    newCol.Add o
  Next

  Set CreateLeftHighlights = newCol
End Function

Private Function CreateHighlights(col As Collection) As Collection
  Dim a As HMIObject
  Dim newCol As New Collection
  Dim o As HMIRectangle
  Dim l, t, w, h

  For Each a In col
    's = Split(a, ",", , vbBinaryCompare)
    l = a.Left 'CInt(s(0))
    t = a.Top 'CInt(s(1))
    w = a.Width 'CInt(s(2))
    h = a.Height 'CInt(s(3))
    Set o = ActiveDocument.HMIObjects.AddHMIObject("HarmonizationAuxiliaries", "HMIRectangle")
    o.Left = l
    o.Top = t
    o.Width = w
    o.Height = h
    o.Transparency = 50
    o.BackColor = RGB(255, 201, 14) 'gold
    o.Layer = 27
    o.GlobalColorScheme = False
    o.GlobalShadow = False
    newCol.Add o
  Next

  Set CreateHighlights = newCol
End Function

Sub DeleteObjectsFoundInCol(col As Collection)
  Dim elem As HMIObject
  For Each elem In col
    elem.delete
  Next
End Sub

Sub DeleteHarmonizationAuxiliaries()
  Dim obj As HMIObject
  Dim startt
  startt = Timer()
  Dim objs As HMIObjects
  Set objs = ActiveDocument.HMIObjects
  
  issueCount = 0
  For Each obj In objs
    If InStr(1, obj.ObjectName, "Harmonization", vbBinaryCompare) = 1 Then
      obj.delete
      issueCount = issueCount + 1
    End If
  Next
  
  Dim uf As UserForm
  Set uf = HarmonizationHelper
  uf.StatusMsg.Caption = "Cleared issue " & issueCount & " highlights in " & FormatNumber(Timer() - startt, 3) & " seconds"
End Sub

Public Function GetHMIObjectsByType(typ As String, os As HMIObjects, Optional str As String, Optional area As Variant, Optional bacColor As Long, Optional borColor As Long) As Collection
' restrict to certain type
' added area to check for area occupied by object ..
  Dim o As HMIObject
  Dim ba
  Dim bo
  Dim col As New Collection
  Dim elem As HMIObject
  
  For Each o In os
    If o.Type = typ Then
      col.Add o, o.ObjectName
    End If
    If IsMissing(str) = False Then
      If str <> "" Then
        If InStr(1, o.ObjectName, str, vbBinaryCompare) Then
        Else
          On Error Resume Next
          col.Remove (o.ObjectName)
        End If
      End If
    End If
    If IsMissing(area) = False Then
      If o.Width * o.Height >= area Then
      Else
        On Error Resume Next
        col.Remove (o.ObjectName)
      End If
    End If
    If IsMissing(bacColor) = False And bacColor <> 0 Then
      On Error Resume Next
      ba = o.BackColor.value
      If ba <> 0 Then
        If ba = bacColor Or ba = bacColor - 1 Or ba = bacColor + 1 Then
        Else
          On Error Resume Next
          col.Remove (o.ObjectName)
        End If
      End If
    End If
    If IsMissing(borColor) = False And borColor <> 0 Then
      On Error Resume Next
      bo = o.BorderColor.value
      If bo = 0 Then
        If bo = borColor Or bo = borColor - 1 Or bo = borColor + 1 Then
        Else
          On Error Resume Next
          col.Remove (o.ObjectName)
        End If
      End If
    End If
  Next
  
  Set GetHMIObjectsByType = col
End Function

Public Function InCollection(col As Collection, Key As String) As Boolean
  Dim var As Variant
  Dim errNumber As Long

  InCollection = False
  Set var = Nothing

  Err.Clear
  On Error Resume Next
    var = col(Key)
    errNumber = CLng(Err.Number)
  On Error GoTo 0

  '5 is not in, 0 and 438 represent incollection
  If errNumber = 5 Then ' it is 5 if not in collection
    InCollection = False
  Else
    InCollection = True
  End If
End Function

Function RoundUp(ByVal d As Double) As Integer
    Dim result As Integer
    result = Math.Round(d)
    If result >= d Then
        RoundUp = result
    Else
        RoundUp = result + 1
    End If
End Function
