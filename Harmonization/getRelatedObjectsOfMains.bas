Attribute VB_Name = "getRelatedObjectsOfMains"

Sub callForCurrent()

Dim fso As FileSystemObject
Set fso = CreateObject("Scripting.FilesystemObject")
Dim txt As Object
Dim strLogFile As String
st = StrReverse(ActiveDocument.Path)
st = Right(st, Len(st) - InStr(1, st, "\", vbTextCompare))
st = StrReverse(st)
strLogFile = st & "\Logger.txt"

On Error Resume Next
txt.Close
On Error Resume Next
Set txt = fso.GetFile(strLogFile)
On Error Resume Next
Kill txt
Set txt = fso.CreateTextFile(strLogFile, False, True)
txt.WriteLine "Main object name,,Left,Top,Width,Height,Area,Type,|,Related object name,,Left,Top,Width,Height,Area,Type,,Connection"
txt.Close
Dim obj As HMIObject
Dim objs As HMIObjects
Set objs = ActiveDocument.HMIObjects
For Each obj In objs
  If InStr(1, obj.ObjectName, "@V3_SMS", vbTextCompare) = 1 And (obj.Width > 30 And obj.Height > 30) Then
    If InStr(1, obj.ObjectName, "ClAt", vbTextCompare) > 0 Then
        GoTo nextobject
    End If
    ActiveDocument.Selection.DeselectAll
    obj.Selected = True
    Call getRelatedObjs(obj.ObjectName)
  End If
nextobject:
Next

End Sub

Sub getRelatedObjs(stdObj As String)

Dim std As HMIObject
Dim obj As HMIObject
Dim objs As HMIObjects
Dim sel(1 To 20) As String

Dim lst As String
Dim tag As String
Dim mk1 As String
Dim muk As String

Dim mrg As Integer
Dim aux As Integer 'for testing if obj has dynamics
Dim s As Integer 'indexing

Dim dyn As HMIDynamicDialog
Dim scr As HMIScriptInfo
Dim var As HMIVariableTrigger
Dim prp As HMIProperty

mrg = 100 'margin of pixels around the standard object

Set std = ActiveDocument.HMIObjects(stdObj)
'Debug.Print std.ObjectName & ", " & std.Left & ", " & std.Top

Set objs = ActiveDocument.HMIObjects

ActiveDocument.Selection.DeselectAll

For i = 1 To std.Properties.Count
  typ = std.Properties(i).DynamicStateType
  If typ <> hmiDynamicStateTypeNoDynamic Then
    If typ = hmiDynamicStateTypeDynamicDialog Then
      Set dyn = std.Properties(i).Dynamic
        With dyn
          If .Trigger.Type = hmiTriggerTypeVariable Then
            If tag = "" And InStr(1, .Trigger.VariableTriggers.Item(1).VarName, "SMS", vbTextCompare) > 0 Then
              tag = .Trigger.VariableTriggers.Item(1).VarName
              mk1 = Right(tag, Len(tag) - (InStr(1, tag, "::", vbTextCompare) + 1))
              mk1 = StrReverse(mk1)
              mk1 = StrReverse(Right(mk1, Len(mk1) - InStr(1, mk1, "_", vbTextCompare)))
              Debug.Print mk1
            End If
          End If
        End With
    End If
  End If
Next

std.Selected = False

For Each obj In objs
aux = 0
  If obj.ObjectName = std.ObjectName Then GoTo nextObj
    If ((obj.Left > (std.Left - mrg)) And (obj.Left < (std.Left + std.Width + 5))) And ((obj.Top > (std.Top - mrg)) And (obj.Top < (std.Top + std.Height + 5))) Then
      lst = lst & obj.ObjectName & vbCrLf
      For i = 1 To obj.Properties.Count
        If obj.Properties(i).DynamicStateType = hmiDynamicStateTypeDynamicDialog Then
          Set dyn = obj.Properties(i).Dynamic
          With dyn
            If .Trigger.Type = hmiTriggerTypeVariable Then
              For v = 1 To .Trigger.VariableTriggers.Count
                If InStr(1, .Trigger.VariableTriggers(v).VarName, mk1, vbTextCompare) > 0 Then
                  'is related
                  aux = aux + 1
                  GoTo selObj
                Else: GoTo nextProp
                End If
              Next
            End If
          End With
        ElseIf obj.Properties(i).DisplayName = "Text" Then
          muk = Replace(obj.Text, " ", "", , , vbTextCompare)
          muk = Replace(muk, "_", "", , , vbTextCompare)
          mk2 = Replace(mk1, " ", "", , , vbTextCompare)
          mk2 = Replace(mk2, "_", "", , , vbTextCompare)
          If muk = mk2 Then GoTo selObj
          If muk <> mk2 Then GoTo nextObj
        End If
nextProp:
      Next
    If aux > 0 Then
selObj:
      obj.Selected = True
      s = s + 1
      sel(s) = obj.ObjectName
    End If
  End If
nextObj:
Next

Rem * writing text file *******************
Dim fso As FileSystemObject
Set fso = CreateObject("Scripting.FilesystemObject")
Dim txt As Object
Dim strLogFile As String
st = StrReverse(ActiveDocument.Path)
st = Right(st, Len(st) - InStr(1, st, "\", vbTextCompare))
st = StrReverse(st)
strLogFile = st & "\Logger.txt"

On Error Resume Next
txt.Close
Set txt = fso.OpenTextFile(strLogFile, ForAppending, False, TristateMixed)

For s = 1 To UBound(sel)
  If (sel(s) <> "") Then
    Set obj = objs(sel(s))
    Dim conn As String
    If obj.Type = "HMICustomizedObject" Then
      For i = 1 To obj.Properties.Count
        If obj.Properties(i).DynamicStateType = hmiDynamicStateTypeDynamicDialog Then
          Set dyn = obj.Properties(i).Dynamic
          With dyn
            If .Trigger.Type = hmiTriggerTypeVariable Then
              For v = 1 To .Trigger.VariableTriggers.Count
                If InStr(1, .Trigger.VariableTriggers(v).VarName, mk1, vbTextCompare) > 0 Then
                  'is related
                  conn = .Trigger.VariableTriggers(v).VarName
                Else: GoTo nextProp2
                End If
              Next
            End If
          End With
        End If
nextProp2:
      Next
    ElseIf obj.Type = "HMIStaticText" Then
      conn = obj.Text
    End If
    jk1 = Space(20 - Len(std.ObjectName))
    jk2 = Space(20 - Len(obj.ObjectName))
    jk3 = Space(20 - Len(obj.Type))
    txt.WriteLine std.ObjectName & ", " & jk1 & ", " & std.Left & ", " & std.Top & ", " & std.Width & ", " & std.Height & ", " & std.Width * std.Height & ", " & std.Type & ", | ," & obj.ObjectName & ", " & jk2 & ", " & obj.Left & ", " & obj.Top & ", " & obj.Width & ", " & obj.Height & ", " & obj.Width * obj.Height & ", " & obj.Type & ", " & jk3 & ", " & conn
  End If
Next
txt.Close
Rem * end text file section ***
End Sub
