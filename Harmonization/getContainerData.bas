Attribute VB_Name = "GetContainerData"
Public k As Integer

Sub multiCall()

  Call defaultAllPages("C:\Users\admin\Desktop\Devices") 'will call singleCall for each file in the specified folder

End Sub

Sub singleCall()

  Call GetMuKData("@V3_SMS_StdDev_DigMsr_", "@V3_SMS_StdDev_DigMsr_ClAt", "_MsAt") 'include string in objectName, do not include string in objectName, suffix

End Sub

Sub defaultAllPages(Path As String)

Rem *******************************************************************************************************
Rem This program cycles through all .pdl files in your folder of choice. For each page it opens it will cycle through all objects it finds in the page.
Rem ensure Microsoft Scripting Runtime is checked in Tools\References menu
Rem
Rem             |------------|------------|------------|
Rem             |   Author   |    Date    |   Version  |
Rem             |   MURA02   | 18.01.2019 |     1.0    |
Rem             |------------|------------|------------|
Rem
Rem *******************************************************************************************************

Dim objDocument As Document
Dim vbs As HMIScriptInfo
Dim FileName As String
Dim FolderPath, strSourceCode As String
Dim fso As FileSystemObject
Dim fdr As Scripting.Folder
Dim f As file
Dim fs As Files
Dim filterPages(10), filterpdl
Dim startTime
Dim endTime
Dim fdate
Dim ndate


startTime = Timer()


Set fso = CreateObject("Scripting.FilesystemObject")

FolderPath = Path

Set fdr = fso.GetFolder(FolderPath)
Set fs = fdr.Files

Rem initialize objects processed counter
objsChanged = 0

For Each f In fs

k = 0
Rem init the property change counter to have a indicator for saving

If InStr(1, f, ".pdl", vbTextCompare) > 0 Then  'filter through .pdl files only

    Rem ** check if date is earlier than * for use to see if file was not recently modified ******
    fdate = CDate(f.DateLastModified)
    ndate = CDate("19/05/2020 01:00")
    If (fdate < ndate) Then
      'MsgBox fdate
    Else
      GoTo nextFile
    End If
    Rem end ** check if date is earlier than * for use to see if file was not recently modified **
        
    On Error Resume Next
    Application.Documents.CloseAll
    Application.Documents.Open f, hmiOpenDocumentTypeVisible
                
        Rem script to be performed on the filtered pages
                               
        Call singleCall

        Rem end of script to be performed
    
    Rem check if the file needs to be saved or not, else it will just close it
    If k > 0 Then
        ActiveDocument.Save
    End If
    ActiveDocument.Close

End If
Rem end multiple page filter *************************************************************
'End If 'end multiple file filter
'Next   'next row in array of files
Rem ***************************************************************************************
nextFile:

Next f 'next file in graCS

exitSub:

Set fso = Nothing
Set fs = Nothing

endTime = Timer()

MsgBox objsChanged & "  objects have been changed in " & FormatNumber(endTime - startTime, 2) & " seconds"

End Sub

Sub GetMuKData(include As String, notinclude As String, suffix As String) 'include string in objectName, do not include string in objectName, suffix

Dim obj As HMIObject
Dim objs As HMIObjects

Dim prop As HMIProperty
Dim props As HMIProperties
Dim var As HMIVariableTrigger
Dim vars As HMIVariableTriggers

Dim dyn As HMIDynamicDialog
Dim vbs As HMIScriptInfo
Dim var As HMIVariableTrigger

Dim oMuK As HMIObject
Dim oTyp As HMIObject
Dim ocont As HMIRectangle
Dim sobj As HMIObject

Dim rng As Integer

Dim MuKNo As String
Dim MuKType As String

Set objs = ActiveDocument.HMIObjects


For Each obj In objs

ActiveDocument.Selection.DeselectAll 'clear selection for next steps

If InStr(1, obj.ObjectName, include, vbBinaryCompare) > 0 And InStr(1, obj.ObjectName, notinclude, vbBinaryCompare) = 0 Then
  GoTo nextobject
  End If

  MuKType = ""
  MuKNo = ""
  suf = suffix
  
  obj.Selected = True
  
tryoContAgain:
  For Each sobj In objs
    If obj.Left < (sobj.Left + sobj.Width) And obj.Left > (sobj.Left) And obj.Top < (sobj.Top + sobj.Height) And obj.Top > (sobj.Top) Then
      If InStr(1, sobj.ObjectName, "@", vbTextCompare) = 0 Then 'layer < 29 ?
        If sobj.BackColor = -2147483556 Then
          Set ocont = sobj
          ocont.Selected = True
          Exit For
        End If
      End If
    End If
  Next
  If ocont Is Nothing Then
    MsgBox "trying search for container object again"
    GoTo tryoContAgain
  End If
  
tryoMuKAgain:
  For Each sobj In objs
    If sobj.Left < (ocont.Left + ocont.Width) And sobj.Left > (ocont.Left) And sobj.Top < (ocont.Top + ocont.Height) And sobj.Top > (ocont.Top) Then
      If (Abs(sobj.Top - (ocont.Top)) > (10 - 5) And Abs(sobj.Top - (ocont.Top)) < (10 + 5)) _
      And (Abs((ocont.Left + ocont.Width) - (sobj.Left + sobj.Width)) < 30) _
      And sobj.Type = "HMIStaticText" Then
        Set oMuK = sobj
        oMuK.Selected = True
        MuKNo = oMuK.Text
        Exit For
      End If
    End If
  Next
  If oMuK Is Nothing Then
    MsgBox "trying search for MuKNo object again"
    GoTo tryoMuKAgain
  End If
  
tryoTypAgain:
  For Each sobj In objs
    If Abs((obj.Top + obj.Height) - sobj.Top) < 8 And Abs(obj.Left - sobj.Left) < 10 And sobj.Type = "HMIStaticText" Then
      Set oTyp = sobj
      MuKType = oTyp.Text
      oTyp.Selected = True
      Exit For
    End If
  Next
  If oTyp Is Nothing Then
    MsgBox "trying search for MukType object again"
    GoTo tryoTypAgain
  End If
  
  If ocont Is Nothing Or oTyp Is Nothing Or oMuK Is Nothing Then
    MsgBox "something is nothing"
  End If
  
  If MuKNo <> "" And MuKType <> "" Then
    tag = "SMS_CSPFM::" & MuKNo & "_" & MuKType & suf
  Else
    MsgBox "empty MuKNo"
  End If
  
Call changesToObject(obj, ocont, oMuK, oTyp) 'desired object, container, muk number for container, muk type for object
  
nextobject:
Next 'end big for

End Sub

Sub changesToObject(o As HMIObject, c As HMIObject, m As HMIObject, t As HMIObject) 'desired object, container, muk number for container, muk type for object
 
  Set props = o.Properties
  For Each prop In props
    If prop.DisplayName = "Symbol - Background Color" Then
'      If prop.DynamicStateType = hmiDynamicStateTypeDynamicDialog Then
'        Set dyn = prop.Dynamic
'          With dyn
'            variable = .sourceCode
'          End With ' does not work , crashes on accessing sourcecode or variable
'      End If
        Set dyn = prop.CreateDynamicDialog("'" & tag & "'", 1)
          With dyn
          .ResultType = hmiResultTypeAnalog
          .Trigger.VariableTriggers.Item(1).CycleType = hmiVariableCycleType_500ms
            If .ResultType = hmiResultTypeAnalog Then
              .AnalogResultInfos.Add 0, -2147483642
              .AnalogResultInfos.Add 1, -2147483641
              .AnalogResultInfos.Add 2, -2147483643
              .AnalogResultInfos.Add 4, -2147483556
              .AnalogResultInfos.Add 6, -2147483639
              .AnalogResultInfos.Add 9, -2147483642
              .AnalogResultInfos.Add 11, -2147483638
              .AnalogResultInfos.Add 13, -2147483637
              .AnalogResultInfos.Add 15, -2147483629
              .AnalogResultInfos.Add 17, -2147483636
              .AnalogResultInfos.Add 100, -2147483642
              .AnalogResultInfos.ElseCase = -2147483642
            End If
          End With
    End If
  Next
  
  k = k + 1
  
End Sub
