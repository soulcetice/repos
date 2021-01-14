VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HarmonizationHelper 
   Caption         =   "HarmonizationHelper"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   OleObjectBlob   =   "HarmonizationHelper.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HarmonizationHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Categories_Click()

End Sub

Private Sub Check_Click()

Dim cont As CheckBox
Dim button As CheckBox
Dim units As CheckBox

Set button = HarmonizationHelper.ButtonsCheck
Set cont = HarmonizationHelper.ContainersCheck
Set units = HarmonizationHelper.UnitsCheck

'Call DeleteHarmonizationAuxiliaries

If button.value = True Then
  'MsgBox "Will check buttons!"
  Call HarmonizeButtons(False)
End If
If cont.value = True Then
  'MsgBox "Will check containers!"
  Call HarmonizeContainers(False) 'false to only check
End If
If units.value = True Then
  'MsgBox "Will check containers!"
  Call HarmonizeUnits(False) 'false to only check
End If

End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub CommandButton1_Click()

Dim cont As CheckBox
Dim button As CheckBox
Dim units As CheckBox

Set button = HarmonizationHelper.ButtonsCheck
Set cont = HarmonizationHelper.ContainersCheck
Set units = HarmonizationHelper.UnitsCheck

If button.value = True Then
  MsgBox "Will harmonize buttons!"
  Call HarmonizeButtons(True)
End If
If cont.value = True Then
  MsgBox "Will harmonize containers!"
  Call HarmonizeContainers(True) 'true to repair as well
End If
If units.value = True Then
  MsgBox "Will harmonize units!"
  Call HarmonizeUnits(True) 'false to only check
End If

End Sub

Private Sub CommandButton2_Click()
Call DeleteHarmonizationAuxiliaries
End Sub

Private Sub ContainersCheck1_Click()

End Sub

Private Sub StatusMsg_Click()

End Sub

Private Sub TabStrip1_Change()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UnitsCheck_Click()

End Sub

Private Sub UserForm_Click()

End Sub

