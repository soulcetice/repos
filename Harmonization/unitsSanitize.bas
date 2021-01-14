Attribute VB_Name = "unitsSanitize"
Sub unitsSanitize()

'this script creates the VBScript required for unit output values, from a direct variable

Dim obj As HMIObject
Dim objs As HMIObjects
Dim objVBS As HMIScriptInfo
Dim objVar As HMIVariableTrigger
Dim varTrig As String
Dim dynSourceCode As String

Set objs = ActiveDocument.HMIObjects

'Set obj = objs("@V3_SMS_Unit_87")

''***************************************************************************************************
''Unit Display
'    Item.OutputValue = "[" & HMIRuntime.Tags("@NOP::#UChg_perc_perc.Unit").Read & "]"
''***************************************************************************************************

For Each obj In objs

    If InStr(1, obj.ObjectName, "Unit_", vbTextCompare) Then
        obj.ObjectName = Replace(obj.ObjectName, "Unit_", "@V3_SMS_Unit_", , , vbTextCompare)
    End If

    If InStr(1, obj.ObjectName, "@V3_SMS_Unit", vbTextCompare) Then
    
        If obj.Properties("OutputValue").DynamicStateType = hmiDynamicStateTypeVariableDirect Then
            Set objVar = obj.Properties("OutputValue").Dynamic
                With objVar
                    varTrig = .VarName
                End With
                
                If InStr(1, varTrig, "kN", vbTextCompare) > 0 Then
                    obj.OutputValue = "[kN]"
                ElseIf InStr(1, varTrig, "kg", vbTextCompare) > 0 Then
                    obj.OutputValue = "[kg]"
                ElseIf InStr(1, varTrig, "bar", vbTextCompare) > 0 Then
                    obj.OutputValue = "[bar]"
                ElseIf InStr(1, varTrig, "perc", vbTextCompare) > 0 Then
                    obj.OutputValue = "[%]"
                ElseIf InStr(1, varTrig, "mm", vbTextCompare) > 0 Then
                    obj.OutputValue = "[mm]"
                ElseIf InStr(1, varTrig, "m/s", vbTextCompare) > 0 Then
                    obj.OutputValue = "[m/s]"
                ElseIf InStr(1, varTrig, "rpm", vbTextCompare) > 0 Then
                    obj.OutputValue = "[rpm]"
                End If
                
                dynSourceCode = "'***************************************************************************************************" & vbCrLf & _
                                  "'Unit Display" & vbCrLf & _
                                  "   Item.OutputValue = ""["" & HMIRuntime.Tags(""" & varTrig & """).Read & ""]""" & vbCrLf & _
                                  "'***************************************************************************************************"
            Set objVBS = obj.Properties("OutputValue").CreateDynamic(hmiDynamicCreationTypeVBScript, dynSourceCode)
                With objVBS
                    .Trigger.VariableTriggers.Add varTrig, hmiVariableCycleType_500ms
                End With
        ElseIf obj.Properties("OutputValue").DynamicStateType = hmiDynamicStateTypeScript Then
            Set objVBS = obj.Properties("OutputValue").Dynamic
                With objVBS
                    If InStr(1, .sourceCode, "_kN_", vbTextCompare) > 0 Then
                        obj.InputValue = "[kN]"
                    ElseIf InStr(1, .sourceCode, "_bar_", vbTextCompare) > 0 Then
                        obj.InputValue = "[bar]"
                    ElseIf InStr(1, .sourceCode, "_kg_", vbTextCompare) > 0 Then
                        obj.InputValue = "[kg]"
                    ElseIf InStr(1, .sourceCode, "_perc_", vbTextCompare) > 0 Then
                        obj.InputValue = "[%]"
                    ElseIf InStr(1, .sourceCode, "_mm_", vbTextCompare) > 0 Then
                        obj.InputValue = "[mm]"
                    ElseIf InStr(1, .sourceCode, "_m/s_", vbTextCompare) > 0 Then
                        obj.InputValue = "[m/s]"
                    ElseIf InStr(1, .sourceCode, "_rpm_", vbTextCompare) > 0 Then
                        obj.InputValue = "[rpm]"
                    End If
                End With
        End If
    
    End If
    
Next obj

End Sub
