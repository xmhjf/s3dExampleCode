Attribute VB_Name = "CommonFuncs"
'*******************************************************************
'  Copyright (C) 2016 Intergraph.  All rights reserved.
'
'  Project:
'
'  Abstract:    ManagedExportFunctions.bas
'
'  History:
'
'
'******************************************************************
Option Explicit

Public Function IsPanelObject(pDispObject As Object) As Boolean
Const sMETHOD As String = "IsPanelObject"
On Error GoTo ErrorHandler

    IsPanelObject = False

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = pDispObject
    
    Dim oMfgPlatePart As IJMfgPlatePart
    If oPlateWrapper.PlateHasMfgPart(oMfgPlatePart) Then
        If oMfgPlatePart.PanelMode = True Then
            IsPanelObject = True
        End If
    End If
    
    Set oMfgPlatePart = Nothing
    Set oPlateWrapper = Nothing
    
    Exit Function
ErrorHandler:
End Function


Public Function IsMemberOfPanel(pDispObject As Object) As Boolean
Const sMETHOD As String = "IsMemberOfPanel"
On Error GoTo ErrorHandler

    IsMemberOfPanel = False

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = pDispObject
    
    Dim oPanelCollection As Collection
    Set oPanelCollection = oPlateWrapper.GetPanelCollection
    
    If Not oPanelCollection Is Nothing Then
        If oPanelCollection.Count > 0 Then
            IsMemberOfPanel = True
        End If
    End If
    
    Set oPanelCollection = Nothing
    Set oPlateWrapper = Nothing
    
    Exit Function
ErrorHandler:
End Function

Public Function GetName(pDispObject As Object) As String
Const sMETHOD As String = "GetName"
On Error GoTo ErrorHandler

    GetName = ""

    Dim oNamedItem As IJNamedItem
    Set oNamedItem = pDispObject
   
    If Not oNamedItem Is Nothing Then
        GetName = oNamedItem.Name
    End If
   
   Set oNamedItem = Nothing
    
    Exit Function
ErrorHandler:
End Function

Public Function SanitizeFileName(ByVal strInput As String) As String
Const sMETHOD As String = "SanitizeFileName"
On Error GoTo ErrorHandler

    SanitizeFileName = "SanitizeFileName"
    '<>:\"/\\|?*
    
    Dim tempString As String
    tempString = strInput
    
    If InStr(tempString, "<") > 0 Then
        tempString = Replace(tempString, "<", "_")
    End If
    
    If InStr(tempString, ">") > 0 Then
        tempString = Replace(tempString, ">", "_")
    End If
    
    If InStr(tempString, ":") > 0 Then
        tempString = Replace(tempString, ":", "_")
    End If
    
    If InStr(tempString, "\") > 0 Then
        tempString = Replace(tempString, "\", "_")
    End If
    
    If InStr(tempString, "/") > 0 Then
        tempString = Replace(tempString, "/", "_")
    End If
    
    If InStr(tempString, "//") > 0 Then
        tempString = Replace(tempString, "//", "_")
    End If
    
    If InStr(tempString, "|") > 0 Then
        tempString = Replace(tempString, "|", "_")
    End If
    
    If InStr(tempString, "?") > 0 Then
        tempString = Replace(tempString, "?", "_")
    End If
    
    If InStr(tempString, "*") > 0 Then
        tempString = Replace(tempString, "*", "_")
    End If
    
    SanitizeFileName = tempString
    
    
    Exit Function
ErrorHandler:

End Function



