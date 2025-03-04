VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalPositionNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2007, Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'
'  Abstract:
'   The file contains implementation of the Global Position Naming Rule.
'
'  Author:
'   Eric Faivre     01/30/07    Creation.
'   Eric Faivre     03/15/07    Fix TR-CP·116601: Global Naming Rule of Grids fails with rotated Coordinate System.
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "GSNamingRules: "

Private Sub Class_Initialize()
    Set mg_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set mg_oErrors = Nothing
End Sub

Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IMSICDPInterfacesLib.IJElements, ByVal pActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo label

    Dim oChildNamedItem As IJNamedItem
    Dim oGridData As ISPGGridData
    Dim oNavigate As ISPGNavigate
    Dim Axis As AxisType
    Dim oCSProperties As ISPGCoordinateSystemProperties
    Dim CSType As CoordinateSystemType
    Dim oCS As IJDCoordinateSystem
    Dim oAxisCS As IJDVector
    Dim oAxisGlobal As IJDVector
    Dim oOriginCS As IJDPosition
    Dim i As Integer
    Dim axisGlobal As AxisType
    Dim maxDot As Double
    Dim csOrigin As Double
    Dim prefix As String
    Dim position As String
    Dim sign As Integer

    If pEntity Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
        Exit Sub
    End If

    Set oChildNamedItem = pEntity
    Set oGridData = pEntity
    Set oNavigate = pEntity

    oNavigate.GetParent oCSProperties
    Axis = oGridData.Axis
    CSType = oCSProperties.Type
    Set oCS = oCSProperties
    
    If Axis = x Or Axis = y Or Axis = z Then
        Set oAxisCS = oCS.UnitVector(Switch(Axis = x, csDirectionX, Axis = y, csDirectionY, Axis = z, csDirectionZ))
        
        maxDot = 0
    
        For i = 1 To 3
            Set oAxisGlobal = New DVector
            oAxisGlobal.Set Switch(i = 1, 1, i = 2, 0, i = 3, 0), _
                            Switch(i = 1, 0, i = 2, 1, i = 3, 0), _
                            Switch(i = 1, 0, i = 2, 0, i = 3, 1)
                            
            If Abs(oAxisCS.Dot(oAxisGlobal)) > maxDot Then
                maxDot = Abs(oAxisCS.Dot(oAxisGlobal))
                axisGlobal = Switch(i = 1, x, i = 2, y, i = 3, z)
                sign = IIf(oAxisCS.Dot(oAxisGlobal) > 0, 1, -1)
            End If
        Next
        
        Set oOriginCS = oCS.Origin
        
        Select Case axisGlobal
            Case x
                prefix = IIf(CSType = Ship, "F", "E")
                csOrigin = oOriginCS.x
            Case y
                prefix = IIf(CSType = Ship, "L", "N")
                csOrigin = oOriginCS.y
            Case z
                prefix = IIf(CSType = Ship, "D", "El")
                csOrigin = oOriginCS.z
        End Select
        
        If Abs(maxDot - 1) < 0.000001 Then
            position = CStr(FormatNumber(csOrigin + sign * oGridData.position, 3)) + " m"
        Else
            position = "[ " + CStr(FormatNumber(csOrigin, 3)) + " m ; " + CStr(FormatNumber(oGridData.position, 3)) + " m ]"
        End If
    End If
    
    Select Case Axis
        Case x
            oChildNamedItem.Name = prefix + " " + position
        Case y
            oChildNamedItem.Name = prefix + " " + position
        Case z
            oChildNamedItem.Name = prefix + " " + position
        Case R
            oChildNamedItem.Name = "R" + " " + CStr(FormatNumber(oGridData.position * 180 / 3.14159265358979, 3)) + " deg"
        Case C
            oChildNamedItem.Name = "C" + " " + CStr(FormatNumber(oGridData.position, 3)) + " m"
    End Select
   
Exit Sub

label:
    ' log the error in middle tier and propagate the error code to the caller
    mg_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
    
End Sub

Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Set IJNameRule_GetNamingParents = Nothing
End Function

