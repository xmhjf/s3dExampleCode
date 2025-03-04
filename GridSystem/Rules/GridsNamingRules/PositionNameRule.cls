VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'
'  Abstract:
'   The file contains  implementation of a naming rule for Elevation
'   Planes
'
'  Author:
'   Celine Faure    10/22/02    Creation
'   Celine Faure    04/03/03    Change Prefix for Ship Naming Rule
'   Eric Faivre     06/06/03    Added meter unit.
'   Eric Faivre     01/04/05    Support of Radial Plane and Cylinder.
'   Eric Faivre     05/17/05    Fix TR-CP·71736: Position name rule shows angle in radians.
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
    Dim oCS As ISPGCoordinateSystemProperties
    Dim CSType As CoordinateSystemType

    If pEntity Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
        Exit Sub
    End If

    Set oChildNamedItem = pEntity
    Set oGridData = pEntity
    Set oNavigate = pEntity

    oNavigate.GetParent oCS
    Axis = oGridData.Axis
    CSType = oCS.Type
    
    Select Case Axis
        Case x
            oChildNamedItem.Name = IIf(CSType = Ship, "F", "E") + " " + CStr(FormatNumber(oGridData.Position, 3)) + " m"
        Case y
            oChildNamedItem.Name = IIf(CSType = Ship, "L", "N") + " " + CStr(FormatNumber(oGridData.Position, 3)) + " m"
        Case z
            oChildNamedItem.Name = IIf(CSType = Ship, "D", "El") + " " + CStr(FormatNumber(oGridData.Position, 3)) + " m"
        Case R
            oChildNamedItem.Name = "R" + " " + CStr(FormatNumber(oGridData.Position * 180 / 3.14159265358979, 3)) + " deg"
        Case C
            oChildNamedItem.Name = "C" + " " + CStr(FormatNumber(oGridData.Position, 3)) + " m"
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


