VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImperialPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------------------'
'    Copyright (C) 2004 Intergraph Corporation. All rights reserved.
'
'
'Abstract
'    'Imperial Position' Grids Naming Rule.
'
'Notes
'    The Naming Rule format is:
'
'    Elevation Plane:
'    EL [Position in Imperial Units] [(Short Description of Elevation Type Code List)]. (ie: EL 28' 5" (BOBP))
'    X-Grid Plane:
'    E [Position in Imperial Units]. (ie: E 20' 0")
'    Y-Grid Plane:
'    N [Position in Imperial Units]. (ie: N 18' 5")
'    If Planes is on Ship Coordinate System then EL, E and N prefix should be replaced by X, Y and Z.
'
'History
'    Eric Faivre         07/21/04                Creation.
'    Eric Faivre         01/04/05                Support of Radial Plane and Cylinder.
'    Eric Faivre         08/08/05                Fix TR-CP·81252: Change naming rule for "Imperial Position" Naming option.
'                                                Use 'ft' and 'in' symbol instead of ' and '' which cause SQL queries to fail.
'--------------------------------------------------------------------------------------------'

Option Explicit

Implements IJNameRule

Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IMSICDPInterfacesLib.IJElements, ByVal pActiveEntity As Object)
    On Error GoTo ErrorHandler
    
    Dim oPlaneData As ISPGGridData
    Set oPlaneData = pEntity
    
    Dim oUOMServices As UomVBInterface
    Set oUOMServices = New UomVBInterface
    
    Dim imperialPosition As String
    If oPlaneData.Axis <> R Then
        oUOMServices.FormatUnit UNIT_DISTANCE, oPlaneData.Position, imperialPosition, , DISTANCE_FOOT, DISTANCE_INCH
    Else
        oUOMServices.FormatUnit UNIT_ANGLE, oPlaneData.Position, imperialPosition
    End If
    
    Dim oNavigate As ISPGNavigate
    Set oNavigate = pEntity
    
    Dim oCS As ISPGCoordinateSystemProperties
    oNavigate.GetParent oCS
    
    Dim CSType As CoordinateSystemType
    CSType = oCS.Type
    
    Dim planeName As String
    
    Select Case oPlaneData.Axis
        Case AxisType.x
            planeName = IIf(CSType = Grids, "E", "X") & " " & imperialPosition
        Case AxisType.y
            planeName = IIf(CSType = Grids, "N", "Y") & " " & imperialPosition
        Case AxisType.z
            planeName = IIf(CSType = Grids, "EL", "Z") & " " & imperialPosition & " (" & GetPlaneType(pEntity) & ")"
        Case AxisType.R
            planeName = "R" & " " & imperialPosition
        Case AxisType.C
            planeName = "C" & " " & imperialPosition
    End Select
    
    Dim oNamedItem As IJNamedItem
    Set oNamedItem = pEntity
    
    oNamedItem.Name = planeName
    
    Exit Sub
    
ErrorHandler:
    mg_oErrors.Add Err.Number, "GSNamingRules::ImperialPosition::IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
End Sub

Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IMSICDPInterfacesLib.IJElements
    Set IJNameRule_GetNamingParents = Nothing
End Function

Private Sub Class_Initialize()
    Set mg_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set mg_oErrors = Nothing
End Sub

