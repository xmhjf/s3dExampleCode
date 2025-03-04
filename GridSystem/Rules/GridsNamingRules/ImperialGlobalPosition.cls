VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImperialGlobalPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------------------'
'    Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'
'Abstract
'    'Imperial Global Position' Grids Naming Rule.
'
'History
'    Eric Faivre         01/30/07                Creation.
'    Eric Faivre         03/15/07                Fix TR-CP·116601: Global Naming Rule of Grids fails with rotated Coordinate System.
'--------------------------------------------------------------------------------------------'

Option Explicit

Implements IJNameRule

Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IMSICDPInterfacesLib.IJElements, ByVal pActiveEntity As Object)
    On Error GoTo ErrorHandler
    
    Dim oNavigate As ISPGNavigate
    Set oNavigate = pEntity

    Dim oCSProperties As ISPGCoordinateSystemProperties
    oNavigate.GetParent oCSProperties
    
    Dim oPlaneData As ISPGGridData
    Set oPlaneData = pEntity
    
    Dim CSType As CoordinateSystemType
    CSType = oCSProperties.Type

    Dim oCS As IJDCoordinateSystem
    Dim Axis As AxisType
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
    
    Set oCS = oCSProperties
    Axis = oPlaneData.Axis

    If Axis <> R Then
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
                position = GetImperialFormat(csOrigin + sign * oPlaneData.position)
            Else
                position = "[ " + GetImperialFormat(csOrigin) + " ; " + GetImperialFormat(oPlaneData.position) + " ]"
            End If
        Else
            prefix = "R"
            position = GetImperialFormat(oPlaneData.position)
        End If
    Else
        prefix = "C"
        position = GetImperialFormat(oPlaneData.position, True)
    End If
    
    Dim planeName As String
    
    Select Case Axis
        Case AxisType.x
            planeName = prefix & " " & position
        Case AxisType.y
            planeName = prefix & " " & position
        Case AxisType.z
            planeName = prefix & " " & position & " (" & GetPlaneType(pEntity) & ")"
        Case AxisType.R
            planeName = prefix & " " & position
        Case AxisType.C
            planeName = prefix & " " & position
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

Private Function GetImperialFormat(position As Double, Optional angleFormat As Boolean = False) As String
    Dim oUOMServices As UomVBInterface
    Set oUOMServices = New UomVBInterface

    Dim imperialFormat As String
    
    If angleFormat = False Then
        oUOMServices.FormatUnit UNIT_DISTANCE, position, imperialFormat, , DISTANCE_FOOT, DISTANCE_INCH
    Else
        oUOMServices.FormatUnit UNIT_ANGLE, position, imperialFormat
    End If
    
    GetImperialFormat = imperialFormat
End Function
