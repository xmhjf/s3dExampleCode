VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpsideInnerWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMarkingProfile
' Module: UpsideInner
'
' Description:   Determines the marking side for the mfg profile if it is inner.
'                Any web left surface whose nomal is pointing to the centre of the hull is Inner side.
'                Note: If a stiffener is not in standard directions(X,Y,Z), this might be helpful to identify the desired face
'
'  History:
'       Suma Mallena         May 02. 2008   created
'
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileUpsideRule

Const MODULE = "MfgMarkingProfile"

Private Function IJDMfgProfileUpsideRule_GetUpside(ByVal oStructProfilePart As Object) As Long
    Const METHOD = "UpsideInner: IJDMfgProfileUpsideRule_GetUpside"
    
    On Error GoTo ErrorHandler
    
    IJDMfgProfileUpsideRule_GetUpside = JXSEC_UNKNOWN
    
    If Not TypeOf oStructProfilePart Is IJProfilePart Then 'Member type
        Exit Function
    End If
    
    Dim oGeomUtils As IJTopologyLocate
    Set oGeomUtils = New TopologyLocate
        
    Dim oStructConnectable  As IJStructConnectable
    Dim oEnumPorts          As IJElements
    Set oEnumPorts = New JObjectCollection
    
    Set oStructConnectable = oStructProfilePart
    oEnumPorts.AddElements oStructConnectable.enumAllPorts

    Dim j As Integer
    Dim oPort As IJPort
    Dim oStructPort As IJStructPort
    Dim oCenter As IJDPosition
    Dim oNormal As IJDVector
    
    For j = 1 To oEnumPorts.Count
        Set oPort = oEnumPorts.Item(j)
        Set oStructPort = oPort
        If oStructPort.OperatorID = CLng(JXSEC_WEB_LEFT) Then

            oGeomUtils.FindApproxCenterAndNormal oPort.Geometry, oCenter, oNormal
            
             If oCenter.y < 0 Then
                If oNormal.y > 0 Then
                    IJDMfgProfileUpsideRule_GetUpside = JXSEC_WEB_LEFT
                   Exit For
                End If
            ElseIf oCenter.y > 0 Then
                If oNormal.y < 0 Then
                    IJDMfgProfileUpsideRule_GetUpside = JXSEC_WEB_LEFT
                    Exit For
                End If
            End If
            
         ElseIf oStructPort.OperatorID = CLng(JXSEC_WEB_RIGHT) Then

            oGeomUtils.FindApproxCenterAndNormal oPort.Geometry, oCenter, oNormal
            
            If oCenter.y < 0 Then
                If oNormal.y > 0 Then
                    IJDMfgProfileUpsideRule_GetUpside = JXSEC_WEB_RIGHT
                    Exit For
                End If
            ElseIf oCenter.y > 0 Then
                If oNormal.y < 0 Then
                    IJDMfgProfileUpsideRule_GetUpside = JXSEC_WEB_RIGHT
                    Exit For
                End If
            End If
        End If
        
    Next j
    
CleanUp:
    Set oGeomUtils = Nothing
    Set oStructConnectable = Nothing
    Set oEnumPorts = Nothing
    Set oPort = Nothing
    Set oStructPort = Nothing
    Set oCenter = Nothing
    Set oNormal = Nothing
    
    Exit Function

ErrorHandler:

    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4022, , "RULES")
    GoTo CleanUp
    
End Function
