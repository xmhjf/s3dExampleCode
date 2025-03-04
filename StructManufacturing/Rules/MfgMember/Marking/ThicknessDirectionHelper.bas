Attribute VB_Name = "ThicknessDirectionHelper"
'*******************************************************************
'  Copyright (C) 2007 Intergraph Corporation.  All rights reserved.
'
'  Project: MfgPlateMarking, MfgMemberMarking
'
'  Abstract:    Helpers for the MfgPlateMarking and MfgMemberMarking rules.
'               Within this module we will generate thickness direction
'               vector for the marking lines depending on the landing curve
'               and the Part geometry
'
'  History:
'       K. Kamph        August 20. 2007     created
'
'******************************************************************
Option Explicit

Private Const MODULE = "MfgPlateMarking.ThicknessDirectionHelper"

Public Function GetThicknessDirectionVector(oCS As IJComplexString, oSDConWrapper As Object, sMoldedSide As String) As IJDVector
    Const METHOD = "GetThicknessDirectionVector"
    
    On Error GoTo ErrorHandler
    Dim oWireBody As IJWireBody
    Dim oStart1 As IJDPosition, oEnd1 As IJDPosition
    Dim oVectorThickness As IJDVector
    
    'we must adjust the returned TD vector for each of the new seperate complex strings
    Dim oSurfaceBodyCon1 As IJSurfaceBody
    If TypeOf oSDConWrapper.object Is IJPlatePart Then
        If sMoldedSide = "Base" Then
            Set oSurfaceBodyCon1 = oSDConWrapper.BasePort(BPT_Offset).Geometry
        ElseIf sMoldedSide = "Offset" Then
            Set oSurfaceBodyCon1 = oSDConWrapper.BasePort(BPT_Base).Geometry
        End If
    ElseIf TypeOf oSDConWrapper.object Is IJProfilePart Then
        If sMoldedSide = "Base" Then
            Set oSurfaceBodyCon1 = oSDConWrapper.BasePort(BPT_Offset).Geometry
        ElseIf sMoldedSide = "Offset" Then
            Set oSurfaceBodyCon1 = oSDConWrapper.BasePort(BPT_Base).Geometry
        End If
    End If
    
    Set oWireBody = m_oMfgRuleHelper.ComplexStringToWireBody(oCS)
                    
    oWireBody.GetEndPoints oStart1, oEnd1
    
    Set oStart1 = m_oMfgRuleHelper.ProjectPointOnSurface(oEnd1, oSurfaceBodyCon1, oVectorThickness)

    oVectorThickness.Set oStart1.x - oEnd1.x, oStart1.y - oEnd1.y, oStart1.z - oEnd1.z

    Set GetThicknessDirectionVector = oVectorThickness
CleanUp:
    Set oSurfaceBodyCon1 = Nothing
    Set oWireBody = Nothing
    Set oVectorThickness = Nothing
    Set oStart1 = Nothing
    Set oEnd1 = Nothing
Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1001, , "RULES")
    GoTo CleanUp
End Function
