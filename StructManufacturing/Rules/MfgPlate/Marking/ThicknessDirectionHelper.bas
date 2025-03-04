Attribute VB_Name = "ThicknessDirectionHelper"
'*******************************************************************
'  Copyright (C) 2007 Intergraph Corporation.  All rights reserved.
'
'  Project: MfgPlateMarking, MfgProfileMarking
'
'  Abstract:    Helpers for the MfgPlateMarking and MfgProfileMarking rules.
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

Public Function GetThicknessDirectionVector(oCS As IJComplexString, oSDConWrapper As Object, sMoldedSide As String, _
                                            Optional oSurface As IJSurface = Nothing, _
                                            Optional oVector As IJDVector = Nothing) As IJDVector
    Const METHOD = "GetThicknessDirectionVector"
    
    On Error GoTo ErrorHandler
    Dim oWireBody As IJWireBody
    Dim oStart1 As IJDPosition, oEnd1 As IJDPosition
    Dim oVectorThickness As IJDVector
    
    'we must adjust the returned TD vector for each of the new seperate complex strings
    Dim oSurfaceBodyCon1 As IJSurfaceBody
    If TypeOf oSDConWrapper.object Is IJPlatePart Then
        Dim oPartSupp As IJPartSupport
        Dim oPlatePartSupp As IJPlatePartSupport
        Set oPartSupp = New PlatePartSupport
        Set oPlatePartSupp = oPartSupp
        Set oPartSupp.Part = oSDConWrapper.object

        If sMoldedSide = "Base" Then
            oPlatePartSupp.GetSurfaceWithoutFeatures PlateOffsetSide, oSurfaceBodyCon1
        ElseIf sMoldedSide = "Offset" Then
            oPlatePartSupp.GetSurfaceWithoutFeatures PlateBaseSide, oSurfaceBodyCon1
        End If
        Set oPlatePartSupp = Nothing
        Set oPartSupp = Nothing
    ElseIf TypeOf oSDConWrapper.object Is IJProfilePart Then
        If sMoldedSide = "Base" Then
            Set oSurfaceBodyCon1 = oSDConWrapper.SubPortBeforeTrim(JXSEC_WEB_RIGHT).Geometry
        ElseIf sMoldedSide = "Offset" Then
            Set oSurfaceBodyCon1 = oSDConWrapper.SubPortBeforeTrim(JXSEC_WEB_LEFT).Geometry
        End If
    ElseIf TypeOf oSDConWrapper.object Is ISPSMemberPartPrismatic Then
        If sMoldedSide = "Base" Then
            Set oSurfaceBodyCon1 = oSDConWrapper.SubPort(JXSEC_WEB_RIGHT).Geometry
        ElseIf sMoldedSide = "Offset" Then
            Set oSurfaceBodyCon1 = oSDConWrapper.SubPort(JXSEC_WEB_LEFT).Geometry
        End If
    End If
    
    If Not oSurface Is Nothing Then
        'Check distance between input surface and oSurfaceBodyCon1
        ' ... if dist > 0.001 then use oVector else continue to following vector calculations
        Dim dMinDist As Double
        Dim pClosestPos1 As IJDPosition, pClosestPos2 As IJDPosition
        Dim oSurfModelBody As IJDModelBody
        
        Set oSurfModelBody = oSurfaceBodyCon1
        
        oSurfModelBody.GetMinimumDistance oSurface, pClosestPos1, pClosestPos2, dMinDist
        If dMinDist > 0.001 Then
            If Not oVector Is Nothing Then
                Set GetThicknessDirectionVector = oVector
                GoTo CleanUp
            End If
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
    GoTo CleanUp
End Function
