Attribute VB_Name = "PinjigBalanceHlpr"
'*******************************************************************
'  Copyright (C) 2011 Intergraph.  All rights reserved.
'
'  Project:
'
'  Abstract:    PinjigBalanceHlpr.bas
'
'  History:
'     Suma Mallena     Nov 08, 2011  Creation
'******************************************************************

Option Explicit
Private Const MODULE As String = "PinjigBalanceHlpr::"


Public Function GetPointsBasedOnRefLine(ByVal oPinJig As IJPinJig, ByVal oInputPointsColl As Collection) As Collection
    Const METHOD = "GetPointsBasedOnRefLine"
    On Error GoTo ErrorHandler

    Set GetPointsBasedOnRefLine = New Collection
    Dim iCnt                 As Integer
    Dim oSurfaceToProject    As IJSurfaceBody
    Dim oStartPos                       As IJDPosition
    Dim oEndPos                         As IJDPosition
    
    Set oSurfaceToProject = GetSurfaceToProject(oPinJig)

    'User Selection: Picked One Reference Line and Two Points
    For iCnt = 1 To oInputPointsColl.Count
        If TypeOf oInputPointsColl.Item(iCnt) Is IJDPosition Then
            GetPointsBasedOnRefLine.Add oInputPointsColl.Item(iCnt)
        Else
            If TypeOf oInputPointsColl.Item(iCnt) Is IJSeam Then
                'Get end points of Reference Line
                GetEndPointsOfEntity oInputPointsColl.Item(iCnt), oSurfaceToProject, oStartPos, oEndPos
            ElseIf TypeOf oInputPointsColl.Item(iCnt) Is IJRefCurveOnSurface Then
            
                Dim RefCurve As IJRefCurveOnSurface
                Set RefCurve = oInputPointsColl.Item(iCnt)
                
                GetEndPointsOfEntity RefCurve, oSurfaceToProject, oStartPos, oEndPos
                                   
            ElseIf TypeOf oInputPointsColl.Item(iCnt) Is IJMfgMarkingLines_AE Then
                Dim oMarkLine As IJMfgMarkingLines_AE
                Set oMarkLine = oInputPointsColl.Item(iCnt)
                
                Dim oMarkCSColl     As IJDObjectCollection
                Set oMarkCSColl = oMarkLine.GeometryAsComplexStrings
                
                Dim oMarkCS As IJComplexString
                Dim oMarkWB As IJWireBody
                Dim oMfgMGHelper        As New MfgMGHelper
                For Each oMarkCS In oMarkCSColl
                    oMfgMGHelper.ComplexStringToWireBody oMarkCS, oMarkWB
                    GetEndPointsOfEntity oMarkWB, oSurfaceToProject, oStartPos, oEndPos
                Next oMarkCS
                                    
            End If
            
            GetPointsBasedOnRefLine.Add oStartPos
            GetPointsBasedOnRefLine.Add oEndPos
        
        End If
    Next


    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5030, , "RULES")
End Function

Public Function GetSurfaceToProject(ByVal oPinJig As IJPinJig) As IJSurfaceBody
    Const METHOD = "GetSurfaceToProject"
    On Error GoTo ErrorHandler
    
    Dim oSurfaceUtil As IJMfgUtilSurface
    Set oSurfaceUtil = New MfgUtilSurface

    Dim oPlateColl As IJElements
    Set oPlateColl = oPinJig.SupportedPlates
    
    Dim RootX As Double, RootY As Double, RootZ As Double
    Dim NormX As Double, NormY As Double, NormZ As Double
    oPinJig.GetBasePlane NormX, NormY, NormZ, RootX, RootY, RootZ

    Dim oPlane As IJPlane
    Set oPlane = New Plane3d

    oPlane.DefineByPointNormal RootX, RootY, RootZ, NormX, NormY, NormZ
           
    Set GetSurfaceToProject = _
       oSurfaceUtil.GenSurfFromPlateSidesFacingPlane(oPlateColl, oPlane)

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
End Function

Public Sub GetEntitiesAndEndPoints(ByVal oPinJig As IJPinJig, ByRef oEntities As Collection, ByRef oEndPoints As Collection)
    Const METHOD = "GetEntitiesAndEndPoints"
    On Error GoTo ErrorHandler
    
    Dim EndPointsColl       As Collection
    Dim oStartPos           As IJDPosition
    Dim oEndPos             As IJDPosition
    Dim oSurfaceToProject   As IJSurfaceBody
    Dim oMfgMGHelper        As New MfgMGHelper
    Dim oEntityWireBody     As IJWireBody
    Dim oEntityCS           As IJComplexString
    
    Set EndPointsColl = New Collection
    Set oEntities = New Collection
    Set oSurfaceToProject = GetSurfaceToProject(oPinJig)
    
    ''''''''''''''''SEAMS'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim oUniqueLogConns As IJElements
    Set oUniqueLogConns = GetConnectionsBetweenPinJigPlates(ConnectionLogical, oPinJig)
    
    Dim oSeamColl As IJElements
    Dim Dummy1() As Long
    Dim Dummy2() As Long
    Dim Dummy3() As Long
    Dim Dummy4() As Long
    Dim i As Integer
    
    If oUniqueLogConns.Count > 0 Then
        Set oSeamColl = GetOperatorsFromConnections(oUniqueLogConns, Dummy1, Dummy2, Dummy3, Dummy4)
    Else
        Set oSeamColl = New JObjectCollection
    End If
        
    ' Now project the seams on the remarking surface.
    For i = 1 To oSeamColl.Count
    
        oEntities.Add oSeamColl.Item(i)
        GetEndPointsOfEntity oSeamColl.Item(i), oSurfaceToProject, oStartPos, oEndPos
        
        'Need not return seam end points as they would result in duplication.(Already part of plate outer corners)
        'EndPointsColl.Add oStartPos
        'EndPointsColl.Add oEndPos
  
    Next i
    
    ''''''''''''''''KNUCKLES / REFERENCE CURVES'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim oRefCurveColl As Collection
    Set oRefCurveColl = GetRefCurveGeomDataFromPinJigSupportedPlates(oPinJig)
        
    Dim oRefCurveData As IJRefCurveData
    Dim oRefCurve As IJRefCurveOnSurface
    For Each oRefCurveData In oRefCurveColl
        Set oRefCurve = oRefCurveData.ParentReferenceCurve
    
        oEntities.Add oRefCurve
            
        GetEndPointsOfEntity oRefCurve, oSurfaceToProject, oStartPos, oEndPos
        
        EndPointsColl.Add oStartPos
        EndPointsColl.Add oEndPos

    Next oRefCurveData
    
    ''''''''''''''''CENTER LINE'''''''''''''''''''''''''''''''''''''''''''''
    
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper
    
    Dim oXZPlane As Plane3d 'Y=0 plane
    oGeomHelper.MakeTransientPlane 0, 0, 0, 0, 1, 0, oXZPlane
    
    Dim oCenterLineCurve As IUnknown
        
    On Error Resume Next
    oGeomHelper.IntersectSurfaceWithPlane oSurfaceToProject, oXZPlane, oCenterLineCurve, oStartPos, oEndPos
    On Error GoTo ErrorHandler
    
    If Not oCenterLineCurve Is Nothing Then
        EndPointsColl.Add oStartPos
        EndPointsColl.Add oEndPos
    End If

    ''''''''''''''''USER DEFINED MARKING LINES'''''''''''''''''''''''''''''''''''''''''''''
    
    Dim oUserMarkColl As IJElements
    ' Return only user marks of type "STRMFG_PINJIG_MARK"
    Const OnlyThisPinJig As Boolean = False
    Set oUserMarkColl = oPinJig.MarkingLinesOnSupportedPlates(OnlyThisPinJig, STRMFG_PINJIG_MARK, PinJigRemarkingSide)
    
    Dim oMarkLine As IJMfgMarkingLines_AE
    For Each oMarkLine In oUserMarkColl
    
        oEntities.Add oMarkLine
        
        Dim oMarkCSColl     As IJDObjectCollection
        Set oMarkCSColl = oMarkLine.GeometryAsComplexStrings
        
        Dim oMarkCS As IJComplexString
        Dim oMarkWB As IJWireBody
        For Each oMarkCS In oMarkCSColl
        
            oMfgMGHelper.ComplexStringToWireBody oMarkCS, oMarkWB
            GetEndPointsOfEntity oMarkWB, oSurfaceToProject, oStartPos, oEndPos
            
            EndPointsColl.Add oStartPos
            EndPointsColl.Add oEndPos
            
        Next oMarkCS
        
    Next oMarkLine

    Set oEndPoints = EndPointsColl
    
CleanUp:
    Set EndPointsColl = Nothing
    Set oSeamColl = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
    GoTo CleanUp
End Sub

Public Sub GetEndPointsOfEntity(ByVal oEntityWireBody As IJWireBody, ByVal oSurface As IJSurfaceBody, ByRef oStartPos As IJDPosition, ByRef oEndPos As IJDPosition)
    Const METHOD = "GetEndPointsOfEntity"
    On Error GoTo ErrorHandler
    
    Dim oMfgMGHelper        As New MfgMGHelper
    Dim oEntityCS           As IJComplexString
    Dim oCSColl             As IJElements
    Dim oCS                 As IJComplexString
    Dim oWB                 As IJWireBody

    oMfgMGHelper.WireBodyToComplexString oEntityWireBody, oEntityCS
    oMfgMGHelper.ProjectCSToSurface oEntityCS, oSurface, Nothing, oCSColl

    For Each oCS In oCSColl
        oMfgMGHelper.ComplexStringToWireBody oCS, oWB
        oWB.GetEndPoints oStartPos, oEndPos
    Next oCS

    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")

End Sub

