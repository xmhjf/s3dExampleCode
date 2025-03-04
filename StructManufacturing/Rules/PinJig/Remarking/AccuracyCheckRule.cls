VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccuracyCheckRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2011, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigRule
' Module: AccuracyCheckRule
'
' Description:
'
' Author: Suma Mallena
'
'*******************************************************************************

Option Explicit
Private Const MODULE As String = "MfgPinJigRule.AccuracyCheckRule"

Implements IJDPinJigAccuracyRule

'---------------------------------------------------------------------------------------
' Procedure : IJDPinJigAccuracyRule_GetAccuracyObjectsForPurpose
' Purpose   : Return accuracy check points corresponding to the "AttributeName"
'---------------------------------------------------------------------------------------
Private Function IJDPinJigAccuracyRule_GetAccuracyObjectsForPurpose(ByVal PartialPinJig As Object, ByVal AttributeName As String, ByVal AttributeValue As Variant) As IJMfgGeomCol3d
 Const METHOD = "IJDPinJigAccuracyRule_GetAccuracyObjectsForPurpose"
    On Error GoTo ErrorHandler
    
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig
    
    Dim ReturnColl As IJMfgGeomCol3d
    Set ReturnColl = CreateMfgGeomCol3dObject
    
    Dim oPinJigReport               As IJDPinJigReport
    Dim oJigOutput                  As IJJigOutput
    Dim oProjectedData              As IJJigProjectedData
    Dim oCuspPointColl              As Collection
    Dim oParentRemarkColl           As Collection
    Dim iCounter                    As Integer
    Dim oNamedItem                  As IJNamedItem
    Dim eSubGeometryType            As StrMfgGeometryType
    Dim strAccuracyPointName        As String
    Dim oAccuracyPointColl          As IJElements
    Dim o3DCountourLine             As IJMfgGeom3d
    Dim o2DCountourLine             As IJMfgGeom3d
    Dim oReferenceCountourLine      As IJMfgGeom3d
    Dim oRemarkingLineColl          As IJElements
    Dim oJigPart3d                  As IJJigPart3D
    Dim oProcessData                As IJJigProcessData
    
    Set oPinJigReport = oPinJig
            
    Set oJigPart3d = oPinJig.GetJigPart3D
    Set oProcessData = oJigPart3d.GetJigProcessData
        
    Set oJigOutput = oPinJig.GetJigOutput
    Set oProjectedData = oJigOutput.GetJigProjectedData

    Dim oRemarkUtil As IJDJigRemarkLineUtil
    Set oRemarkUtil = oPinJigReport.GetEntity(RemarkingLineHandler)

    Select Case AttributeName
        Case "BendPoints"
        
            Set oRemarkingLineColl = oRemarkUtil.GetCollection(AllRemarkingLines, ProjectedJigRemarking)
            ' Get tha cusp points with cusp angle = 5 degrees
            oPinJigReport.GetCuspPointsCollection oRemarkingLineColl, 0.09, oCuspPointColl, oParentRemarkColl
            
            eSubGeometryType = STRMFG_PinJig_CuspIntxPt
            
            For iCounter = 1 To oCuspPointColl.Count
                'Set the Name
                Set oNamedItem = oParentRemarkColl.Item(iCounter)
                strAccuracyPointName = "Cusp Point on " & oNamedItem.Name

                oProjectedData.AddGeom3dAndSetParentRML oCuspPointColl.Item(iCounter), oParentRemarkColl.Item(iCounter), eSubGeometryType, strAccuracyPointName, ReturnColl
            Next
            
            Set IJDPinJigAccuracyRule_GetAccuracyObjectsForPurpose = ReturnColl
            Exit Function
            
        Case "AftBoundary"
           
            Set oRemarkingLineColl = oRemarkUtil.GetCollection(AftContourJigRemarking, JigSurfaceRemarking)
            If oRemarkingLineColl.Count = 0 Then Exit Function
            Set o3DCountourLine = oRemarkingLineColl.Item(1)
            
             Select Case AttributeValue
                 Case 1 'Along Girth From Lower
                    Set oRemarkingLineColl = oRemarkUtil.GetCollection(LowerContourJigRemarking, JigSurfaceRemarking)
                    If oRemarkingLineColl.Count = 0 Then Exit Function
                    
                    Set oReferenceCountourLine = oRemarkingLineColl.Item(1)
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAlongGirth(o3DCountourLine, 0.5, oReferenceCountourLine)
                    
                 Case 2 'Along Girth From Upper
                    Set oRemarkingLineColl = oRemarkUtil.GetCollection(UpperContourJigRemarking, JigSurfaceRemarking)
                    If oRemarkingLineColl.Count = 0 Then Exit Function
                    
                    Set oReferenceCountourLine = oRemarkingLineColl.Item(1)
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAlongGirth(o3DCountourLine, 0.5, oReferenceCountourLine)
                    
                 Case 3 'At Pin Lines
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAtPinGridLines(o3DCountourLine, False)
                    
                 Case 4 'At and Middle of Pin Lines
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAtPinGridLines(o3DCountourLine, True)
                    
             End Select
       
        Case "LowerBoundary"
            
            Set oRemarkingLineColl = oRemarkUtil.GetCollection(LowerContourJigRemarking, JigSurfaceRemarking)
            If oRemarkingLineColl.Count = 0 Then Exit Function
            Set o3DCountourLine = oRemarkingLineColl.Item(1)
            
             Select Case AttributeValue
                 Case 1 'Along Girth From Aft
                    Set oRemarkingLineColl = oRemarkUtil.GetCollection(AftContourJigRemarking, JigSurfaceRemarking)
                    If oRemarkingLineColl.Count = 0 Then Exit Function
                    
                    Set oReferenceCountourLine = oRemarkingLineColl.Item(1)
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAlongGirth(o3DCountourLine, 0.5, oReferenceCountourLine)
                    
                 Case 2 'Along Girth From Fore
                    Set oRemarkingLineColl = oRemarkUtil.GetCollection(ForeContourJigRemarking, JigSurfaceRemarking)
                    If oRemarkingLineColl.Count = 0 Then Exit Function
                    
                    Set oReferenceCountourLine = oRemarkingLineColl.Item(1)
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAlongGirth(o3DCountourLine, 0.5, oReferenceCountourLine)
                    
                 Case 3 'At Pin Lines
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAtPinGridLines(o3DCountourLine, False)
                    
                 Case 4 'At and Middle of Pin Lines
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAtPinGridLines(o3DCountourLine, True)
                    
             End Select
       
        Case "ForeBoundary"
            
            Set oRemarkingLineColl = oRemarkUtil.GetCollection(ForeContourJigRemarking, JigSurfaceRemarking)
            If oRemarkingLineColl.Count = 0 Then Exit Function
            Set o3DCountourLine = oRemarkingLineColl.Item(1)
            
             Select Case AttributeValue
                 Case 1 'Along Girth From Lower
                    Set oRemarkingLineColl = oRemarkUtil.GetCollection(LowerContourJigRemarking, JigSurfaceRemarking)
                    If oRemarkingLineColl.Count = 0 Then Exit Function
                    
                    Set oReferenceCountourLine = oRemarkingLineColl.Item(1)
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAlongGirth(o3DCountourLine, 0.5, oReferenceCountourLine)
                    
                 Case 2 'Along Girth From Upper
                    Set oRemarkingLineColl = oRemarkUtil.GetCollection(UpperContourJigRemarking, JigSurfaceRemarking)
                    If oRemarkingLineColl.Count = 0 Then Exit Function
                    
                    Set oReferenceCountourLine = oRemarkingLineColl.Item(1)
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAlongGirth(o3DCountourLine, 0.5, oReferenceCountourLine)
                    
                 Case 3 'At Pin Lines
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAtPinGridLines(o3DCountourLine, False)
                    
                 Case 4 'At and Middle of Pin Lines
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAtPinGridLines(o3DCountourLine, True)
                    
             End Select
       
        Case "UpperBoundary"
            
            Set oRemarkingLineColl = oRemarkUtil.GetCollection(UpperContourJigRemarking, JigSurfaceRemarking)
            If oRemarkingLineColl.Count = 0 Then Exit Function
            Set o3DCountourLine = oRemarkingLineColl.Item(1)
            
             Select Case AttributeValue
                 Case 1 'Along Girth From Aft
                    Set oRemarkingLineColl = oRemarkUtil.GetCollection(AftContourJigRemarking, JigSurfaceRemarking)
                    If oRemarkingLineColl.Count = 0 Then Exit Function
                    
                    Set oReferenceCountourLine = oRemarkingLineColl.Item(1)
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAlongGirth(o3DCountourLine, 0.5, oReferenceCountourLine)
                    
                 Case 2 'Along Girth From Fore
                    Set oRemarkingLineColl = oRemarkUtil.GetCollection(ForeContourJigRemarking, JigSurfaceRemarking)
                    If oRemarkingLineColl.Count = 0 Then Exit Function
                    
                    Set oReferenceCountourLine = oRemarkingLineColl.Item(1)
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAlongGirth(o3DCountourLine, 0.5, oReferenceCountourLine)
                    
                 Case 3 'At Pin Lines
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAtPinGridLines(o3DCountourLine, False)
                    
                 Case 4 'At and Middle of Pin Lines
                    Set oAccuracyPointColl = oPinJigReport.GetAccuracyCheckPtsAtPinGridLines(o3DCountourLine, True)
                    
             End Select
       
        Case "Custom"
        
            'Not Implemented


    End Select
    
    If oAccuracyPointColl Is Nothing Then Exit Function
    
    eSubGeometryType = STRMFG_PinJig_AccuracyPoint
    
    Dim o2DPosition As IJDPosition
    For iCounter = 1 To oAccuracyPointColl.Count
    
        'Get Position on Jig Floor
        Set o2DPosition = oProcessData.GetCorresponding2dAccuracyPoint(oAccuracyPointColl.Item(iCounter), o3DCountourLine)
        Set o2DCountourLine = oProcessData.GetCorresponding2dRemarkingLine(o3DCountourLine)

        ' Set the Name
        Set oNamedItem = o3DCountourLine
        strAccuracyPointName = "AP_" & iCounter & "_" & oNamedItem.Name
                               
        oProjectedData.AddGeom3dAndSetParentRML o2DPosition, o2DCountourLine, eSubGeometryType, strAccuracyPointName, ReturnColl
                                        
    Next

    Set IJDPinJigAccuracyRule_GetAccuracyObjectsForPurpose = ReturnColl

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5029, , "RULES")
End Function
