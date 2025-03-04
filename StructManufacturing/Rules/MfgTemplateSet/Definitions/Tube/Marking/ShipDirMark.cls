VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShipDirMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project  : StrMfgTemplateMarkingTube
'
'  Abstract : Creates Ship Direction Marks on Template
'
'  History  :
'      Siva   20th October 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateMarkingTube.ShipDirMark"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler

    Dim oTempElems      As IJElements
    Set oTempElems = New JObjectCollection
    
    Dim oBottomCurve2D  As IJCurve
    Dim oTopCurve2D     As IJCurve
    Dim oRefCurve       As IJCurve
    
    ' Get the 2D Bottom curve, top curve and ref curve
    Dim iIndex  As Long
    For iIndex = 1 To oReferenceObjColl.Count
        Dim oTempObj    As Object
        Set oTempObj = oReferenceObjColl.Item(iIndex)
        
        If TypeOf oTempObj Is IJMfgGeom2d Then
            Dim oGeom2D As IJMfgGeom2d
            Set oGeom2D = oTempObj
            If oGeom2D.GetGeometryType = STRMFG_TemplateLocationMarkLine Then
                Set oBottomCurve2D = oGeom2D.GetGeometry
            ElseIf oGeom2D.GetGeometryType = STRMFG_TopLine Then
                Set oTopCurve2D = oGeom2D.GetGeometry
            End If
        Else
            If TypeOf oTempObj Is IJCurve Then
                Set oRefCurve = oReferenceObjColl.Item(iIndex)
            End If
        End If
    Next
    
    Dim oMarkVec    As IJDVector
    Set oMarkVec = New DVector
    
    oMarkVec.Set 0, -1, 0
    
    Dim dRefCurveLen    As Double
    dRefCurveLen = oRefCurve.length
    
    Dim oInputPos As IJDPosition
    Set oInputPos = New DPosition
    
    ' Create direction mark right side of the 2D template geometry at dist of (0.4 * dRefCurveLen) from BCL
    oInputPos.Set (0.4 * dRefCurveLen), 0, 0
    
    Dim oVertCurve  As IJCurve
    Set oVertCurve = CreateLineAtPosition(oInputPos, oMarkVec, 5, True)
    
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    ' Get the nearest point of "input pos on 2D ref curve" on the 2D top curve
    Dim oOutputPos As IJDPosition
    Dim dDist   As Double
    
    ' Get the intersection of vertical line created above and top curve.
    oMfgGeomHelper.GetDistBet2Curves oTopCurve2D, oVertCurve, dDist, oOutputPos, oOutputPos
    
    ' Offset the intersection position for placing ship direction mark
    oMarkVec.length = 0.1
    Dim oMarkPos As IJDPosition
    Set oMarkPos = oOutputPos.Offset(oMarkVec)
    
    Dim oMfgGeomChild   As IJMfgGeomChild
    Set oMfgGeomChild = oMfgTemplate
    
    Dim oMfgTemplateSet As IJDMfgTemplateSet
    Set oMfgTemplateSet = oMfgGeomChild.GetParent
    
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    oMfgTemplateSet.GetBasePlaneNormalVector dNormalX, dNormalY, dNormalZ
    
    Dim oControlCurve   As IJCurve
    Set oControlCurve = oMfgTemplateSet.GetControlLine
    
    Dim oBasePlane  As IJPlane
    Set oBasePlane = oMfgTemplateSet.GetBasePlane
    
    Dim oNormalVec As IJDVector
    Set oNormalVec = New DVector
    oNormalVec.Set 0, 1, 0
    
    Dim oRootPoint As IJDPosition
    Set oRootPoint = New DPosition
    
    Dim oCenPlane   As IJPlane
    Set oCenPlane = CreatePlane(oRootPoint, oNormalVec)
    
    On Error Resume Next    ' Intersection routine can fail if there is no intersection
    
    Dim oIntersectionObj    As Object
    Set oIntersectionObj = oMfgGeomHelper.IntersectCurveWithPlane(oControlCurve, oCenPlane)
    
    Dim bCenterCross As Boolean
    If Not oIntersectionObj Is Nothing Then
        bCenterCross = True
    End If
       
    On Error GoTo ErrorHandler
    
    ' Create ship direction mark at the input position
    Dim oShipDirMark As Object
    Set oShipDirMark = CreateShipDirectionMark(bCenterCross, oMarkPos, oBasePlane, TEMPLATE_SHIP_DIR_SECONDARY_LENGTH)
      
    oTempElems.Add oShipDirMark
    Set oShipDirMark = Nothing
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oTempElems
    Exit Function
    
ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3022, , "RULES")
End Function

Public Function CreateShipDirectionMark(bCenterCross As Boolean, oInputPos As IJDPosition, oBasePlane As IJPlane, oMarkLen As Double) As Object
    Const METHOD = "CreateShipDirectionMark"
    On Error GoTo ErrorHandler
    
    Dim oMarkVec    As IJDVector
    Set oMarkVec = New DVector
    
    oMarkVec.Set 0, -1, 0
    
    Dim oShipDirMark    As IJCurve
    Set oShipDirMark = CreateLineAtPosition(oInputPos, oMarkVec, oMarkLen)
    
    Dim dRootX As Double, dRootY As Double, dRootZ As Double
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    oBasePlane.GetNormal dNormalX, dNormalY, dNormalZ
    oBasePlane.GetRootPoint dRootX, dRootY, dRootZ
    
    Dim oNormalVec  As IJDVector
    Set oNormalVec = New DVector
    oNormalVec.Set dNormalX, dNormalY, dNormalZ
    
    ' As Base plane normal points to the end cut, reverse the normal
    'oNormalVec.length = -1
    
    Dim oRootPos  As IJDPosition
    Set oRootPos = New DPosition
    oRootPos.Set dRootX, dRootY, dRootZ
    
    ' Get the direction string of the mark
    Dim strDirection    As String
    strDirection = GetDirectionString(oNormalVec, oRootPos)
    
    ' Create Geom2D for the ship direction mark
    Dim oMfgGeom2D As IJMfgGeom2d
    Set oMfgGeom2D = CreateGeom2D(oShipDirMark, STRMFG_DIRECTION, Nothing, strDirection)
    
    Set CreateShipDirectionMark = oMfgGeom2D
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
