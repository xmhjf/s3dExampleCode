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
'  Project  : StrMfgTemplateMarkingEdge
'
'  Abstract : Creates Ship Direction Marks on Edge Template
'
'  History  :
'      Siva   20th April 2012    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateMarkingEdge.ShipDirMark"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler

    Dim oBottomCurve3D  As IJCurve
    Dim oBottomCurve2D  As IJCurve
    Dim oUVLine2D As IJCurve
    Dim oUVLine3D As IJCurve
    
    ' Get the reference objects like 2D bottom curve etc
    Dim iIndex  As Long
    For iIndex = 1 To oReferenceObjColl.Count
        Dim oTempObj    As Object
        Set oTempObj = oReferenceObjColl.Item(iIndex)
        
        If TypeOf oTempObj Is IJMfgGeom3d Then
            Dim oTempGeom3d As IJMfgGeom3d
            Set oTempGeom3d = oTempObj
            
            If oTempGeom3d.GetGeometryType = STRMFG_TemplateLocationMarkLine Then
                Set oBottomCurve3D = oTempGeom3d.GetGeometry
            ElseIf oTempGeom3d.GetGeometryType = STRMFG_UV_MARK Then
                Set oUVLine3D = oTempGeom3d.GetGeometry
            End If
        ElseIf TypeOf oTempObj Is IJMfgGeom2d Then
            Dim oTempGeom2d As IJMfgGeom2d
            Set oTempGeom2d = oTempObj
            
            If oTempGeom2d.GetGeometryType = STRMFG_TemplateLocationMarkLine Then
                Set oBottomCurve2D = oTempGeom2d.GetGeometry
            ElseIf oTempGeom2d.GetGeometryType = STRMFG_UV_MARK Then
                Set oUVLine2D = oTempGeom2d.GetGeometry
            End If
        End If
    Next
    
    ' Exit when inputs are not available
    If oUVLine2D Is Nothing Then
        Exit Function
    End If
    
    Dim oMGHelper   As IJMfgMGHelper
    Set oMGHelper = New MfgMGHelper
    
    Dim oWire3D As IJWireBody
    oMGHelper.ComplexStringToWireBody oBottomCurve3D, oWire3D
    
    Dim oWire2D As IJWireBody
    oMGHelper.ComplexStringToWireBody oBottomCurve2D, oWire2D

    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New Helper

    Dim oMidPos3D As IJDPosition
    Set oMidPos3D = oMfgRuleHelper.GetMiddlePoint(oWire3D)
    
    Dim oMidPos2D As IJDPosition
    Set oMidPos2D = oMfgRuleHelper.GetMiddlePoint(oWire2D)
    
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    Dim oV_Vec3D As IJDVector, oU_Vec3D As IJDVector, oTempVec As IJDVector
    Dim oV_Vec2D As IJDVector, oU_Vec2D As IJDVector
    
    Set oU_Vec3D = oMfgGeomHelper.GetTangentByPointOnCurve(oBottomCurve3D, oMidPos3D)
    oU_Vec3D.length = 1
    
    Set oU_Vec2D = oMfgGeomHelper.GetTangentByPointOnCurve(oBottomCurve2D, oMidPos2D)
    oU_Vec2D.length = 1
    
    Dim dDirX As Double, dDirY As Double, dDirZ As Double, dDummy As Double
    oUVLine3D.Evaluate 0, dDummy, dDummy, dDummy, dDirX, dDirY, dDirZ, dDummy, dDummy, dDummy
    
    Set oV_Vec3D = New DVector
    oV_Vec3D.Set dDirX, dDirY, dDirZ
    oV_Vec3D.length = 1
        
    oUVLine2D.Evaluate 0, dDummy, dDummy, dDummy, dDirX, dDirY, dDirZ, dDummy, dDummy, dDummy
    
    Set oV_Vec2D = New DVector
    oV_Vec2D.Set dDirX, dDirY, dDirZ
    oV_Vec2D.length = 1
    
    Set oTempVec = oV_Vec2D.Add(oU_Vec2D)
    oTempVec.length = 0.3
    
    Dim oRootPos As IJDPosition
    Set oRootPos = oMidPos2D.Offset(oTempVec)
    
    ' Create BCL mark of fixed length
    Dim oDirMark    As IJCurve
    Set oDirMark = CreateLineAtPosition(oRootPos, oU_Vec2D, TEMPLATE_SHIP_DIR_PRIMARY_LENGTH)
        
    Dim oTempElems      As IJElements
    Set oTempElems = New JObjectCollection
    
    Dim strDirection    As String
    strDirection = GetDirectionString(oU_Vec3D, oMidPos3D)
    
    ' Create Geom3D for Ship Direction mark - Primary
    Dim oMfgGeom2D As IJMfgGeom2d
    Set oMfgGeom2D = CreateGeom2D(oDirMark, STRMFG_DIRECTION, Nothing, strDirection)
    
    oTempElems.Add oMfgGeom2D
    
    Set oDirMark = Nothing
    Set oDirMark = CreateLineAtPosition(oRootPos, oV_Vec2D, TEMPLATE_SHIP_DIR_SECONDARY_LENGTH)
    
    strDirection = GetDirectionString(oV_Vec3D, oMidPos3D)
    
    ' Create Geom3D for Ship Direction mark - Secondary
    Set oMfgGeom2D = Nothing
    Set oMfgGeom2D = CreateGeom2D(oDirMark, STRMFG_DIRECTION, Nothing, strDirection)
    
    oTempElems.Add oMfgGeom2D
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oTempElems
    Exit Function
    
ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3022, , "RULES")
End Function
