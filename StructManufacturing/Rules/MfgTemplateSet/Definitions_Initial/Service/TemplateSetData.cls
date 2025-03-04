VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateSetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************************
'Copyright (C) 2001, Intergraph Limited. All rights reserved.
'
'Abstract:
'    TemplateSetData object encapsulates the attributes such as side, orientation etc. of a TemplateSet.
'
'Description:
'History :
'   Ram Kandimalla            08/19/2005      Creation.
'******************************************************************************************************************

Option Explicit

Private Const MODULE = "GSCADStrMfgTemplate.TemplateSetData"
Private Const PART_SURFACE = 0
Private Const MOLDED_SURFACE = 1
Private Const PART_SURFACE_BASE = 2

' READ ONLY Properties that are exposed through Public Property Get
Private m_bBaseSide As Boolean
Private m_dMaxHeight As Double
Private m_dMinHeight As Double
Private m_strOrientation As String
Private m_strDirection As String
Private m_strSecondaryDirection As String
Private m_dOffset As Double
Private m_dMinExtn As Double
Private m_dEdgeOffset As Double
Private m_nTemplates As Long
Private m_strPositionEven As String
Private m_strPositionFrames As String
Private m_oFrameElems As IJElements
Private m_dTemplateExtension As Double
Private m_strTemplateType As String
Private m_dInterval As Double
Private m_oAvgNormal As IJDVector
Private m_oAvgPosition As IJDPosition
Private m_oSurfaceBody As IJSurfaceBody
Private m_oPlatePart As IJPlatePart



Public Sub InitSettings(ByVal oPlatePart As IJPlatePart, ByVal oProcessSettings As IJMfgTemplateProcessSettings, ByVal oTemplateSet As IJDMfgTemplateSet)
    Const METHOD = "InitSettings"
    On Error Resume Next
        
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oPlatePart Is Nothing Then
        strErrorMsg = strErrorMsg & "oPlatePart is Nothing "
        GoTo ErrorHandler
    End If
    If oProcessSettings Is Nothing Then
        strErrorMsg = strErrorMsg & "oProcessSettings is Nothing "
        GoTo ErrorHandler
    End If
   
    ' Private variables
    
    Dim oSettingsHelper  As MfgSettingsHelper
    Dim oMfgGeomHelper As MfgGeomHelper
    Dim oMfgMGHelper   As GSCADMathGeom.MfgMGHelper
    Set oSettingsHelper = oProcessSettings
    Set oMfgGeomHelper = New MfgGeomHelper
    Set oMfgMGHelper = New MfgMGHelper
    
    If oSettingsHelper Is Nothing Then
        strErrorMsg = strErrorMsg & "oSettingsHelper is Nothing "
        GoTo ErrorHandler
    End If
    If oMfgGeomHelper Is Nothing Then
        strErrorMsg = strErrorMsg & "oMfgGeomHelper is Nothing "
        GoTo ErrorHandler
    End If
    If oMfgMGHelper Is Nothing Then
        strErrorMsg = strErrorMsg & "oMfgMGHelper is Nothing "
        GoTo ErrorHandler
    End If
    
    'store the plate
    Set m_oPlatePart = oPlatePart
    
    ' Get values from template process settings object
     
    'TemplateType
    strErrorMsg = "TemplateType failed"
    
    Dim strProgId As String
    strProgId = oSettingsHelper.GetProgIDFromAttr("Type")
    
    Dim oTemplateType   As IJMfgTemplateTypeRule
    Set oTemplateType = SP3DCreateObject(strProgId)
    
    oTemplateType.GetTemplateType oPlatePart, m_strTemplateType, m_dInterval
   
    'MaxHeight
    strErrorMsg = "MaxHeight failed"
    
    strProgId = oSettingsHelper.GetProgIDFromAttr("MaxHeight")
    
    Dim oMaxHeight As IJMfgTemplateMaxHeightRule
    Set oMaxHeight = SP3DCreateObject(strProgId)
    m_dMaxHeight = oMaxHeight.GetMaxHeight(oPlatePart)
    
    'MinHeight
    strErrorMsg = "MinHeight failed"
    
    strProgId = oSettingsHelper.GetProgIDFromAttr("MinHeight")
    
    Dim oMinHeight As IJMfgTemplateMinHeightRule
    Set oMinHeight = SP3DCreateObject(strProgId)
    m_dMinHeight = oMinHeight.GetMinHeight(oPlatePart)
    
    'TemplateOrientation
    strErrorMsg = "TemplateOrientation failed"
    'override since other values are invalid
    If m_strTemplateType = "Perpendicular" Then
        m_strOrientation = "Perpendicular"
    Else
        m_strOrientation = oProcessSettings.TemplateOrientation
    End If
    
    'Cache Template Surface
    Dim lSurfaceType As Long
    lSurfaceType = oTemplateSet.SurfaceType
        
    Dim eSurfaceTypeToUse As eStrMfgSurfaceType
    If lSurfaceType = PART_SURFACE Then
        eSurfaceTypeToUse = TRUE_PART
    ElseIf lSurfaceType = MOLDED_SURFACE Then
        eSurfaceTypeToUse = TRUE_MOLD
    ElseIf lSurfaceType = PART_SURFACE_BASE Then
        ' Nothing to do
    Else
        'Error
        GoTo ErrorHandler
    End If
    
    'Get plate surface
    Dim eThicknessSide As PlateThicknessSide
    
    If lSurfaceType = PART_SURFACE Or lSurfaceType = MOLDED_SURFACE Then
        'TemplateSide
        strErrorMsg = "TemplateType failed"
        strProgId = oSettingsHelper.GetProgIDFromAttr("Side")
        
        Dim oPlateSideRule   As IJDMfgPlateUpSideRule
        Dim UpSide As enumPlateSide
        Set oPlateSideRule = SP3DCreateObject(strProgId)
        UpSide = oPlateSideRule.GetPlateUpSide(oPlatePart)
        If (UpSide = BaseSide) Then
            m_bBaseSide = True
        ElseIf (UpSide = OffsetSide) Then
            m_bBaseSide = False
        End If
        
        If m_bBaseSide Then
            eThicknessSide = PlateBaseSide
        Else
            eThicknessSide = PlateOffsetSide
        End If
    End If
    
    On Error Resume Next
    
    If lSurfaceType = PART_SURFACE Or lSurfaceType = MOLDED_SURFACE Then
        Set m_oSurfaceBody = oMfgGeomHelper.GetSurfaceFromPlate(m_oPlatePart, eSurfaceTypeToUse, eThicknessSide, 0)
                            
         If m_oSurfaceBody Is Nothing Then
            ' Try using MOLDED_SURFACE option

            If lSurfaceType = PART_SURFACE Then
                oTemplateSet.SurfaceType = MOLDED_SURFACE
                eSurfaceTypeToUse = TRUE_MOLD
                Set m_oSurfaceBody = oMfgGeomHelper.GetSurfaceFromPlate(m_oPlatePart, eSurfaceTypeToUse, eThicknessSide, 0)
            End If
        End If
    
    End If
    
    If m_oSurfaceBody Is Nothing Then
        ' Try using PART_SURFACE_BASE option i.e.,
        '   1. If plate part is surface then use part geometry as SurfaceBody
        '   2. If plate part is solid then use the system surface without offset

        oTemplateSet.SurfaceType = PART_SURFACE_BASE
        
        Dim oPlateGeom As IJStructGeometry
        Set oPlateGeom = m_oPlatePart

        Set m_oSurfaceBody = oPlateGeom
        
        If m_oSurfaceBody Is Nothing Then ' The plate is solid
            eSurfaceTypeToUse = TRUE_MOLD
            eThicknessSide = PlateSideUnspecified
            
            On Error GoTo GSWF_ErrorHandler
            Set m_oSurfaceBody = oMfgGeomHelper.GetSurfaceFromPlate(m_oPlatePart, eSurfaceTypeToUse, eThicknessSide, 0)
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    'TemplateDirection
    strErrorMsg = "TemplateDirection failed"
    m_strDirection = oProcessSettings.TemplateDirection
    
    If m_strTemplateType = "Box" Or m_strTemplateType = "UserDefined Box" _
        Or m_strTemplateType = "UserDefined Box With Edges" Then
        'Get secondary direction for Templates
        'Get primary and secondary direction vectors
        Dim oDirectionVec As IJDVector
        Set oDirectionVec = New DVector
        If m_strDirection = "Longitudinal" Then 'X - Direction(Buttock)
            oDirectionVec.Set 0, 1, 0
        ElseIf m_strDirection = "Transversal" Then 'Y - Direction(Frame)
            oDirectionVec.Set 1, 0, 0
        Else 'Z - Direction(WaterLine)
            oDirectionVec.Set 0, 0, 1
        End If
        
        'Get Frames in secondary direction
        Dim oParallelAxisVec As IJDVector
        Dim oSecondaryDirVec As IJDVector
        Set oParallelAxisVec = GetParallelAxis(m_oSurfaceBody)
        Set oSecondaryDirVec = oParallelAxisVec.Cross(oDirectionVec)
        oSecondaryDirVec.Length = 1 'make it unit vector (some times its length could be -1)
        
        If Abs(oSecondaryDirVec.x) = 1 And Abs(oSecondaryDirVec.y) = 0 And Abs(oSecondaryDirVec.z) = 0 Then
            m_strSecondaryDirection = "Transversal"
        ElseIf Abs(oSecondaryDirVec.x) = 0 And Abs(oSecondaryDirVec.y) = 1 And Abs(oSecondaryDirVec.z) = 0 Then
            m_strSecondaryDirection = "Longitudinal"
        Else
            m_strSecondaryDirection = "Waterline"
        End If
    End If

    'TemplatePositionEven
    strErrorMsg = "TemplatePositionEven failed"
     m_strPositionEven = oProcessSettings.TemplatePositionEven
    
    If m_strPositionEven = "PositionEven" Then
        strErrorMsg = "GetPositionEven failed"
        
        strProgId = oSettingsHelper.GetProgIDFromAttr("PositionEven")
    
        Dim oPositionEven As IJMfgTemplatePositionEvenRule
        Set oPositionEven = SP3DCreateObject(strProgId)
        oPositionEven.GetPositionEven oPlatePart, m_dEdgeOffset, m_nTemplates
    End If
    
    'TemplatePositionFrames
    strErrorMsg = "TemplatePositionFrames failed"
    m_strPositionFrames = oProcessSettings.TemplatePositionFrames
    
    'we need the frames in all cases
    'strPositionFrames = "PositionFrame"
    If m_strPositionFrames = "PositionFrame" And Not oProcessSettings.TemplateType = "UserDefined" Then
        strErrorMsg = "GetPositionFrame failed"
        
        Dim oFrameSystem As IHFrameSystem
        Dim oMfgFrameSys As IJDMfgFrameSystem
        If Not oTemplateSet Is Nothing Then
            Set oMfgFrameSys = oTemplateSet
            Set oFrameSystem = oMfgFrameSys.FrameSysParent
            If oFrameSystem Is Nothing Then
                Err.Number = StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3003, , "RULES")
            End If
        End If
        
        strProgId = oSettingsHelper.GetProgIDFromAttr("PositionFrames")
        Dim oPositionFrames As IJMfgTemplatePositionFrameRule
        Set oPositionFrames = SP3DCreateObject(strProgId)
        
        Set m_oFrameElems = oPositionFrames.GetPositionFrame(oFrameSystem, oPlatePart, oProcessSettings, oTemplateSet)
    End If
    
    'TemplateExtension
    strErrorMsg = "TemplateExtension failed"
     
    strProgId = oSettingsHelper.GetProgIDFromAttr("Extension")
    
    Dim oExtension As Object
    Dim oTempExtnRule As IJMfgTemplateExtensionRule
    Dim oTempExtnRule2 As IJMfgTemplateExtensionRule2
    
    Set oExtension = SP3DCreateObject(strProgId)
    
    If TypeOf oExtension Is IJMfgTemplateExtensionRule Then
        Set oTempExtnRule = oExtension
        m_dTemplateExtension = oTempExtnRule.GetExtension(oPlatePart)
    ElseIf TypeOf oExtension Is IJMfgTemplateExtensionRule2 Then
        Set oTempExtnRule2 = oExtension
        oTempExtnRule2.GetExtension oPlatePart, m_dTemplateExtension, m_dOffset, m_dMinExtn
    End If
    
    Dim dAvgNorDirOfPlateX As Double, dAvgNorDirOfPlateY As Double, dAvgNorDirOfPlateZ As Double
    Dim dRX As Double, dRY As Double, dRZ As Double
    
    oMfgGeomHelper.GetPlatePartAvgPointAvgNormal m_oSurfaceBody, m_bBaseSide, dRX, dRY, dRZ, _
                                                 dAvgNorDirOfPlateX, dAvgNorDirOfPlateY, dAvgNorDirOfPlateZ
     
    
    Set m_oAvgNormal = New DVector
    m_oAvgNormal.Set dAvgNorDirOfPlateX, dAvgNorDirOfPlateY, dAvgNorDirOfPlateZ
    m_oAvgNormal.Length = 1
    
    Set m_oAvgPosition = New DPosition
    m_oAvgPosition.Set dRX, dRY, dRZ
    
Cleanup:
    Set oSettingsHelper = Nothing
    Set oExtension = Nothing
    Set oPositionFrames = Nothing
    Set oPositionEven = Nothing
    Set oMinHeight = Nothing
    Set oMaxHeight = Nothing
    Set oTemplateType = Nothing
    Set oPlateSideRule = Nothing
    Exit Sub
    
GSWF_ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_FailedToGetSurfaceFromPlatePart, , "RULES")
    GoTo Cleanup
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo Cleanup
End Sub
Public Property Get IsOnBaseSide() As Boolean
    IsOnBaseSide = m_bBaseSide
End Property
Public Property Get TemplateSide() As Boolean
    TemplateSide = m_bBaseSide
End Property
Public Property Get MaxHeight() As Double
    MaxHeight = m_dMaxHeight
End Property
Public Property Get MinHeight() As Double
    MinHeight = m_dMinHeight
End Property
Public Property Get Orientation() As String
    Orientation = m_strOrientation
End Property
Public Property Get Direction() As String
    Direction = m_strDirection
End Property

Public Property Get Offset() As Double
    Offset = m_dOffset
End Property

' valid for PosistionEven templates only
Public Property Get EdgeOffset() As Double
    EdgeOffset = m_dEdgeOffset
End Property

' valid for PosistionEven templates only
Public Property Get TemplatesCount() As Long
    TemplatesCount = m_nTemplates
End Property

Public Property Get PositionEven() As String
    PositionEven = m_strPositionEven
End Property
Public Property Get PositionFrames() As String
    PositionFrames = m_strPositionFrames
End Property
Public Property Get FrameElems() As IJElements
    Set FrameElems = m_oFrameElems
End Property
Public Property Get TemplateExtension() As Double
    TemplateExtension = m_dTemplateExtension
End Property
Public Property Get TemplateMinimumExtension() As Double
    TemplateMinimumExtension = m_dMinExtn
End Property
Public Property Get TemplateType() As String
    TemplateType = m_strTemplateType
End Property
Public Property Get Interval() As Double
    Interval = m_dInterval
End Property

Public Property Get PlatePosition() As IJDPosition
    Set PlatePosition = m_oAvgPosition
End Property

Public Property Get PlateNormal() As IJDVector
    Set PlateNormal = m_oAvgNormal
End Property
Public Property Get SurfaceBody() As IJSurfaceBody
    Set SurfaceBody = m_oSurfaceBody
End Property


Private Sub Class_Terminate()
    Set m_oFrameElems = Nothing
    Set m_oAvgNormal = Nothing
    Set m_oAvgPosition = Nothing
    Set m_oSurfaceBody = Nothing
End Sub

Public Property Get SecondaryDirection() As String
    SecondaryDirection = m_strSecondaryDirection
End Property


