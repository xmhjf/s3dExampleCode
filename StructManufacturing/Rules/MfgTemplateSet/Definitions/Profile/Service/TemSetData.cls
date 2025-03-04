VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemSetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Template Set Data Private Class
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************


Option Explicit

Private Const MODULE = "GSCADStrMfgTemplate.TemSetData"
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
Private m_dEdgeOffset As Double
Private m_nTemplates As Long
Private m_strPosition As String
Private m_oFrameElems As IJElements
Private m_dLSPortion As Double
Private m_dLSFixedExtn As Double
Private m_dLSOffsetExtn As Double
Private m_dLSMinExtnforOffset As Double
Private m_dUSPortion As Double
Private m_dUSFixedExtn As Double
Private m_dUSOffsetExtn As Double
Private m_dUSMinExtnforOffset As Double
Private m_strTemplateType As String
Private m_dInterval As Double
Private m_oAvgNormal As IJDVector
Private m_oAvgPosition As IJDPosition
Private m_oSurfaceBody As IJSurfaceBody
Private m_oProfilePart As IJProfilePart
Private m_bSupportEdges As Boolean
Private m_oForwardVec As IJDVector
Private m_oUpVec As IJDVector

Public Sub InitSettings(ByVal oProfilePart As IJProfilePart, ByVal oProcessSettings As IJMfgTemplateProcessSettings, ByVal oTemplateSet As IJDMfgTemplateSet)
    Const METHOD = "InitSettings"
    On Error Resume Next
        
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oProfilePart Is Nothing Then
        strErrorMsg = strErrorMsg & "oProfilePart is Nothing "
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
    Set m_oProfilePart = oProfilePart
    
    ' Get values from template process settings object
     
    'TemplateType
    strErrorMsg = "TemplateType failed"
    
    Dim strProgId As String
    strProgId = oSettingsHelper.GetProgIDFromAttr("Type")
    
    Dim oTemplateType   As IJMfgTemplateTypeRule
    Set oTemplateType = SP3DCreateObject(strProgId)
    
    Dim dUnUsedOffSet As Double 'Offset now comes from ExtensionByOffSet Rule
    oTemplateType.GetTemplateType oProfilePart, m_strTemplateType, m_dInterval
   
    'MaxHeight
    strErrorMsg = "MaxHeight failed"
    
    strProgId = oSettingsHelper.GetProgIDFromAttr("MaxHeight")
    
    Dim oMaxHeight As IJMfgTemplateMaxHeightRule
    Set oMaxHeight = SP3DCreateObject(strProgId)
    m_dMaxHeight = oMaxHeight.GetMaxHeight(oProfilePart)
    
    'MinHeight
    strErrorMsg = "MinHeight failed"
    
    strProgId = oSettingsHelper.GetProgIDFromAttr("MinHeight")
    
    Dim oMinHeight As IJMfgTemplateMinHeightRule
    Set oMinHeight = SP3DCreateObject(strProgId)
    m_dMinHeight = oMinHeight.GetMinHeight(oProfilePart)
    
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
    
    'TemplateSide
    strErrorMsg = "TemplateType failed"
    strProgId = oSettingsHelper.GetProgIDFromAttr("Side")
    
    Dim oProfileSideRule   As IJDMfgProfileUpsideRule
    Dim UpSide As Long
    Set oProfileSideRule = SP3DCreateObject(strProgId)
    UpSide = oProfileSideRule.GetUpside(oProfilePart)
    
    Dim oUpsideSurface As IUnknown
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oProfilePart
    
    Dim oSurfacePort As IJPort
    Set oSurfacePort = oProfileWrapper.GetSurfacePort(UpSide)
    Set m_oSurfaceBody = oSurfacePort.Geometry
    
    On Error GoTo ErrorHandler
    
    'TemplateDirection
    strErrorMsg = "TemplateDirection failed"
    m_strDirection = oProcessSettings.TemplateDirection
    
    'Template Position
     m_strPosition = oProcessSettings.TemplatePosition
    
    strErrorMsg = "GetPositionEven failed"
    
    strProgId = oSettingsHelper.GetProgIDFromAttr("Position")

    If strProgId = "Ignore" Then
        'nothing to do..
    Else
        Dim oPosition As IJMfgTemplatePositionRule
        Set oPosition = SP3DCreateObject(strProgId)
        Dim oFrameSystem As IHFrameSystem
        Dim oMfgFrameSys As IJDMfgFrameSystem
        If Not oTemplateSet Is Nothing Then
            Set oMfgFrameSys = oTemplateSet
            Set oFrameSystem = oMfgFrameSys.FrameSysParent
            If oFrameSystem Is Nothing Then
                Err.Number = StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3003, , "RULES")
            End If
        End If
        
        oPosition.GetPositionData oFrameSystem, oProfilePart, oProcessSettings, oTemplateSet, m_dEdgeOffset, m_nTemplates, m_bSupportEdges, m_oFrameElems
    End If
    
    
    'TemplateExtension
    strErrorMsg = "TemplateExtension failed"
     
    strProgId = oSettingsHelper.GetProgIDFromAttr("Extension")
    
    Dim oExtension As Object
    Dim oTempExtnRule3 As IJMfgTemplateExtensionRule3
    
    Set oExtension = SP3DCreateObject(strProgId)
    
    If TypeOf oExtension Is IJMfgTemplateExtensionRule3 Then
        Set oTempExtnRule3 = oExtension
        oTempExtnRule3.GetExtension oProfilePart, _
                                    m_dLSPortion, _
                                    m_dLSFixedExtn, _
                                    m_dLSOffsetExtn, _
                                    m_dLSMinExtnforOffset, _
                                    m_dUSPortion, _
                                    m_dUSFixedExtn, _
                                    m_dUSOffsetExtn, _
                                    m_dUSMinExtnforOffset
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
    
    'initialize report directions
    oTemplateSet.GetDirectionVector m_oForwardVec, m_oUpVec, 1
    
CleanUp:
    Set oSettingsHelper = Nothing
    Set oExtension = Nothing
    Set oPosition = Nothing
    Set oMinHeight = Nothing
    Set oMaxHeight = Nothing
    Set oTemplateType = Nothing
    Exit Sub
    
GSWF_ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_FailedToGetSurfaceFromPlatePart, , "RULES")
    GoTo CleanUp
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub
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

Public Property Get LSPortion() As Double
    LSPortion = m_dLSPortion
End Property
Public Property Get LSFixedExtn() As Double
    LSFixedExtn = m_dLSFixedExtn
End Property
Public Property Get LSOffsetExtn() As Double
    LSOffsetExtn = m_dLSOffsetExtn
End Property
Public Property Get LSMinExtnForOffset() As Double
    LSMinExtnForOffset = m_dLSMinExtnforOffset
End Property
Public Property Get USPortion() As Double
    USPortion = m_dUSPortion
End Property
Public Property Get USFixedExtn() As Double
    USFixedExtn = m_dUSFixedExtn
End Property
Public Property Get USOffsetExtn() As Double
    USOffsetExtn = m_dUSOffsetExtn
End Property
Public Property Get USMinExtnForOffset() As Double
    USMinExtnForOffset = m_dUSMinExtnforOffset
End Property
' valid for PosistionEven templates only
Public Property Get EdgeOffset() As Double
    EdgeOffset = m_dEdgeOffset
End Property

' valid for PosistionEven templates only
Public Property Get TemplatesCount() As Long
    TemplatesCount = m_nTemplates
End Property

Public Property Get Position() As String
    Position = m_strPosition
End Property
Public Property Get FrameElems() As IJElements
    Set FrameElems = m_oFrameElems
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

Public Property Get SupportEdges() As Boolean
    SupportEdges = m_bSupportEdges
End Property

Public Property Get ForwardVector() As IJDVector
    Set ForwardVector = m_oForwardVec
End Property
Public Property Get UpVector() As IJDVector
    Set UpVector = m_oUpVec
End Property

