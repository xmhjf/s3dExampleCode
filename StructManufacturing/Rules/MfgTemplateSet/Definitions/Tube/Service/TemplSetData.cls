VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplSetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************************
'Copyright (C) 2001, Intergraph Limited. All rights reserved.
'
'Abstract:
'    TemplSetData object encapsulates the attributes such as side, orientation etc. of a TemplateSet.
'
'Description:
'History :
'   Siva              20th Oct 2011      Creation
'******************************************************************************************************************

Option Explicit

Private Const MODULE = "StrMfgTemplateProcessTube.TubeTemplSetData"

' READ ONLY Properties that are exposed through Public Property Get
Private m_dMaxHeight    As Double
Private m_dMinHeight    As Double
Private m_strTemplateType   As String
Private m_strBasePlane      As String
Private m_strTemplateSide     As String
Private m_strTemplateSideEnd  As String

Private m_lMethod_Left As Long, m_lMethod_Right As Long, m_lMeasure_Left As Long, m_lMeasure_Right As Long
Private m_dValue1_Left As Double, m_dValue2_Left As Double, m_dValue1_Right As Double, m_dValue2_Right As Double

Public Sub InitSettings(ByVal oTubePart As Object, ByVal oProcessSettings As IJMfgTemplateProcessSettings, ByVal oTemplateSet As IJDMfgTemplateSet)
    Const METHOD = "InitSettings"
    On Error Resume Next
        
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTubePart Is Nothing Then
        strErrorMsg = strErrorMsg & "oTubePart is Nothing "
        GoTo ErrorHandler
    End If
    If oProcessSettings Is Nothing Then
        strErrorMsg = strErrorMsg & "oProcessSettings is Nothing "
        GoTo ErrorHandler
    End If
    
    Dim oSettingsHelper  As MfgSettingsHelper
    Set oSettingsHelper = oProcessSettings
    
    If oSettingsHelper Is Nothing Then
        strErrorMsg = strErrorMsg & "oSettingsHelper is Nothing "
        GoTo ErrorHandler
    End If
    
    ' Get values from template process settings object
    ' TemplateType
    strErrorMsg = "TemplateType failed"
    
    Dim strProgId As String
   
    'MaxHeight
    strErrorMsg = "MaxHeight failed"
    strProgId = oSettingsHelper.GetProgIDFromAttr("MaxHeight")
    Dim oMaxHeight As IJMfgTubeTemplateMaxHeightRule
    Set oMaxHeight = SP3DCreateObject(strProgId)
    m_dMaxHeight = oMaxHeight.GetMaxHeight(oTubePart)
    
    'MinHeight
    strErrorMsg = "MinHeight failed"
    strProgId = oSettingsHelper.GetProgIDFromAttr("MinHeight")
    Dim oMinHeight As IJMfgTubeTemplateMinHeightRule
    Set oMinHeight = SP3DCreateObject(strProgId)
    m_dMinHeight = oMinHeight.GetMinHeight(oTubePart)
        
    Dim bstrTableName   As String
    bstrTableName = oSettingsHelper.GetTableNameFromAttr("BasePlane")
    m_strBasePlane = oSettingsHelper.GetShortDescription(bstrTableName, "BasePlane")
        
    bstrTableName = oSettingsHelper.GetTableNameFromAttr("Side")
    m_strTemplateSide = oSettingsHelper.GetShortDescription(bstrTableName, "Side")
    
    bstrTableName = oSettingsHelper.GetTableNameFromAttr("SideEnd")
    m_strTemplateSideEnd = oSettingsHelper.GetShortDescription(bstrTableName, "SideEnd")
    
    bstrTableName = oSettingsHelper.GetTableNameFromAttr("Type")
    m_strTemplateType = oSettingsHelper.GetShortDescription(bstrTableName, "Type")
        
    strProgId = oSettingsHelper.GetProgIDFromAttr("Extension")
    
    Dim oTempExtnRule As IJMfgTubeTemplateExtensionRule
    Set oTempExtnRule = SP3DCreateObject(strProgId)
     
    oTempExtnRule.GetExtension oTubePart, m_lMethod_Left, m_lMethod_Right, m_lMeasure_Left, m_lMeasure_Right, _
                                            m_dValue1_Left, m_dValue2_Left, m_dValue1_Right, m_dValue2_Right
CleanUp:
    Set oSettingsHelper = Nothing
    Set oMinHeight = Nothing
    Set oMaxHeight = Nothing
    Set oTempExtnRule = Nothing
    
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

Public Property Get TemplateType() As String
    TemplateType = m_strTemplateType
End Property

Public Property Get TemplateBasePlane() As String
    TemplateBasePlane = m_strBasePlane
End Property

Public Property Get TemplateSide() As String
    TemplateSide = m_strTemplateSide
End Property

Public Property Get TemplateSideEnd() As String
    TemplateSideEnd = m_strTemplateSideEnd
End Property

Public Property Get TemplateMethodLeft() As Double
    TemplateMethodLeft = m_lMethod_Left
End Property

Public Property Get TemplateMethodRight() As Double
    TemplateMethodRight = m_lMethod_Right
End Property

Public Property Get TemplateMeasureLeft() As Double
    TemplateMeasureLeft = m_lMeasure_Left
End Property

Public Property Get TemplateMeasureRight() As Double
    TemplateMeasureRight = m_lMeasure_Right
End Property

Public Property Get TemplateExtnValueLeft() As Double
    TemplateExtnValueLeft = m_dValue1_Left
End Property

Public Property Get TemplateExtnValueRight() As Double
    TemplateExtnValueRight = m_dValue1_Right
End Property

