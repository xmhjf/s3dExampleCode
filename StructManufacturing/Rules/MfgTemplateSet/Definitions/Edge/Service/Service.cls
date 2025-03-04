VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************************
'Copyright (C) 2012, Intergraph Limited. All rights reserved.
'
'Abstract:
'    Edge Template service for Creation of BasePlane, ControlLine, TemplateOutContour etc.
'    which can be customized by shipyard.
'
'Description:
'History :
'   Siva              20th April 2012      Creation
'******************************************************************************************************************
Option Explicit
Implements IJDMfgTemplateService
Implements IJDMfgTemplateService2

Private Const MODULE = "StrMfgTemplateProcessEdge.Service"

Private Sub Class_Initialize()
 Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    ' No Implementation
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'********************************************************************************************************************************
' Function Name:    CreateBasePlane
' Interface:        IJDMfgTemplateService
' Inputs:           Part -- selected part
'                   oProcessSettings -- Template settings, type, side etc.
' Output:           Object - Baseplane IJPlane
' Assumption:
' Algorithm:        1.
'                   2.
'********************************************************************************************************************************
Private Function IJDMfgTemplateService_CreateBasePlane(ByVal oPart As Object, ByVal oProcessSettings As Object, ByVal oTemplateSet As Object, Optional bUserDefined As Boolean = False) As Object
 Const METHOD = "IJDMfgTemplateService_CreateBasePlane"
    On Error GoTo ErrorHandler

    ' No Implementation

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBP_FaliedToCreateBasePlane, , "RULES")
End Function

'********************************************************************************************************************************
' Function Name:    CreateControlLine
' Interface:        IJDMfgTemplateService
' Inputs:           Part -- selected part
'                   oProcessSettings -- Template settings, type, side etc.
' Output:           Object - BCL
' Assumption:
' Algorithm:

'********************************************************************************************************************************
Private Function IJDMfgTemplateService_CreateControlLine(ByVal oPart As Object, ByVal oProcessSettings As Object, ByVal oBasePlane As Object, ByVal oTemplateSet As Object) As Object
Const METHOD = "IJDMfgControlLineService_CreateControlLine"
    On Error GoTo ErrorHandler
    
    ' No Implementation
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FaliedToCreateBaseControlLine, , "RULES")
End Function

Private Function IJDMfgTemplateService_CreateMarkingLine(ByVal oPart As Object, ByVal oProcessSettings As Object, ByVal pBasePlane As Object) As Object

End Function

Private Function IJDMfgTemplateService_CreateSightLines(ByVal pTemplateSet As Object, ByVal numOfSightLine As Long, ByVal pMarkingSetting As Object) As Object
Const METHOD = "IJDMfgSightLineService_CreateSightLines"
On Error GoTo ErrorHandler
    
    ' No Implementation
    
   Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'********************************************************************************************************************************
' Function Name:    CreateTemplateOutContour
'
' Interface:        IJDMfgTemplateService
'
' Inputs:           Part -- selected part
'                   TemplateSet -- TemplateSet object created by semantic
'                   ProcessSetting -- contains the settings for TemplateSet ex: Type, Side etc
'                   SketchingLines -- Collection of lines sketched by user
'
' Output:           Collection of Geom3d Objects.
'
' Assumption:       Within the output collection, a set of 4 objects make one template. each object represent
'                   the bottom curve, side, top , side of the template
'
' Algorithm:        1. GetSettingValues From ProcessRules
'********************************************************************************************************************************

Private Function IJDMfgTemplateService_CreateTemplateOutContour(ByVal oPart As Object, ByVal pUnkTemplateSet As Object, ByVal oProcessSettings As Object, ByVal pSketchingLines As Object) As Object
    Const METHOD = "IJDMfgTemplateService_CreateTemplateOutContour"
    On Error Resume Next
    
    Dim lTemplateSide As Long
    Dim dTemplateOffset As Double, dTemplateExtn_Base As Double, dTemplateExtn_Offset As Double
    
    ' Create the needed services
    Dim oTemplateInfo As IJElements
    Set oTemplateInfo = New JObjectCollection
    
    Dim oMfgGeomHelper     As GSCADStrMfgUtilities.MfgGeomHelper
    Set oMfgGeomHelper = New GSCADStrMfgUtilities.MfgGeomHelper

    Dim oMfgMGHelper    As GSCADMathGeom.MfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper

    Dim oTemplateHelper As IJMfgTemplateHelper
    Set oTemplateHelper = New MfgTemplateHelper

    Dim oRuleHelper     As MfgRuleHelpers.Helper
    Set oRuleHelper = New MfgRuleHelpers.Helper

    ' Create template set data and cache the settings data
    Dim oTemplSetData   As TemplSetData
    Set oTemplSetData = New TemplSetData
    
    oTemplSetData.InitSettings oPart, oProcessSettings, pUnkTemplateSet
    
    Dim dNA_XValue As Double, dNA_YValue As Double
    Dim dSheer_XValue As Double, dSheer_YValue As Double
    
    Dim strProgId  As String
    strProgId = "MfgProfileProcess.NeutralAxisFixed"
    Dim oNARule As IJDMfgProfileNeutralAxisRule
    Set oNARule = SP3DCreateObject(strProgId)
    oNARule.GetNeutralAxis oPart, dNA_XValue, dNA_YValue
    
    Dim oProfileClass As Object
    Set oProfileClass = New StructDetailObjects.ProfilePart
    Set oProfileClass.object = oPart

    ' Get SheerX and SheerY values
    oProfileClass.Get_SheerCenter dSheer_XValue, dSheer_YValue

    ' Create 3D bottom line
    lTemplateSide = oTemplSetData.TemplateSide
    dTemplateOffset = oTemplSetData.TemplateOffset
    
    dTemplateExtn_Base = oTemplSetData.TemplateBaseExtension
    dTemplateExtn_Offset = oTemplSetData.TemplateOffsetExtension
    
    Dim oTempCurve          As IJCurve
    Dim oBottomCurve        As IJCurve
    
    ' PURPOSE_3D              1
    ' PURPOSE_AFTER_UNTWIST   2
    ' PURPOSE_AFTER_UNBEND    3
    '
    ' LENGTH_LANDING          1
    ' LENGTH_BEFORE_FEATURES  2
    ' LENGTH_AFTER_FEATURES   3

    Dim oGeom2DColl As IJMfgGeomCol2d
    Dim oMfgUnfold As IJMfgUnfold
    Set oMfgUnfold = New MfgUnfold
    
    oMfgUnfold.GetProfileCurve 1, 3, oPart, lTemplateSide, dTemplateOffset, dNA_XValue, dNA_YValue, dSheer_XValue, dSheer_YValue, Nothing, oGeom2DColl, oBottomCurve
    
	'Dim FileName As String
	'FileName = Environ("TEMP")
	'If FileName = "" Or FileName = vbNullString Then
	'	FileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
	'End If
	'FileName = FileName & "\Profilecurve.txt"
    ' oMfgGeomHelper.PrintComplexString oBottomCurve, FileName
    ' Apply extension(linear/perpendicular) to unfolded bottom curve
    If Not dTemplateExtn_Base = 0 Or Not dTemplateExtn_Offset = 0 Then

        ' [in]int extOpt   - 0: extend to point, 1: extend by distance
        ' [in]int extType  - 0: tangent line; 1: polynomial extension
        ' [in]int extEnd   - 0: extend the beginning, 1: extend the end, 2: extend both
        
        If Not dTemplateExtn_Base = 0 Then
            Set oTempCurve = oMfgGeomHelper.ExtraPolateCurve(oBottomCurve, 1, 0, 0, Nothing, Nothing, dTemplateExtn_Base)
            Set oBottomCurve = oTempCurve
        End If
        
        If Not dTemplateExtn_Offset = 0 Then
            Set oTempCurve = oMfgGeomHelper.ExtraPolateCurve(oBottomCurve, 1, 0, 1, Nothing, Nothing, dTemplateExtn_Offset)
        End If
        
    Else
        Set oTempCurve = oBottomCurve
    End If
    
    oTemplateInfo.Add oTempCurve

    Dim oTemplateColl As IJElements
    Set oTemplateColl = New JObjectCollection

    'Add GroupIndex and Sketched information
    Dim oTempPropColl As Collection
    Set oTempPropColl = New Collection

    Dim lGroupIndex     As Long
    Dim strKey          As String

    ' Group index is 1 as there is only one template and key is empty
    lGroupIndex = 1
    strKey = ""

    oTempPropColl.Add lGroupIndex
    oTempPropColl.Add strKey

    oTemplateInfo.Add oTempPropColl
    oTemplateColl.Add oTemplateInfo

    If Not oTemplateColl Is Nothing Then
        Set IJDMfgTemplateService_CreateTemplateOutContour = oTemplateColl
    Else
        Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToCreateTemplateContours, , "RULES")
        GoTo CleanUp
    End If

CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToCreateTemplateContours, , "RULES")
End Function

Private Function IJDMfgTemplateService_CreateTemplatePlane(ByVal oPart As Object, ByVal pUnkPosition As Object, ByVal oProcessSettings As Object, ByVal pDispTemplateSet As Object, Optional ByVal lGroupNumber As Long = 1&) As Object
Const METHOD = "IJDMfgTemplateService_CreateTemplatePlane"
On Error GoTo ErrorHandler

    ' No Implementation
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToCreateBottomLines, , "RULES")
End Function

Private Function IJDMfgTemplateService_ValidateProcessSettings(ByVal oProcessSettings As Object) As Boolean
Const METHOD = "IJDMfgTemplateService_ValidateProcessSettings"
On Error GoTo ErrorHandler

    ' No Implementation
    IJDMfgTemplateService_ValidateProcessSettings = True
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_FailedToValidateProcessSettings, , "RULES")
End Function

' ***********************************************************************************
' Public Function IJDMfgTemplateService2_CreateTemplateUnfoldOutContour()
'
' Description:  creates unfolded output for the template 3D geometries
'
' ***********************************************************************************
Private Function IJDMfgTemplateService2_CreateTemplateUnfoldOutContour(ByVal oPart As Object, ByVal oUnkTemplateSet As Object, ByVal oProcessSettings As Object) As Object
 Const METHOD = "IJDMfgTemplateService2_CreateTemplateUnfoldOutContour"
    On Error GoTo ErrorHandler
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToCreateTemplateContours, , "RULES")
End Function

' ***********************************************************************************
' Public Function IJDMfgTemplateService2_GetTemplateUnfoldRefCurve()
'
' Description:  Returns reference curve used for unfolding of template 3D geometries
'
' ***********************************************************************************
Private Function IJDMfgTemplateService2_GetTemplateUnfoldRefCurve(ByVal oPart As Object, ByVal oUnkTemplateSet As Object, ByVal oProcessSettings As Object) As Object
Const METHOD = "IJDMfgTemplateService2_GetTemplateUnfoldRefCurve"
On Error GoTo ErrorHandler
    
    ' No Implementation
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
