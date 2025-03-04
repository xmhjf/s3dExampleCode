VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KnuckleMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project: MfgTemplateMarking
'
'  Abstract:    Creates Knuckle Marks on Template
'
'  History:
'      Siva        2nd September 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateMarkingPlate.KnuckleMark"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler
    
    Dim oMfgGeomChild   As IJMfgGeomChild
    Set oMfgGeomChild = oMfgTemplate
        
    Dim oMfgTemplateSetRep  As IJMfgTemplateSetReport
    Set oMfgTemplateSetRep = oMfgGeomChild.GetParent
    
    Dim oPlatePart  As Object
    Set oPlatePart = oMfgTemplateSetRep.GetParentPlatePart
    
    'Create the SD plate Wrapper and initialize it
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = oPlatePart
    
    ' If the part is of hull type, reference marks rule will do
    If oSDPlateWrapper.plateType = Hull Then
        Exit Function
    End If
    
    Dim oTemplateBottomCurve   As IJBSplineCurve
    Set oTemplateBottomCurve = oMfgTemplate.GetTemplateLocationMarkLine
    
    Dim oTemplateBottomCS As IJComplexString
    Set oTemplateBottomCS = New ComplexString3d
    
    oTemplateBottomCS.AddCurve oTemplateBottomCurve, False
    
    If oTemplateBottomCS Is Nothing Then
        Exit Function
    End If
    
    Dim oGeomElems As IJElements
    Set oGeomElems = New JObjectCollection
    
    Dim nIndex          As Long
    Dim oMfgRuleHelper  As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New Helper
    
    Dim oWireBodyUtils  As IJSGOWireBodyUtilities
    Set oWireBodyUtils = New SGOWireBodyUtilities
    
    Dim oBottomLineWB   As IJWireBody
    Set oBottomLineWB = oMfgRuleHelper.ComplexStringToWireBody(oTemplateBottomCS)
    
    Dim oKnucklePtColl  As Collection
    Set oKnucklePtColl = oWireBodyUtils.GetKnucklePoints(oBottomLineWB)
    
    If oKnucklePtColl Is Nothing Then
        'Since there are no knuckles we can leave the marking rule
        GoTo CleanUp
    End If
    
    For nIndex = 1 To oKnucklePtColl.Count
        Dim oMarkPosition   As IJDPosition
        Set oMarkPosition = oKnucklePtColl.Item(nIndex)
        
        Dim oMarkVec    As IJDVector
        Set oMarkVec = GetMarkVector(oMfgTemplate, oTemplateBottomCurve, oMarkPosition)
        
        Dim oMarkCS As IJComplexString
        Set oMarkCS = CreateMarkAtPosition(oMarkPosition, oMarkVec, TEMPLATE_KNUCKLE_MARK_LENGTH, True)
        
        Dim oGeom3D As IJMfgGeom3d
        Set oGeom3D = CreateGeom3D(oMarkCS, STRMFG_KNUCKLE_MARK, , "Knuckle Mark")
        
        oGeomElems.Add oGeom3D
    Next
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oGeomElems
    
CleanUp:
    Exit Function
    
ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3019, , "RULES")
End Function
