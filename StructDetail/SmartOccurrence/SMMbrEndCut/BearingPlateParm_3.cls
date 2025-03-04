VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BearingPlateParm_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2014 Intergraph Corporation. All rights reserved.
'
'File : BearingPlateParm_3.cls
'
'Author : Alligators
'
'Description :
'   Control Parameters of the Member Items for the MbrBearingEndSel Selector Class
'
'
'History:

'   26/11/2014   Created
'
'   14/July/2015 -knukala
'           TR-CP-270890  Cannot Place Generic AC when Handrail Post member is bounded to a plate
'********************************************************************
Option Explicit

Const m_sClassName As String = "BearingPlateParm_3"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices


'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ParameterRule Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"

    pOH.SetOutput "Offset"
    pOH.SetOutput "Radius"
    pOH.SetOutput "FlangeThickness"
    pOH.SetOutput "WebThickness"
    

  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    sMsg = "Defining Parameter Values"
    
    '***   We need to set four parameters  ***
    
    'Offset and Radius parameters are retained here for future use (incase of different dimensions for different size of crosssection):
    'As of now, below code gets catalog values and applies the same
    
	Dim dOffset As Double
    Dim dRadius As Double
    Dim oPortBounded As IJPort
    Dim oPortBounding As IJPort
    Dim oBoundedObject As Object
    sMsg = "Defining Parameter Values"
    Dim oAttributes  As IJDAttributes
    Dim oSmartOccurrence As IJSmartOccurrence
    
    Set oSmartOccurrence = pPRL.SmartOccurrence
    Set oAttributes = oSmartOccurrence
	'***   First parameter:
    dOffset = oAttributes.CollectionOfAttributes("IJUAMbrEndCuts").Item("Offset").value
    pPRL.Add "Offset", dOffset
    
	'***   Second parameter:
    dRadius = oAttributes.CollectionOfAttributes("IJUAMbrEndCuts").Item("Radius").value
    pPRL.Add "Radius", dRadius
    
    ' To Get the flange thickness and web thickness from member
    
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
    Set oPortBounding = pPRL.InputObject(INPUT_BOUNDING)
    Set oBoundedObject = oPortBounded.Connectable
    
    Dim dFlangeThickness As Double
    Dim dWebThickness As Double
    
    Dim oBoundedMemberPart As New StructDetailObjects.MemberPart
    Set oBoundedMemberPart.object = oBoundedObject
    dFlangeThickness = oBoundedMemberPart.flangeThickness
    dWebThickness = oBoundedMemberPart.webThickness
    
    Dim oBoundedPart As ISPSMemberPartPrismatic
    Dim oBoundingPart As ISPSMemberPartPrismatic
    Dim oBoundedAxisPort As ISPSSplitAxisPort
    
    Set oBoundedPart = oPortBounded.Connectable
    Set oBoundedAxisPort = oPortBounded
    If TypeOf oPortBounding.Connectable Is ISPSMemberPartCommon Then
        Set oBoundingPart = oPortBounding.Connectable
    Else
        'Assumed as plate part for now.
    End If
    
    
    
    Dim oBddMat As IJDT4x4
    Dim oBdgMat As IJDT4x4
    Dim sX0#, sY0#, sZ0#, eX0#, eY0#, eZ0#
    Dim sX1#, sY1#, sZ1#, eX1#, eY1#, eZ1#
    oBoundedPart.Axis.EndPoints sX0, sY0, sZ0, eX0, eY0, eZ0
    
    Dim oBoundedAxis_Vector As IJDVector
    Set oBoundedAxis_Vector = New dVector
    If oBoundedAxisPort.PortIndex = SPSMemberAxisStart Then
        oBoundedPart.Rotation.GetTransformAtPosition sX0, sY0, sZ0, oBddMat, Nothing
        oBoundedAxis_Vector.Set oBddMat.IndexValue(0), oBddMat.IndexValue(1), oBddMat.IndexValue(2)
    ElseIf oBoundedAxisPort.PortIndex = SPSMemberAxisEnd Then
        oBoundedPart.Rotation.GetTransformAtPosition eX0, eY0, eZ0, oBddMat, Nothing
        oBoundedAxis_Vector.Set -oBddMat.IndexValue(0), -oBddMat.IndexValue(1), -oBddMat.IndexValue(2)
    Else
        Exit Sub
    End If
    
    Dim oBoundedLocation As IJDPosition
    Set oBoundedLocation = New DPosition
    oBoundedLocation.Set oBddMat.IndexValue(12), oBddMat.IndexValue(13), oBddMat.IndexValue(14)
    
    Dim oBdgPltPortNormal As IJDVector
    Dim oV_Vector As IJDVector
    Dim oU_Vector As IJDVector
    Set oV_Vector = New dVector
    Set oU_Vector = New dVector
    
    If Not oPortBounding Is Nothing Then
        If TypeOf oPortBounding.Connectable Is ISPSMemberPartCommon Then
            oBoundingPart.Axis.EndPoints sX1, sY1, sZ1, eX1, eY1, eZ1
            oBoundingPart.Rotation.GetTransformAtPosition eX1, eY1, eZ1, oBdgMat, Nothing
            'We need to know whether the U-vector or V-vector of bounding member need to be used as
            'the bounding surface normal: we are considering angles greater than 45 degrees
            oV_Vector.Set oBdgMat.IndexValue(8), oBdgMat.IndexValue(9), oBdgMat.IndexValue(10)
            
            Dim dBddAxisAndBdgV_DotProduct As Double
            dBddAxisAndBdgV_DotProduct = Abs(oBoundedAxis_Vector.Dot(oV_Vector))
            
            Set oU_Vector = New dVector
            oU_Vector.Set oBdgMat.IndexValue(4), oBdgMat.IndexValue(5), oBdgMat.IndexValue(6)
            
            Dim dBddAxisAndBdgU_DotProduct As Double
            dBddAxisAndBdgU_DotProduct = Abs(oBoundedAxis_Vector.Dot(oU_Vector))
            
            Dim bIsWebBoundingSurface As Boolean
            If GreaterThan(dBddAxisAndBdgU_DotProduct, dBddAxisAndBdgV_DotProduct) Then
                bIsWebBoundingSurface = True
            Else
                bIsWebBoundingSurface = False
            End If
            
            'Considering bearing plate normal to be the same as that of the bounding surface normal: here,
            ' bIsWebBoundingSurface True/False values would correspond to use oU_vector/oV_vector of the bounding respectively
            If bIsWebBoundingSurface Then
                Set oBdgPltPortNormal = oU_Vector '(of the bounding)
            Else
                Set oBdgPltPortNormal = oV_Vector '(of the bounding)
            End If
            
        ElseIf (TypeOf oPortBounding.Connectable Is SPSSlabEntity) Or (TypeOf oPortBounding.Connectable Is SPSWallPart) _
                    Or (TypeOf oPortBounding.Connectable Is IJPlate) Then
            If TypeOf oPortBounding.Geometry Is IJSurfaceBody Then
                Dim oSurfaceBdy As IJSurfaceBody
                Set oSurfaceBdy = oPortBounding.Geometry
                oSurfaceBdy.GetNormalFromPosition oBoundedLocation, oBdgPltPortNormal
            End If
        Else
            sMsg = "This Case is yet not handled"
            GoTo ErrorHandler
        End If
        
        'Check if the bounded axis and the bearing plate normal are parallel (no need to adjust flange/web thickness),
        'otherwise compute these dimensions as measured in bearing plate plane
        Dim dDotProduct As Double
        dDotProduct = Abs(oBoundedAxis_Vector.Dot(oBdgPltPortNormal))

        If GreaterThan(dDotProduct, Cos(GetPI / 2 - degreeToRadian(89))) Then
            'Angle of bounded member axis is around 89 degrees or more with respect to bearing plate surface, so
            'use Flange/Web thickness values as is which is at bottom of this 'If' block
        ElseIf GreaterThan(dDotProduct, Cos(GetPI / 2 - degreeToRadian(20))) Then
            'Angle of bounded member axis is between 20 and 89 degrees (approximately)

            '***   Third parameter (adjusted flange thickness value):
            '*** Bounded V vector ***:
            oV_Vector.Set oBddMat.IndexValue(8), oBddMat.IndexValue(9), oBddMat.IndexValue(10)
            
            Dim oResultantVec As IJDVector
            Set oResultantVec = oV_Vector.Cross(oBdgPltPortNormal)
            oResultantVec.Length = 1#
            
            'Get a unit vector which is along projection of V-Vector on to bearing plate surface:
            Dim oProjectedVvecOnBrgPlt As IJDVector
            Set oProjectedVvecOnBrgPlt = oBdgPltPortNormal.Cross(oResultantVec)
            oProjectedVvecOnBrgPlt.Length = 1#
            
            dDotProduct = Abs(oProjectedVvecOnBrgPlt.Dot(oV_Vector))
            dFlangeThickness = dFlangeThickness / dDotProduct
            
            '***   Fourth parameter (adjusted web thickness value):
            '*** Bounded U vector ***:
            oU_Vector.Set oBddMat.IndexValue(4), oBddMat.IndexValue(5), oBddMat.IndexValue(6)

            Set oResultantVec = oU_Vector.Cross(oBdgPltPortNormal)
            oResultantVec.Length = 1#

            'Get a unit vector which is along projection of U-Vector on to bearing plate surface:
            Dim oProjectedUvecOnBrgPlt As IJDVector
            Set oProjectedUvecOnBrgPlt = oBdgPltPortNormal.Cross(oResultantVec)
            oProjectedVvecOnBrgPlt.Length = 1#

            dDotProduct = Abs(oProjectedUvecOnBrgPlt.Dot(oU_Vector))
            dWebThickness = dWebThickness / dDotProduct
        Else
            sMsg = "Bounded member inclination is less than limit 20 degrees (these cases are handled)"
            GoTo ErrorHandler
        End If
    End If

    '***   Third parameter:
    pPRL.Add "FlangeThickness", dFlangeThickness
    
    '***   Fourth parameter:
    pPRL.Add "WebThickness", dWebThickness
    

  Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
End Function

' ********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
  On Error Resume Next
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  pPR.IJDInputs.RemoveAllInput
  pPR.IJDRepresentations.RemoveAllRepresentation

  Dim pDFact As New DefinitionFactory
  Dim pIH As IJDInputsHelper
  Dim pOH As IJDOutputsHelper
  
  pDFact.InitAbstractParameterRule pPR
  
  Set pIH = New InputHelper
  pIH.definition = pPR
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  
  Set pOH = New OutputHelper
  pOH.Representation = pPR.IJDRepresentations.Item(1)
  pOH.InitAs m_FamilyProgid
  ParameterRuleOutputs pOH
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_DefinitionProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

'*********************************************************************************************
' Method      : CMParameterRule
' Description :
'
'*********************************************************************************************
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPRL As IJDParameterLogic
  Set pPRL = New ParameterLogic
  pPRL.Representation = pRep
  ParameterRuleLogic pPRL
End Sub


' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
