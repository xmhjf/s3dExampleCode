VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InsideFaceRathol3Parm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2014, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMCornerFeatRules
'  File        : InsideRathole3Parm.cls
'
'  Description :
'       It computes the item paramaters in the context of the smart occurrence.
'
'       - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'       - Outputs defined by name the collection of parameters
'
'  Author      : Alligators
'
'  History     :
'    5/SEP/2014 - Creation
'    29/Nov/2014 - GHM  DI-259276 Updated ParameterRuleLogic()
'*********************************************************************************************

Const m_ParameterRuleProgid As String = CUSTOMERID + "CornerFeatRules.InsideFaceRathol3Parm"
Const m_ParameterRuleName As String = CUSTOMERID + "CornerFeatRules.InsideFaceRathol3Parm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\InsideFaceRathol3Parm.cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    ' Add port inputs
    pIH.SetInput INPUT_PORT1FACE
    pIH.SetInput INPUT_PORT2EDGE
    pIH.SetInput INPUT_PORT3EDGE
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    On Error GoTo ErrorHandler
  
    pOH.SetOutput "Ulength"
    pOH.SetOutput "OffsetU"
    pOH.SetOutput "OffsetV"
    pOH.SetOutput "FlangeClearance"

    pOH.SetOutput "Radius"
    pOH.SetOutput "Angle"
    
    pOH.SetOutput "CornerToTop"
    pOH.SetOutput "FlangeThickness"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number

End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    
    On Error GoTo ErrorHandler
        
    ' ---------------------------------------------------------------
    ' Get the inside overlap between bounded flange and bounding edge
    ' ---------------------------------------------------------------
    
    Dim dCornerToInside As Double
    dCornerToInside = 0#
    
    Dim oStructFeatUtils As IJSDFeatureAttributes
    Dim oFacePortObj As Object
    Dim oEdge1 As IJPort
    Dim oEdge2 As IJPort
    
    Set oStructFeatUtils = New SDFeatureUtils
    'Get Corner Feature inputs
    oStructFeatUtils.get_CornerCutInputsEx pPRL.SmartOccurrence, oFacePortObj, oEdge1, oEdge2
    
    Dim oEdgePort1 As IJStructPort
    Dim oEdgePort2 As IJStructPort
    
    Set oEdgePort1 = oEdge1
    Set oEdgePort2 = oEdge2
    
    Dim oModelBody As IJDModelBody
    Dim dMinDistance As Double
    Dim oFlangeInnerPort As IJPort
    Dim eSubPort As JXSEC_CODE
    
    'Check for Multi bounding case using EdgeIDs
    If oEdgePort1.SectionID >= e_JXSEC_MultipleBounding_5001 And oEdgePort1.SectionID <= e_JXSEC_MultipleBounding_5005 Then
        'Get Bounded Port
        If oEdgePort1.SectionID = e_JXSEC_MultipleBounding_5001 Or oEdgePort2.SectionID = e_JXSEC_MultipleBounding_5001 Then
            eSubPort = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
        Else
            eSubPort = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
        End If
        
        'Bounded Inner port
        Set oFlangeInnerPort = GetLateralSubPortBeforeTrim(oEdge1.Connectable, eSubPort)
               
        Set oModelBody = oEdge1.Geometry
        'Get distance between Corner Feature Edgeports and Bounded ports
        oModelBody.GetMinimumDistance oFlangeInnerPort.Geometry, Nothing, Nothing, dMinDistance
        dCornerToInside = dMinDistance
                
        Set oModelBody = oEdge2.Geometry
        dMinDistance = -1
        oModelBody.GetMinimumDistance oFlangeInnerPort.Geometry, Nothing, Nothing, dMinDistance
        
        'If edge1 and edge2 are equidistant from the flange inner port,then the bounded flange is thicker than the bounded port.
        'In this case the CornerToInside parameter should be negative since its drives the CornerToTop parameter.
        
        'If edge1 and edge2 are NOT equidistant from the flange inner port,then the bounded flange is thinner than the bounded port
        'and the sign of the CornerToInside parameter should be postive. If thickness of bounded and bounding ports are same,
        'considering it as positive as it does not have any impact since the CornerToInside value will be zero.

        If Equal(dCornerToInside, dMinDistance) Then
            dCornerToInside = -dMinDistance
        ElseIf GreaterThan(dCornerToInside, dMinDistance) Then
            ' dCornerToInside is already filled with required value.
        Else
            dCornerToInside = dMinDistance
        End If

    Else 'If not multi bounding case
        dCornerToInside = GetSurfaceOffsetForEdgeCornerFeature(pPRL.SmartOccurrence)
    End If
    
    ' -------------------------------------------------------------------------------------------
    ' Get bounded flange thickness and calulate distance from corner to top of the bounded flange
    ' -------------------------------------------------------------------------------------------
    ' Distance to the top of the bounded flange should be the surface offset returned below + the flange thickness + the offset
    ' between the bounding flange inner surface and the web cut.  Though it is poor coding practice, we will simplify our
    ' code by using the fact we know this offset is static 0.03mm
    Dim boundedFlangeThickness As Double
    GetFlangeThickness pPRL.SmartOccurrence, boundedFlangeThickness
    
    pPRL.Add "FlangeThickness", boundedFlangeThickness
 
    Dim cornerToTop As Double
    cornerToTop = dCornerToInside + boundedFlangeThickness + 0.00003
    
    pPRL.Add "CornerToTop", cornerToTop
    
    ' --------------------
    ' Set flange clearance
    ' --------------------
    ' We don't know how close a torch can get to the inside of the flange, but it is unlikely to be less than
    ' 2mm.  Using a very small value, such as 0.01mm causes problems in the manfucturing output (creates very thin
    ' faces in the web cut part monitor output).
    Dim flangeClearance As Double
    flangeClearance = 0.002
    
    pPRL.Add "FlangeClearance", 0.002
    
    ' ---------------------------------------------------------
    ' Set the other parameters that do not vary with the inputs
    ' ---------------------------------------------------------
    pPRL.Add "Ulength", 0.025
    pPRL.Add "OffsetU", 0.00001
    pPRL.Add "Radius", 0.025
    
    ' ------------------------------------------------------
    ' Set the angle between the bounded and bounding flanges
    ' ------------------------------------------------------
    ' 180 indicates they are parallel
    ' For now we assume 180. The angle between them is accounted for in the
    ' symbol so that this can be enhanced in the future without changing symbol inputs.
    pPRL.Add "Angle", Atn(1) * 4
    
    ' ------------------------------------------------------------------------------------------------------
    ' Calculate the setback from the inside of the bounding flange so that the contour does not get inverted
    ' ------------------------------------------------------------------------------------------------------
    ' Shifts what the user sees as the straight line segment from one side to the other.  One side will always be
    ' close to zero (0.01mm).
    ' The offset on the bounded end is 0.01 mm when the bounded flange inner surface + the flange clearance does
    ' not go as far as 0.01 below the web cut at the bounding flange inner surface.  When it meets or exceeds this,
    ' the offset is increased to avoid the line at the bounded end from becoming inverted.
    Dim offsetV As Double

    If (boundedFlangeThickness + flangeClearance) > cornerToTop Then
        offsetV = Abs(dCornerToInside - flangeClearance) + 0.00001
    Else
        offsetV = 0.00001
    End If
            
    pPRL.Add "OffsetV", offsetV
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
 
End Sub
  
' ** Start CM **
'*********************************************************************************************
' If needed Add Custom Method HERE
'*********************************************************************************************
' ** End CM **

'*********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
'*********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
  
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
    On Error Resume Next
      
    ' Remove all existing defined Input and Output (Representations) before defining the current Inputs and Outputs
    pPR.IJDInputs.RemoveAllInput
    pPR.IJDRepresentations.RemoveAllRepresentation
      
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractParameterRule pPR
    
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pPR
    pIH.InitAs m_FamilyProgid
    ParameterRuleInputs pIH
    
    Dim pOH As IJDOutputsHelper
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
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
    
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

'*********************************************************************************************
'         !!!!! End Private Code !!!!!
'*********************************************************************************************

