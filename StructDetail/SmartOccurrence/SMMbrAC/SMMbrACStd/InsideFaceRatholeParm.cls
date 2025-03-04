VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InsideFaceRatholeParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2014, Intergraph Corporation.  All rights reserved.
'
'  File        : CornerToInsideCFParm.cls
'
'  Author      : Alligators
'
'  History     :
'
'*********************************************************************************************

Const m_sClassName As String = "InsideFaceRatholeParm"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sStdACProjectName + "." + m_sClassName
Const MODULE = m_sStdACProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

    'Add port inputs
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
  
    pOH.SetOutput "OffsetV"
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

    dCornerToInside = GetSurfaceOffsetForEdgeCornerFeature(pPRL.SmartOccurrence)
    
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
        offsetV = Abs(cornerToTop - boundedFlangeThickness - flangeClearance) + 0.00001
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
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionProgid
  
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
  Dim oPRL As IJDParameterLogic
  Set oPRL = New ParameterLogic
  oPRL.Representation = pRep
  ParameterRuleLogic oPRL

End Sub

'*********************************************************************************************
'         !!!!! End Private Code !!!!!
'*********************************************************************************************

