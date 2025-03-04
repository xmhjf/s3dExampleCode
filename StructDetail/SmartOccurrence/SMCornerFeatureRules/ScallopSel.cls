VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScallopSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : ScallopSel.cls
'
'  Description :
'
'   Selector:
'   ---------
'   It selects the list of available Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parameterize the selection
'
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'
'*********************************************************************************************
'   pnalugol - Nov 1 2011 - 202656 - Changes made to support Corner features on Slots.
'                         - Used StructDetailObjectsEx to create corner feature
'   12 Dec 2011 - pnalugol - enhancements to fix Corner feature on slot issues 205720
Const m_SelectorProgid As String = CUSTOMERID + "CornerFeatRules.ScallopSel"
Const m_SelectorName As String = CUSTOMERID + "CornerFeatRules.ScallopSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\ScallopSel.cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : SelectorInputs
' Description :
'
'*********************************************************************************************
Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
'Add port inputs
pIH.SetInput INPUT_PORT1FACE
pIH.SetInput INPUT_PORT2EDGE
pIH.SetInput INPUT_PORT3EDGE
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
 
End Sub

'*********************************************************************************************
' Method      : SelectorQuestions
' Description :
'
'*********************************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add your question below
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
 
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description :
'
'*********************************************************************************************
Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    'Get corner feature
    Dim oCornerFeature As IJSDOCornerFeature
    Set oCornerFeature = New StructDetailObjectsex.CornerFeature
    Set oCornerFeature.object = pSL.SmartOccurrence
    
    'Get corner feature part object
    Dim oPart As Object
    Set oPart = oCornerFeature.GetPartObject
    
    'Get Corner location and angle of edge ports
    Dim dAngle1 As Double
    Dim dAngle2 As Double
    
    Dim oStructFeatUtils As IJSDFeatureAttributes
    Dim oFacePort As Object
    Dim oEdgePort1 As Object
    Dim oEdgePort2 As Object
    
    Set oStructFeatUtils = New SDFeatureUtils
      
    oStructFeatUtils.get_CornerCutInputsEx pSL.SmartOccurrence, _
                                          oFacePort, _
                                         oEdgePort1, _
                                         oEdgePort2
                                               
    Dim oFP As IJPort, oEP1 As IJPort, oEP2 As IJPort
    
    If TypeOf oFacePort Is IJPort Then
        Set oFP = oFacePort
    End If
    
    If TypeOf oEdgePort1 Is IJPort Then
        Set oEP1 = oEdgePort1
    Else
        GetPortFromProxy oEdgePort1, oCornerFeature.object, oEP1
    End If
    
    
    If TypeOf oEdgePort2 Is IJPort Then
        Set oEP2 = oEdgePort2
    Else
        GetPortFromProxy oEdgePort2, oCornerFeature.object, oEP2
    End If
    
    If oEP1 Is Nothing Or oEP2 Is Nothing Then
        pSL.Add "Scallop Along Corner"
        Exit Sub
    End If
    
    dAngle1 = GetAngleBetweenCornerEdgePorts(oFP, oEP1)
    dAngle2 = GetAngleBetweenCornerEdgePorts(oFP, oEP2)
    
    'Selection code
    If CInt(dAngle1 * 180 / 3.14156) = 90 And CInt(dAngle2 * 180 / 3.14156) = 90 Then
        pSL.Add "Scallop"
        pSL.Add "Scallop R50"
        pSL.Add "ScallopWithCollar"
    Else
        pSL.Add "Scallop Along Corner"
        pSL.Add "Scallop Along Corner R50"
        pSL.Add "Scallop"
        pSL.Add "Scallop R50"
        pSL.Add "ScallopWithCollar"
    End If
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
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
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
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
' Method      : CMSelector
' Description :
'
'*********************************************************************************************
Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

