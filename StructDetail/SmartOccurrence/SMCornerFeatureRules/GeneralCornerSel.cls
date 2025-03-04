VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeneralCornerSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : GeneralCornerSel.cls
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
'   pnalugol - Nov 1 2011 -202656 -  Changes made to support Corner features on Slots.
'                         - Used StructDetailObjectsEx to create corner feature
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'   12 Dec 2011 - pnalugol - enhancements to fix Corner feature on slot issues 205720
'*********************************************************************************************

Const m_SelectorProgid As String = CUSTOMERID + "CornerFeatRules.GeneralCornerSel"
Const m_SelectorName As String = CUSTOMERID + "CornerFeatRules.GeneralCornerSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\GeneralCornerSel.cls"
Const PI As Double = 3.14159265

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
    Dim oLocation As IJDPosition
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
                                           
    oCornerFeature.GetLocationOfCornerFeature oLocation
    
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
    
    dAngle1 = GetAngleBetweenCornerEdgePorts(oFP, oEP1)
    dAngle2 = GetAngleBetweenCornerEdgePorts(oFP, oEP2)

    Dim oFeatureParent As Object
    Dim oDesignChild As IJDesignChild
    Dim bOnEndCutWithGap As Boolean ' Indicates if this corner feature is on end cut
    Dim dGap1 As Double
    Dim dGap2 As Double
    Dim dAngle3 As Double  ' Angle between 2 supporting ports
    
    
    bOnEndCutWithGap = False
    Set oDesignChild = pSL.SmartOccurrence
    Set oFeatureParent = oDesignChild.GetParent
    If TypeOf oFeatureParent Is IJStructFeature Then
       ' Corner feature on end cuts
       On Error Resume Next
       oCornerFeature.MeasureCornerGap 0.001, dGap1, dGap2
       If dGap1 > 0.001 Then
          bOnEndCutWithGap = True
          
          
          
          Dim oPointOnPort1 As IJDPosition
          Dim oPort1Normal As IJDVector
          Dim oPointOnPort2 As IJDPosition
          Dim oPort2Normal As IJDVector
          Dim oTopologyLocate As TopologyLocate
                
          Set oTopologyLocate = New TopologyLocate
          oTopologyLocate.GetProjectedPointOnModelBody oEdgePort1.Geometry, _
                                                       oLocation, _
                                                       oPointOnPort1, _
                                                       oPort1Normal
                                                       
          oTopologyLocate.GetProjectedPointOnModelBody oEdgePort2.Geometry, _
                                                       oLocation, _
                                                       oPointOnPort2, _
                                                       oPort2Normal
          Dim dAngle As Double
          Dim x As Double
          Dim y As Double
          Dim z As Double
          
          oPort1Normal.Get x, y, z
          oPort1Normal.Set -x, -y, -z
          
          oPort2Normal.Get x, y, z
          oPort2Normal.Set -x, -y, -z
          
          dAngle3 = oPort1Normal.Angle(oPort2Normal, oPort1Normal)
          dAngle3 = 180 - (dAngle3 * 180) / PI
          
          On Error GoTo ErrorHandler
       End If
    End If
    
    dAngle1 = CInt(dAngle1 * 180 / PI)
    dAngle2 = CInt(dAngle2 * 180 / PI)
    
    'Selection code
    If dAngle1 = 90 And dAngle2 = 90 Then
        '**********************************************
        'new code for checking the corner gap
        'first, check to see if the gap exists
        If oCornerFeature.CornerGapExists(0.003) Then
            pSL.Add "LongScallop"
            pSL.Add "LongScallop Calculate"
            pSL.Add "LongScallopWithCollar"
         Else
            pSL.Add "LongScallop"
            pSL.Add "LongScallop 50x100"
            
            pSL.Add "LongScallop Calculate"
            pSL.Add "LongScallop 50x100 Calculate"
            pSL.Add "CrackArrest"
            pSL.Add "LongScallopWithCollar"
        End If
    Else
        pSL.Add "LongScallop Along Corner"
        pSL.Add "LongScallop Along Corner 50x100"
        pSL.Add "CrackArrest Along Corner"
        pSL.Add "CrackArrest"
        pSL.Add "LongScallop"
        pSL.Add "LongScallop 50x100"
        
        pSL.Add "LongScallop Along Corner Calculate"
        pSL.Add "LongScallop Along Corner 50x100 Calculate"
        pSL.Add "LongScallop Calculate"
        pSL.Add "LongScallop 50x100 Calculate"
        pSL.Add "LongScallopWithCollar"
    End If
    
    If bOnEndCutWithGap = True Then
       dGap1 = Round(dGap1, 3)
       If dGap1 >= 0.015 Then
          If dAngle3 <= 90 Then
             pSL.Add "OffsetRadius"
             pSL.Add "OffsetRadius With Straight Part"
          Else
             pSL.Add "OffsetRadius Obtuse"
             pSL.Add "OffsetRadius Obtuse With Straight Part"
          End If
       ElseIf dGap1 >= 0.005 Then
          ' If bounding profile penetrates tight plate,the slot might have a
          ' bulding out corner at this corner feature location,dGap1 is
          ' be dependent that corner radius
          If dAngle3 <= 90 Then
             pSL.Add "OffsetRadius With Straight Part"
          Else
             pSL.Add "OffsetRadius Obtuse With Straight Part"
          End If
       End If
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

