VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCutERParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Parameter Rule:
'   ---------------
'   It computes the item paramaters in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Outputs defined by name the collection of parameters
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
Option Explicit

Const m_ParameterRuleProgid As String = CUSTOMERID + "EndCutRules.WebCutERParm"
Const m_ParameterRuleName As String = CUSTOMERID + "EndCutRules.WebCutERParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\WebCutERParm.cls"

Implements IJDUserSymbolServices
Implements IJEndCutDrawingType

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

  pOH.SetOutput "CutDepth"
  
  '****following added for FreeEnd functionality
  pOH.SetOutput "SetBack"
  pOH.SetOutput "TopSetBack"
  pOH.SetOutput "BottomSetBack"
  '****end code added for free end functionality
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler

    Dim strError As String
    Dim oBoundedObject As Object
    Set oBoundedObject = pPRL.InputObject(INPUT_BOUNDED)
    If (oBoundedObject Is Nothing) Then
        strError = "pPRL.InputObject(INPUT_BOUNDED) is NOTHING"
        GoTo ErrorHandler
    End If
    
    ' get Bounded Part
    Dim oBoundedPart As Object
    Dim oBoundedPort As IJPort
    If TypeOf oBoundedObject Is IJPort Then
        Set oBoundedPort = oBoundedObject
        Set oBoundedPart = oBoundedPort.Connectable
    Else
        ' error, BoundedObject MUST be a IJPort object
        strError = "BoundedObject MUST be a IJPort object"
        GoTo ErrorHandler
    End If
     
     
     Dim phelper As New StructDetailObjects.Helper
     
    'get the end cut type answer from the WebCutSel rule
    Dim strEndCutType As String
    'TR-169311 Trying to retrieve question's answer without
    'hardcoding the path
    GetSelectorAnswer pPRL, QUES_ENDCUTTYPE, strEndCutType
    
    Dim TopSetBack As Double
    Dim BottomSetBack As Double
    
    Dim sTypeObject As String
    Dim sObjectType As String
    phelper.GetObjectTypeData oBoundedPart, sTypeObject, sObjectType
    
    'If the Bounded Profile an Edge Reinforcement
    If sTypeObject = "EdgeReinforcement" Then
        Select Case strEndCutType
            Case "Snip"
                'If the profile is a Flat Bar
                Dim oProfile As New StructDetailObjects.ProfilePart
                Set oProfile.object = oBoundedPart
                If oProfile.SectionType = "FB" Then
                    'Is the Edge Reinforcement mounted On Face
                    'then treat it as a normal stiffener
                    Dim oEdgeReinforcement As New StructDetailObjects.EdgeReinforcement
                    Set oEdgeReinforcement.object = oBoundedPart
                    'Edge Reinforcement is on edge
                    If Not oEdgeReinforcement.ReinforcementPosition = "OnFace" Then
                        'Calculate TopSetBack and BottomSetBack
                        
                        Dim oStructPort As IJStructPort
                        Set oStructPort = oBoundedPort
                        
                        Dim ctx As eUSER_CTX_FLAGS
                        ctx = oStructPort.ContextID
                        
                        ' If not from a split, this has been extended
                        ' Use a position where the boundary intersects
                        ' the ER, instead of the bounded port, which has
                        ' been offset a significant amount
                        Dim oObjForPlateTopBottom As Object
                        
                        If ctx And CTX_3DSPLIT Then
                            Set oObjForPlateTopBottom = oBoundedPort
                        Else
                            ' Get the object that created the end before it was extended
                            Dim opt As Long
                            Dim oPR As Long
                            
                            opt = oStructPort.OperationID
                            oPR = oStructPort.OperatorID
                            
                            Dim oDetailHelper As IJStructDetailHelper
                            Set oDetailHelper = New StructDetailHelper
                            
                            Dim oOperation As IJStructOperation
                            Dim oOperator As Object
                            
                            oDetailHelper.FindOperatorForOperationInGraphByID oBoundedPart, _
                                                                              opt, oPR, _
                                                                              oOperation, _
                                                                              oOperator
                                                                              
                            Set oDetailHelper = Nothing
                                                                              
                            ' If no operator, use the landing curve endpoint
                            If oOperator Is Nothing Then
                                Dim oProfileUtil As IJProfileAttributes
                                Set oProfileUtil = New ProfileUtils
                                
                                Dim oBoundedIJProfile As IJProfile
                                Set oBoundedIJProfile = oBoundedPart
                                Dim oLandCurve As Object
                                Dim oLandWire As IJWireBody
                                
                                oProfileUtil.GetLandingCurveFromProfile oBoundedIJProfile, _
                                                                        oLandCurve
                                
                                Dim oStart As IJDPosition
                                Dim oEnd As IJDPosition
                                
                                Set oLandWire = oLandCurve
                                
                                oLandWire.GetEndPoints oStart, oEnd
                                
                                If ctx And CTX_BASE Then
                                    Set oObjForPlateTopBottom = oStart
                                Else
                                    Set oObjForPlateTopBottom = oEnd
                                End If
                                
                                Set oStart = Nothing
                                Set oEnd = Nothing
                                Set oLandWire = Nothing
                                Set oLandCurve = Nothing
                                Set oBoundedIJProfile = Nothing
                                Set oProfileUtil = Nothing
                            ' If operator found
                            Else
                                ' Get a model representing the boundary
                                Dim oBoundaryModel As IJModelBody
                                Set oBoundaryModel = oOperator
                                
                                If oBoundaryModel Is Nothing Then
                                
                                    Dim oGeomMisc As IJGeometryMisc
                                    Dim oAggr As Object
                                    Dim NullObj As Object
                                    Dim oModelFromGType As Object
    
                                    Set oGeomMisc = New DGeomOpsMisc
                                    
                                    oGeomMisc.CreateModelGeometryFromGType NullObj, _
                                                                           oOperator, _
                                                                           NullObj, _
                                                                           oModelFromGType
                                                                             
                                    Set oBoundaryModel = oModelFromGType

                                    Set oGeomMisc = Nothing
                                    Set oModelFromGType = Nothing
                                End If
                                
                                ' Get any point on the port
                                Dim oModelHelper As IJSGOModelBodyUtilities
                                Set oModelHelper = New SGOModelBodyUtilities
                               
                                Dim oVertexColl As Collection
                                oModelHelper.GetVertices oBoundaryModel, oVertexColl
                                
                                Dim oPointOnPort As IJDPosition
                                Set oPointOnPort = oVertexColl.Item(1)
                                
                                ' Find the closest point on the boundary
                                
                                Dim distance As Double
                                Dim oPointOnBoundary As IJDPosition
                                oModelHelper.GetClosestPointOnBody oBoundaryModel, _
                                                                   oPointOnPort, _
                                                                   oPointOnBoundary, _
                                                                   distance
                                
                                Set oObjForPlateTopBottom = oPointOnBoundary
                                
                                Set oVertexColl = Nothing
                                Set oModelHelper = Nothing
                                Set oPointOnPort = Nothing
                                Set oPointOnBoundary = Nothing
                                Set oBoundaryModel = Nothing
                                
                            End If ' boundary found
                        End If ' a split port
                        
                        Set oStructPort = Nothing
                        
                        oEdgeReinforcement.GetPlateToTopBottom _
                                                        oObjForPlateTopBottom, _
                                                        TopSetBack, _
                                                        BottomSetBack
                        
                        Set oObjForPlateTopBottom = Nothing
                        
                        ' If the TopSetBack and the BottomSetBack value is bigger than
                        ' the web length, it is likely caused by some configuration below:
                        ' 1. The end cut is a free end cut. Not bounded.
                        ' 2. The end cut is exactly at the corner.
                        
                        ' For this problem, the good workaround would be bounding the ER back
                        ' by a very small distance. (5mm)
                        
                        ' do not want a zero(0) or negative SetBack value
                        If TopSetBack < 0.00001 Or TopSetBack > oProfile.WebLength Then
                            TopSetBack = (oProfile.WebLength / 4#)
                        End If
                        
                        If BottomSetBack < 0.00001 Or BottomSetBack > oProfile.WebLength Then
                            BottomSetBack = (oProfile.WebLength / 4#)
                        End If
                        
                        pPRL.Add "TopSetBack", TopSetBack
                        pPRL.Add "BottomSetBack", BottomSetBack
                    End If
                End If
        End Select
    End If
    
    'SetBack default to 35mm
    Dim dSetBack As Double
    dSetBack = 0.035
        
    'Get Bounded part to determine value of "SetBack"
    Dim oSDO_Bounded As Object
    If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        Dim oBoundedProfile As New StructDetailObjects.ProfilePart
        Set oBoundedProfile.object = oBoundedPart
        Set oSDO_Bounded = oBoundedProfile
        
        '$$$ Free End functionality
        Dim bFreeEndCut As Boolean
        Dim oWebCut As StructDetailObjects.WebCut
        Set oWebCut = New StructDetailObjects.WebCut
        Set oWebCut.object = pPRL.SmartOccurrence
        bFreeEndCut = oWebCut.IsFreeEndCut
        
        If bFreeEndCut Then
            Select Case strEndCutType
                Case "Snip"  'end cut type is F  (free)
                    'Profile is Not bounded by Connectables
                    'but if
                    'Edge Reinforcement bounded by Ref.Planes or Assoc. Points
                    If Not oWebCut.IsUnBounded Then
                        'SetBack default to .01mm
                        dSetBack = 0.00001
                    End If
            End Select
        End If
            
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        Dim oBoundedBeam As New StructDetailObjects.BeamPart
        Set oBoundedBeam.object = oBoundedPart
        Set oSDO_Bounded = oBoundedBeam
      
    Else
        ' error, BoundedObject MUST be a Stiffener or Beam
        strError = "BoundedObject MUST be a Stiffener or Beam"
        GoTo ErrorHandler
    
    End If
        
    pPRL.Add "SetBack", dSetBack
    
    'define the cutting depth as four(4) times the Web Thickness
    Dim dCuttingDepth As Double
    dCuttingDepth = GetBoundedCutDepth(oBoundedPart, False) * 3#
        

    pPRL.Add "CutDepth", dCuttingDepth

  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleLogic", strError).Number
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
  IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
  On Error Resume Next
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
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
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPRL As IJDParameterLogic
  Set pPRL = New ParameterLogic
  pPRL.Representation = pRep
  ParameterRuleLogic pPRL
End Sub
'**********************************************************
' Method Name   : IJEndCutDrawingType_UpdateDrawingType
' Inputs        : feature occurence object
' Output        : string, returns the end cut type.
' Description   : this method checks for the type of endcut, straight or sniped and offset or no-offset
'                and assigns the corresponding enum to a string variable.
'*********************************************************
Public Function IJEndCutDrawingType_UpdateDrawingType(ByVal pSmartOccurrence As Object) As String
    On Error GoTo ErrorHandler
    Dim dSetBack As Double
    Dim dTopSetBack As Double
    Dim dBottomSetBack As Double
    
    dSetBack = GetCustomAttribute(pSmartOccurrence, "IJUAWebCuts", "SetBack")
    dTopSetBack = GetCustomAttribute(pSmartOccurrence, "IJUAERWebCuts", "TopSetBack")
    dBottomSetBack = GetCustomAttribute(pSmartOccurrence, "IJUAERWebCuts", "BottomSetBack")
    
    Dim sDrawingtype As String
    
    If GreaterThan(dSetBack, 0.004) Then
        ' Offset
        If GreaterThan(dTopSetBack, 0.004) Or GreaterThan(dBottomSetBack, 0.004) Then
            sDrawingtype = Sniped_Offset_WebCuts
        Else
            sDrawingtype = Straight_Offset_WebCuts
        End If
    Else
         sDrawingtype = Sniped_No_Offset_WebCuts
    End If
    
    
    
    AddFeatureEndCutData pSmartOccurrence, EndCutRelativePosition.Primary, sDrawingtype
    IJEndCutDrawingType_UpdateDrawingType = sDrawingtype
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJEndCutDrawingType_UpdateDrawingType").Number
End Function
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************




