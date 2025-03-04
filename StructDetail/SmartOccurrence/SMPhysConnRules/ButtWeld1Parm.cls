VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtWeld1Parm"
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
'
Option Explicit

Const m_ParameterRuleProgid As String = CUSTOMERID & "PhysConnRules.ButtWeld1Parm"
Const m_ParameterRuleName As String = CUSTOMERID & "PhysConnRules.ButtWeld1Parm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "\StructDetail\SmartOccurrence\" & CUSTOMERID & "PhysConnRules\ButtWeld1Parm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
    
    'Add port inputs
    pIH.SetInput INPUT_CONN_OBJECT1
    pIH.SetInput INPUT_CONN_OBJECT2
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler
    
  pOH.SetOutput "RootGap"
  pOH.SetOutput "NRRootGap"
  pOH.SetOutput "Nose"                      'User code "N"
  pOH.SetOutput "NRNose"
  pOH.SetOutput "NoseOrientationAngle"      'User code "n"
  pOH.SetOutput "NRNoseOrientationAngle"
  pOH.SetOutput "ButtCategory"
  pOH.SetOutput "RefPartName", imsARGUMENT_IS_BSTR
  pOH.SetOutput "ReferenceSide", imsARGUMENT_IS_BSTR
  pOH.SetOutput "NRReferenceSide", imsARGUMENT_IS_BSTR
  pOH.SetOutput "NonRefPartName", imsARGUMENT_IS_BSTR
  pOH.SetOutput "ReferencePart"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
    
    ' Get Class Arguments
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    Dim oPlate1 As New StructDetailObjects.PlatePart
    Dim oPlate2 As New StructDetailObjects.PlatePart
    Dim dThickness As Double
    
    Set oPhysConn.object = pPRL.SmartOccurrence
    
    'Variables to be used for getting the part names
    Dim sPartNameRef As String
    Dim sPartNameNR As String
    Dim oNamedItemRef As IJNamedItem
    Dim oNamedItemNR As IJNamedItem
    
    ' variable to indicate if plate normals are reversed
    Dim bNormalsReversed As Boolean
    bNormalsReversed = False
    
    'dim variables for bevel parameters
    Dim dNose As Double
    Dim dNRNose As Double
    
    ' keep track of the reference part number and objects
    Dim iRefPartNum As Integer
    Dim oPartRef As Object
    Dim oPartNR As Object
   
    ' *************************************************************************************
    ' create some variables to hold the "addition" values to account for any thickness
    ' differences between the parts that is not accounted for by chamfers.  There are four
    ' addition values, one for each side of each plate at the connection, but at least two
    ' of them will always be 0.  The other two will be 0 if there is no thickness change
    ' at that point or greater than 0 if this face of this part is larger than the same
    ' face on the connected part.
    '
    ' This rule will compute the overlapping thickness between the two parts.  These addition
    ' values are used to grow the appropriate nose and/or dBevel length to account for
    ' thickness changes between the parts.
    '
    ' This is currently only computed for edge connections between two plate parts.  The
    ' values will be 0 for other connections.
    ' *************************************************************************************
    Dim dAdditionBaseRef As Double
    Dim dAdditionOffsetRef As Double
    Dim dAdditionBaseNR As Double
    Dim dAdditionOffsetNR As Double
    
    ' initialize them all to 0.
    dAdditionBaseRef = 0
    dAdditionOffsetRef = 0
    dAdditionBaseNR = 0
    dAdditionOffsetNR = 0
        
    '*************************
    'get answers from selector
    '*************************
    Dim sCategory As String
    
    GetSelectorAnswer pPRL, "Category", sCategory
    
    Dim sWorkcenter As String
    
    GetSelectorAnswer pPRL, "Workcenter", sWorkcenter
    
    Dim sFirstWeldingSide As String
    
    GetSelectorAnswer pPRL, "FirstWeldingSide", sFirstWeldingSide
    
    Dim sChamferType As String
   
    GetSelectorAnswer pPRL, "ChamferType", sChamferType
    
    ' ********************************************************************
    ' determine which part is the reference and which is the non-reference
    ' ********************************************************************
    If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
        ' reference part is item 2
        iRefPartNum = 2
        Set oPartRef = oPhysConn.ConnectedObject2
        Set oPartNR = oPhysConn.ConnectedObject1
        Set oNamedItemRef = oPhysConn.ConnectedObject2
        Set oNamedItemNR = oPhysConn.ConnectedObject1
    Else
        ' reference part is item 1
        iRefPartNum = 1
        Set oPartRef = oPhysConn.ConnectedObject1
        Set oPartNR = oPhysConn.ConnectedObject2
        Set oNamedItemRef = oPhysConn.ConnectedObject1
        Set oNamedItemNR = oPhysConn.ConnectedObject2
    End If
    
    ' *************************************************
    ' set the part names for the reference and NR parts
    ' *************************************************
    sPartNameRef = oNamedItemRef.Name
    sPartNameNR = oNamedItemNR.Name
    
    ' ************************************************************************************************
    ' get the reference sides.  This returns "Base" or "Offset", depending which is the "molded" side.
    ' The answer can be different for the two parts.
    ' ************************************************************************************************
    Dim sRefSideRef As String
    Dim sRefSideNR As String
    sRefSideRef = GetRefSide(oPhysConn.object, oPartRef)
    sRefSideNR = GetRefSide(oPhysConn.object, oPartNR)
    
    ' ***************************************************************************************
    ' Get the overlapping thickness between the parts and the additions that would have to
    ' be made to either face of either part to make them align with the corresponding face
    ' of the other part.  If this is not a plate-edge to plate-edge connection, the thickness
    ' will be the thickness of the thinner part and the difference in thickness will be
    ' returned as an addition in dAdditionOffsetNR
    ' ***************************************************************************************
    GetButtWeldOverlappingThicknessAndAdditions pPRL, _
                                                oPartRef, _
                                                oPartNR, _
                                                sChamferType, _
                                                iRefPartNum, _
                                                dThickness, _
                                                dAdditionBaseRef, _
                                                dAdditionOffsetRef, _
                                                dAdditionBaseNR, _
                                                dAdditionOffsetNR, _
                                                bNormalsReversed

    
    dNose = dThickness
    
    ' **********************************************************************************
    ' Set and adjust the final nose value to accout for thickness differences
    ' **********************************************************************************
    If dAdditionBaseRef < 0.00001 And _
       dAdditionOffsetRef < 0.00001 And _
       dAdditionBaseNR < 0.00001 And _
       dAdditionOffsetNR < 0.00001 Then
        ' this covers the profile cases and any case where there is no thickness difference
        ' or all thickness differences are accounted for by bevels
        dNRNose = dNose
    Else
        ' we have offsets to consider, start with the values equal and adjust them as needed
        ' at this point, we are correct if there are no adjustments to make due to thickness steps
        dNRNose = dNose
        
        ' **************************************************************************************
        ' because we specifically set the addition values above, and set the others to zero,
        ' we do not have to check the signs of the offsets.  We can simply apply the appropriate
        ' additions.  The zeros will have no effect.
        ' This check will use only the sRefSideRef and ignore the sRefSideNR because all of the
        ' values are referenced to the reference side of the reference part
        ' **************************************************************************************
        
        dNose = dNose + dAdditionOffsetRef + dAdditionBaseRef
        dNRNose = dNRNose + dAdditionOffsetNR + dAdditionBaseNR
    End If
    
    'Set the Nose width
    pPRL.Add "Nose", dNose
    pPRL.Add "NRNose", dNRNose
    
    'Set the Gap
    Select Case sWorkcenter
      Case gsMachine1
          pPRL.Add "RootGap", 0.0005
          pPRL.Add "NRRootGap", 0.0005
      Case gsMachine2
          pPRL.Add "RootGap", 0.0006
          pPRL.Add "NRRootGap", 0.0006
    End Select '
     
    'Calculation of angles
    Dim dButtMountingAngle As Double
    Dim dNOA As Double
    Dim dNRNOA As Double
    
            'Modification is done as per TR-171305 --- manish
    ' This used to use CornerButtMountingAngle, but this method causes some
    ' data loss.  Changing to use TeeMountingAngle (TR170710)
    'dButtMountingAngle = oPhysConn.CornerButtMountingAngle
           dButtMountingAngle = oPhysConn.TeeMountingAngle
    ' the code is expecting the angle to be less than 180 degrees (PI),
    ' but it could be greater than PI because of an issue in GetMountingAngle
    ' adjust it before using it  (TR170710) --
    ' **** this should be changed when DI 171304 is implemented ****
    If (dButtMountingAngle - TOL) >= PI Then
        dButtMountingAngle = (2 * PI) - dButtMountingAngle
    End If
    
    If (dButtMountingAngle - TOL) >= 1.570796 Then
        dButtMountingAngle = PI - dButtMountingAngle
    End If
    
    If (dButtMountingAngle - TOL) >= -0.0001 And (dButtMountingAngle - TOL) <= 0.0001 Or _
        (dButtMountingAngle - TOL) >= (PI - 0.0001) And (dButtMountingAngle - TOL) <= (PI + 0.0001) Then 'no knuckle
        dNOA = 1.570796
    Else 'knuckle
        If oPhysConn.InsideFaceOfKnuckle = "Base" Then
            If sRefSideRef = "Base" Then 'the molded side and inside face of knuckle match
                dNOA = 1.570796 + dButtMountingAngle / 2
            Else 'molded side is offset
                dNOA = 1.570796 - dButtMountingAngle / 2
            End If
        Else 'the sRefSideRef face of knuckle is offset
            If sRefSideRef = "Offset" Then 'the molded side and inside face of knuckle match
                'reset the bevel angles based on the knuckle; same angle used for ref and non ref
                dNOA = 1.570796 + dButtMountingAngle / 2
            Else 'molded side is base
                dNOA = 1.570796 - dButtMountingAngle / 2
            End If
        End If
    End If
    
    dNRNOA = dNOA
        
    If sRefSideRef = "Base" Then
          pPRL.Add "ReferenceSide", "molded"
    ElseIf sRefSideRef = "Offset" Then
          pPRL.Add "ReferenceSide", "antimolded"
    Else
          pPRL.Add "ReferenceSide", sRefSideRef
    End If
    
    If sRefSideNR = "Base" Then
          pPRL.Add "NRReferenceSide", "molded"
    ElseIf sRefSideNR = "Offset" Then
          pPRL.Add "NRReferenceSide", "antimolded"
    Else
          pPRL.Add "NRReferenceSide", sRefSideNR
    End If
              
    If sCategory = gsOneSided Then
          pPRL.Add "ButtCategory", 65536
    Else
          pPRL.Add "ButtCategory", 65537
    End If
    
    ' **********************
    ' set the reference part
    ' **********************
    pPRL.Add "ReferencePart", iRefPartNum
    
    'Set nose and ref part values
    pPRL.Add "NoseOrientationAngle", dNOA
    pPRL.Add "NRNoseOrientationAngle", dNRNOA
    pPRL.Add "NonRefPartName", sPartNameNR
    pPRL.Add "RefPartName", sPartNameRef
    
    Set oPhysConn = Nothing
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Sub


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
  pIH.Definition = pPR
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
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

