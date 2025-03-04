VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SingleBaseParm"
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

Const m_ParameterRuleProgid As String = CUSTOMERID + "ChamferRules.SingleBaseParm"
Const m_ParameterRuleName As String = CUSTOMERID + "ChamferRules.SingleBaseParm"
Const m_FamilyProgid As String = ""
Const PI As Double = 3.141592654

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "ChamferRules\SingleBaseParm.cls"

Implements IJDUserSymbolServices

Public Enum ChamferCondition
        CanCylinderToCone = 1
        CanConeToCylinder = 2
        CanConeToMember = 4
        CanCylinderToMember = 8
        Standard = 16
       
End Enum
     

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

    pIH.SetInput CHAMFER_PART
    pIH.SetInput OPPOSITE_PART
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

  pOH.SetOutput "Depth1Base"
  pOH.SetOutput "AngleRadius1Base"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler

    Dim lStatusCode As Long
    Dim bBaseToBase As Boolean
    Dim dBase As Double
    Dim dChamferBase As Double
    Dim dOppositeBase As Double
    Dim dChamferOffset As Double
    Dim dOppositeOffset As Double
    Dim dChamferThickness As Double
    Dim dAngle As Double
    Dim dEdgeAngle As Double
    Dim dBaseDelta As Double
    Dim dOffsetDelta As Double
    
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    
    Dim oAssyConn As New StructDetailObjects.AssemblyConn
    Dim oAssemblyConnection As IJAssemblyConnection
    Dim dThicknessDiff As Double
    
    Dim oRootParent1 As Object
    Dim oRootParent2 As Object
    
    Dim oChamfer As StructDetailObjects.Chamfer
    
    'Initialize Chamfer Part/Ports
    Set oChamfer = New StructDetailObjects.Chamfer
    Set oChamfer.object = pPRL.SmartOccurrence
    
    Dim oParentOccurance As IJSmartOccurrence
    Dim strChamferMeasurement As String
    Dim dThickDifference As Double
    
    'Get Chamfer Parent Occurance
    GetSmartOccurrenceParent pPRL.SmartOccurrence, oParentOccurance
    
    'Check if Chamfer is created by StructAssemblyConnection
    If TypeOf oParentOccurance Is IJStructAssemblyConnection Then
        
        'Get the Chamfer Value from Parent AC Sector Question Answer
        GetSelectorAnswer oParentOccurance, "ChamferMeasurement", strChamferMeasurement
        GetSelectorAnswer oParentOccurance, "ChamferValue", dAngle
        
        Dim oAppConnection As IJAppConnection
        Set oAppConnection = oParentOccurance

        Dim oPenetratingPort As IJPort
        Dim oPenetratedPort As IJPort
        GetPenetratedAndPenetratingPorts oAppConnection, oPenetratedPort, oPenetratingPort

        
        'Calculate Chamfer Angle based on Chamfer Measurement type
        If strChamferMeasurement = "Slope" Then
            pPRL.Add "AngleRadius1Base", PI / 2 - Atn(dAngle)
        Else
            pPRL.Add "AngleRadius1Base", PI / 2 - (dAngle * PI / 180)
        End If
        
        'Get Thickenss diff between Plate and Member Flange
        GetThicknessDiffOfPlateOverMbrFlange oPenetratedPort.Connectable, oPenetratingPort.Connectable, dThickDifference
        
        pPRL.Add "Depth1Base", dThickDifference
      Exit Sub
    ElseIf TypeOf oParentOccurance Is IJStructFeature Then
    
        Dim oSmartParent As IJSmartOccurrence
        GetSmartOccurrenceParent oParentOccurance, oSmartParent
        
        If TypeOf oSmartParent Is IJAssemblyConnection Then
            Set oAssyConn = New StructDetailObjects.AssemblyConn
            Set oAssemblyConnection = oSmartParent
            Set oAssyConn.object = oAssemblyConnection
            Dim oProfilePart1 As StructDetailObjects.ProfilePart
            Dim oProfilePart2 As StructDetailObjects.ProfilePart
            Set oProfilePart1 = New StructDetailObjects.ProfilePart
            Set oProfilePart2 = New StructDetailObjects.ProfilePart
            Set oPort1 = oAssyConn.Port1
            Set oPort2 = oAssyConn.Port2
            Set oProfilePart1.object = oPort1.Connectable
            Set oProfilePart2.object = oPort2.Connectable
            
            Dim bFromKnuckle As Boolean
            Dim dKnuckleAngle As Double
            Dim oProfileKnuckle As IJProfileKnuckle
            bFromKnuckle = oAssyConn.FromKnuckle(oProfileKnuckle)
            Dim dInclined As Double
            dInclined = 1#
            
            If not oProfileKnuckle Is Nothing Then
                dKnuckleAngle = (oProfileKnuckle.Angle / PI) * 180
                Dim dAngle1 As Double
                dAngle1 = (180 - dKnuckleAngle) / 2
                dInclined = Cos(dAngle1 * PI / 180)
            End If
            
            dThicknessDiff = Abs(oProfilePart1.WebThickness - oProfilePart2.WebThickness) / Abs(dInclined)
            
            pPRL.Add "Depth1Base", dThicknessDiff
            pPRL.Add "AngleRadius1Base", DegToRad(72)
            Exit Sub
        ElseIf TypeOf oSmartParent Is IJFreeEndCut Then
            Dim oEndCut As IJFreeEndCut
            Set oEndCut = oSmartParent
            Dim oBoundedObject As Object
            Dim oBoundingObject As Object
            oEndCut.get_FreeEndCutInputs oBoundedObject, oBoundingObject
            If TypeOf oBoundedObject Is IJPort Then
                Dim oPort As IJPort
                Set oPort = oBoundedObject
                
                Dim oBoundedProfile As IJConnectable
                Set oBoundedProfile = oPort.Connectable
                If TypeOf oBoundedProfile Is IJStiffener Then
                    Dim oProfile As StructDetailObjects.ProfilePart
                    Set oProfile = New StructDetailObjects.ProfilePart
                    Set oProfile.object = oBoundedProfile
                    Dim dERThickness As Double
                    dERThickness = oProfile.WebThickness
                    
                    Set oProfile = Nothing
                    Set oBoundedProfile = Nothing
                End If
            End If
                
             pPRL.Add "Depth1Base", dERThickness / 2
             pPRL.Add "AngleRadius1Base", DegToRad(72)
             Exit Sub
        End If
    End If
    Dim oDesignChild As IJDesignChild
    
    Dim bBoolean As Boolean
    Dim bChamferSlope As Double
    Dim oConnectable1 As IJConnectable
    Dim oConnectable2 As IJConnectable
    Dim dAngle_unknown As Double
    Dim oShpStrConnProperties As IJShpStrConnProperties
    Dim dSlope As Double
    Dim mOffset As Double
       
    dSlope = DegToRad(72)
    Set oDesignChild = oChamfer.object
    Set oAssemblyConnection = oDesignChild.GetParent
    Set oAssyConn.object = oAssemblyConnection
    Set oShpStrConnProperties = oAssyConn.object
    Set oPort1 = oChamfer.Chamfered
    Set oPort2 = oChamfer.DrivesChamfer
    
    ' Check for Plate Edge being "Chamfered" by Profile (Stiffener, Edge Reinforcement) Edge
    Set oConnectable1 = oPort1.Connectable
    Set oConnectable2 = oPort2.Connectable
    If Not TypeOf oConnectable1 Is IJPlatePart Then
    ElseIf Not TypeOf oConnectable2 Is IJPlatePart Then
        PlateEdgeByStiffenerEdgeChamferData oPort1, oPort2, dChamferBase, dChamferOffset
'''MsgBox "SMChamferRules.SingleBaseParm::ParameterRuleLogic" & vbCrLf & _
'''       "dChamferBase   = " & Format(dChamferBase, "0.0000") & vbCrLf & _
'''       "dChamferOffset = " & Format(dChamferOffset, "0.0000")
        
        ' Know that a Chamfer Depth of 0.0 is not valid
        If dChamferBase < 0.001 Then
            dChamferBase = 0.001
        End If
        
        pPRL.Add "Depth1Base", dChamferBase
        pPRL.Add "AngleRadius1Base", dSlope
        Exit Sub
    End If
    
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oChamferPart As StructDetailObjects.PlatePart
    Dim oOppositePart As StructDetailObjects.PlatePart
    
    Set oChamferPart = New StructDetailObjects.PlatePart
    Set oOppositePart = New StructDetailObjects.PlatePart
    Set oChamferPart.object = oChamfer.ChamferedPart
    Set oOppositePart.object = oChamfer.DrivesChamferPart
    
    dThicknessDiff = oChamferPart.PlateThickness - oOppositePart.PlateThickness

    Dim oChamferCondition As ChamferCondition
    
    If pPRL.SmartItem.Name = "SingleSidedBaseTeePC" Then
        'this is a tee base chamfer
        GetSelectorAnswer pPRL, "ChamferThickness", dChamferThickness
        dBase = oChamferPart.PlateThickness - dChamferThickness
    
    ElseIf pPRL.SmartItem.Name = "SingleSidedBasePC" Or _
           pPRL.SmartItem.Name = "SingleSidedBase" Then
  
        ' check if Chamfer Plate Parts are Leaf Plate Parts from the same Root
        Set oSDO_Helper = New StructDetailObjects.Helper
        Set oRootParent1 = oSDO_Helper.Object_RootParentSystem(oChamferPart.object)
        Set oRootParent2 = oSDO_Helper.Object_RootParentSystem(oOppositePart.object)
         
        If oRootParent1 Is oRootParent2 Then
            ' Chamfer Plate Parts are Leaf Plate Parts from the same Root
            'for "SingleSidedBasePC" cases: Determine the Base depth
            'for "SingleSidedBase" cases: Determine the Base depth
            
            'we know we want a chamfer on our base side, just need to get the value
            'get the offsets on ourself from molded to base and offset
            'get the offsets on the opposite part, using us as a reference
            'compute the deltas for base and offset, relative to our base and offset
            'if the chamfer is on the base side, our base delta should be >0 (expected)
            'compute the delta on the offset side with the opposite sign so that a chamfer on
            'our part would have a delta >0
            'use the larger value, which we would expect to be the base delta
            'if this value is < .0001, use .0001
            
            'check to see it is a CANBYCANCASE
            oChamferCondition = GetChamferCondition(oChamferPart, oOppositePart, bChamferSlope)
            If oChamferCondition = CanCylinderToCone Then
                dSlope = Atn(bChamferSlope) 'dSLope is ChamferSlope
                oShpStrConnProperties.GetMountingAngle oPort1, oPort2, oChamferPart, dAngle, dAngle_unknown
                If dAngle > PI Then
                    dAngle = ((2 * PI) - dAngle)
                End If
                dBase = dThicknessDiff / (Cos(dAngle / 2))
                GetAdjustmentValue dSlope, dAngle, dThicknessDiff, mOffset
                If dSlope > ((PI / 2) - dAngle) Then
                dBase = dBase - mOffset
                Else
                dBase = dBase + mOffset
                End If
               
                pPRL.Add "Depth1Base", dBase
                pPRL.Add "AngleRadius1Base", dSlope
            Exit Sub
            ElseIf oChamferCondition = CanConeToCylinder Then
                oShpStrConnProperties.GetMountingAngle oPort1, oPort2, oChamferPart, dAngle, dAngle_unknown
                If dAngle > PI Then
                    dAngle = ((2 * PI) - dAngle)
                End If
                dSlope = (PI / 2) - dAngle
            ElseIf oChamferCondition = Standard Then
                dSlope = DegToRad(72)
            End If
            
            'get our offsets
            dChamferBase = oChamferPart.OffsetToBaseFace
            dChamferOffset = oChamferPart.OffsetToOffsetFace
            
            'get opposite offsets, relative to ours
            dOppositeBase = oOppositePart.OffsetToBaseFace(oChamferPart)
            dOppositeOffset = oOppositePart.OffsetToOffsetFace(oChamferPart)
            
            dBaseDelta = dOppositeBase - dChamferBase       'compute this forward so that >0 means chamfer is on us
            dOffsetDelta = dChamferOffset - dOppositeOffset 'compute this backward so that >0 means chamfer is on us
           
            If dBaseDelta > dOffsetDelta Then
                dBase = dBaseDelta
            Else
                dBase = dOffsetDelta
            End If
            
            If dBase < 0.0001 Then
                dBase = 0.0001
            End If
        Else
            oChamferCondition = GetChamferCondition(oChamferPart, oOppositePart, bChamferSlope)
            'case to handle when there is no cone. when chamfer is between Tube portion of Can and
            'attached designed member.
            If oChamferCondition = CanCylinderToMember Then
                dSlope = Atn(bChamferSlope) 'dont need to calculate dBase as for CANS.
            ElseIf oChamferCondition = CanConeToMember Then
                oShpStrConnProperties.GetMountingAngle oPort1, oPort2, oAssyConn.ConnectedObject1, dAngle, dAngle_unknown
                If dAngle > PI Then
                    dAngle = ((2 * PI) - dAngle)
                End If
                dSlope = (PI / 2) - dAngle
            ElseIf oChamferCondition = Standard Then
                dSlope = DegToRad(72)
            End If
            
            ' Chamfer Plate Parts are not from the same Root Plate System
            ' Calculate thickness differences using the Chamfer Sketch Plane
            ' (This might have an Impact on Performance)
            
            oChamfer.GetChamferBaseOffsetData dBaseDelta, dOffsetDelta, _
                                              bBaseToBase, dEdgeAngle, lStatusCode
            If lStatusCode <> 0 Then
                dBase = 0#
            Else
                ' The Single Sided Chamfer uses "girth" Constraints for Offsets
                ' Need to adjust the Vertical distance by the Edge Angle
                dBase = Abs(dBaseDelta / dEdgeAngle)
            End If
        End If


    Else
        'for (UnKnown) cases: Determine the Offset depth
        dBase = 0.0001
    End If
    
    'Make sure Base is not 0.  If so, this parameter is being
    'updated before the symbol will be deleted due to assoc
    'sequence (or there is an error in the rules)
    If dBase < 0.0001 Then dBase = 0.0001
    
    'Set Depth1Base parameter
    pPRL.Add "Depth1Base", dBase
    pPRL.Add "AngleRadius1Base", dSlope
    
  Set oChamfer = Nothing
  Set oChamferPart = Nothing
  Set oOppositePart = Nothing
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
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
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

