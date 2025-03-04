VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VVADASym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:00:03 AM"
'******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    VVADASym.cls
'
'Author
'    RM
'
'Description
'    Definition of Variable Voulme Air Distribution Assembly.
'
'Notes
'
'    <notes>
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'*******************************************************************
Option Explicit

Dim m_outputColl As IJDOutputCollection

Const NEGLIGIBLE_THICKNESS = 0.0001


Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DAirDistribAssemblyAsm.VVADASym"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DAirDistribAssemblyAsm.VVADASym"
  oSymbolDefinition.CodeBase = CodeBase
  
  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId
  
  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  Set oSymbolDefinition = Nothing
  
  Exit Function

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Function


Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    ' Feed GSCADAirDistribAssembly.VVADA Definition
    ' Inputs:
    '          1. "Part"  ( Catalog Part )
    '          2. "BasicAsmWidth"
    '          3. "BasicAsmHeight"
    '          4. "BasicAsmLength"
    '          5. "HeatingCoilLength"
    '          6. "MultiOutletLength"
    '          7. "InletDia"
    '          8. "MultiOutletDia"
    '          9. "NozzleLength1"
    '         10. "NozzleLength2"
    '         11. "NozzleLength3"
    '
    ' The representations "Physical" and "Envelope" are set to the definition
    ' Physical :    (RepresentationId = 1)
    ' Envelope :    (RepresentationId = 128)
  
    On Error GoTo ErrorHandler
  
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition
    
    ' Create a new input by new operator
    Dim Inputs(1 To 7) As IMSSymbolEntities.IJDInput
    Dim Index As Integer
  
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue

    Set Inputs(1) = New IMSSymbolEntities.DInput
    Inputs(1).Name = "BasicAsmWidth"
    Inputs(1).Description = "Width of the Basic Assembly"
    Inputs(1).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.3048
    Inputs(1).DefaultParameterValue = PC
    
    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).Name = "BasicAsmHeight"
    Inputs(2).Description = "Height of the basic Assembly"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.2032
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).Name = "BasicAsmLength"
    Inputs(3).Description = "Length of the basic assembly"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.3239
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).Name = "HeatingCoilLength"
    Inputs(4).Description = "Length of the heating coil"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Inputs(4).DefaultParameterValue = PC
    
    Set Inputs(5) = New IMSSymbolEntities.DInput
    Inputs(5).Name = "MultiOutletLength"
    Inputs(5).Description = "Length of the multi outlet"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.9
    Inputs(5).DefaultParameterValue = PC
    
    Set Inputs(6) = New IMSSymbolEntities.DInput
    Inputs(6).Name = "InletDia"
    Inputs(6).Description = "Inlet Diameter"
    Inputs(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.0984
    Inputs(6).DefaultParameterValue = PC

    Set Inputs(7) = New IMSSymbolEntities.DInput
    Inputs(7).Name = "MultiOutletDia"
    Inputs(7).Description = "MultiOutlet Diameter"
    Inputs(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1492
    Inputs(7).DefaultParameterValue = PC

    For Index = 1 To 7
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 4) As IMSSymbolEntities.IJDOutput
    
    For Index = 1 To 4
        Set O(Index) = New IMSSymbolEntities.DOutput
        O(Index).Properties = 0
    Next

    O(1).Name = "BasicAsmHVACNozzle"
    O(1).Description = "HvacPort on the basic assembly"

    O(2).Name = "BasicAssembly"
    O(2).Description = "Box representing the Basic Assembly"

    O(3).Name = "HeatingCoil"
    O(3).Description = "Box representing the Heating Coil"
    
    O(4).Name = "MultiOutlet"
    O(4).Description = "Box representing the Multi-Outlet"

    'Define the representation "Physical"
    Dim rep1 As IMSSymbolEntities.IJDRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation
  
    rep1.Name = "Physical"
    rep1.Description = "Physical Represntation of the Air Distrib Assembly"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim oRep1Outputs As IMSSymbolEntities.IJDOutputs
    Set oRep1Outputs = rep1

    For Index = 1 To 4
        oRep1Outputs.SetOutput O(Index)
        Set O(Index) = Nothing
    Next

    'Set the representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1
    
'===================================================================
' DEFINE MAINTENANCE ENVELOPE REPRESENTATION WITH DIFFERENT OUTPUTS
'===================================================================
    ' Redefine outputs
    oRep1Outputs.RemoveAllOutput
      
    ' define output
    Set O(1) = New IMSSymbolEntities.DOutput
    
    O(1).Name = "VVADAEnvelope"
    O(1).Description = "VVADA's Maintenance envelope"
    O(1).Properties = 0

    oRep1Outputs.SetOutput O(1)
    Set O(1) = Nothing

    ' Define representation
    rep1.Name = "MaintenanceEnvelope"
    rep1.Description = "Maintenance Envelope Representation of the VVADA"
    'Set the repID to Maintenance. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = Maintenance
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION

    ' Set representation
    RepsIf.SetRepresentation rep1
    
    Set rep1 = Nothing
    Set RepsIf = Nothing
    Set oRep1Outputs = Nothing

    'Define evaluation for Physical representation
    Dim PhysicalRepEval As IJDRepresentationEvaluation
    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "script for the Physical representation"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "SP3DAirDistribAssemblyAsm.VVADASym"

    'Define the evaluation associated to the Envelope representation
    Dim EnvelopeRepEval As DRepresentationEvaluation
    Set EnvelopeRepEval = New DRepresentationEvaluation
    EnvelopeRepEval.Name = "MaintenanceEnvelope"
    EnvelopeRepEval.Description = "Maintenance Envelope representation"
    EnvelopeRepEval.Properties = igREPRESENTATION_HIDDEN
    EnvelopeRepEval.Type = igREPRESENTATION_VBFUNCTION
    EnvelopeRepEval.ProgId = "SP3DAirDistribAssemblyAsm.VVADASym"

    'Set the evaluations for the Physical representation on the definition
    Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    RepEvalsIf.AddRepresentationEvaluation EnvelopeRepEval
    Set PhysicalRepEval = Nothing
    Set EnvelopeRepEval = Nothing
    Set RepEvalsIf = Nothing

'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================

    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

  Exit Sub

ErrorHandler:
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False

End Sub


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
    ElseIf StrComp(repName, "MaintenanceEnvelope") = 0 Then
        MaintenanceEnvelopeRep arrayOfInputs
    End If

End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF VVADA
'=========================================================================

Private Sub Physical(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    
    Dim ii As Integer
    
    Dim X1 As Double, Y1 As Double, Z1 As Double
    Dim X2 As Double, Y2 As Double, Z2 As Double
    Dim BasicAsmWidth As Double, BasicAsmHeight As Double
    Dim BasicAsmLength As Double, HeatingCoilLength As Double
    Dim MultiOutletLength As Double
    Dim InletDia As Double
    Dim MultiOutletDia As Double
'    Dim NozzleLength1 As Double, NozzleLength2 As Double, NozzleLength3 As Double
    Dim Coords(1 To 6) As Double
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)

    'assign to meaningful variables from the input array
    BasicAsmWidth = arrayOfInputs(2)
    BasicAsmHeight = arrayOfInputs(3)
    BasicAsmLength = arrayOfInputs(4)
    HeatingCoilLength = arrayOfInputs(5)
    MultiOutletLength = arrayOfInputs(6)
    InletDia = arrayOfInputs(7)
    MultiOutletDia = arrayOfInputs(8)
    
         
    
'====================
'BUILD BASIC ASSEMBLY
'====================

    Coords(1) = 0#
    Coords(2) = 0#
    Coords(3) = 0#
    Coords(4) = BasicAsmLength
    Coords(5) = BasicAsmWidth
    Coords(6) = BasicAsmHeight

    ' Get or create the definition
    Dim defColl As IJDDefinitionCollection
    Set defColl = oSymbolFactory.DefinitionCollection(m_outputColl.ResourceManager)
    Dim boxDef As IJDSymbolDefinition
    Dim definitionParams As Variant
    definitionParams = ""
    Set boxDef = defColl.GetDefinitionByProgId(True, "Box.BoxServices", vbNullString, definitionParams)
    
    Dim BasicAssembly As IMSSymbolEntities.DSymbol
    Set BasicAssembly = oSymbolFactory.PlaceSymbol(boxDef, m_outputColl.ResourceManager)

    Dim newEnumArg As IJDEnumArgument
    Dim IJEditJDArg As IJDEditJDArgument
    Set newEnumArg = New DEnumArgument
    Set IJEditJDArg = newEnumArg.IJDEditJDArgument

    Dim PC As IJDParameterContent
    Dim argument As IJDArgument
    For ii = 1 To 6
        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.UomValue = Coords(ii)
        PC.Type = igValue
        PC.UomType = 1
        ' Feed the Argument
        argument.Index = ii
        argument.Entity = PC
        ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
    Next
    
    BasicAssembly.IJDValuesArg.SetValues newEnumArg
    Dim IJDInputsArg As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg = BasicAssembly
    IJDInputsArg.Update
    Set IJDInputsArg = Nothing
    Set IJEditJDArg = Nothing
    Set newEnumArg = Nothing

    
    m_outputColl.AddOutput "BasicAssembly", BasicAssembly
    Set BasicAssembly = Nothing

'=====================================
' CONSTRUCTION OF HEATING COIL
'=====================================
     
    Coords(1) = BasicAsmLength
    Coords(2) = 0#
    Coords(3) = 0#
    Coords(4) = HeatingCoilLength + BasicAsmLength
    Coords(5) = BasicAsmWidth
    Coords(6) = BasicAsmHeight

    Dim HeatingCoil As IMSSymbolEntities.DSymbol
    Set HeatingCoil = oSymbolFactory.PlaceSymbol(boxDef, m_outputColl.ResourceManager)

    Set newEnumArg = New DEnumArgument
    Set IJEditJDArg = newEnumArg.IJDEditJDArgument

    For ii = 1 To 6
        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.UomValue = Coords(ii)
        PC.Type = igValue
        PC.UomType = 1
        ' Feed the Argument
        argument.Index = ii
        argument.Entity = PC
        ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
    Next

    HeatingCoil.IJDValuesArg.SetValues newEnumArg
    Set IJDInputsArg = HeatingCoil
    IJDInputsArg.Update
    Set IJDInputsArg = Nothing
    Set IJEditJDArg = Nothing
    Set newEnumArg = Nothing

    
    m_outputColl.AddOutput "HeatingCoil", HeatingCoil
    Set HeatingCoil = Nothing
       
'=====================================
' CONSTRUCTION OF MULTI-OUTLET
'=====================================
    Coords(1) = BasicAsmLength + HeatingCoilLength
    Coords(2) = 0#
    Coords(3) = 0#
    Coords(4) = BasicAsmLength + HeatingCoilLength + MultiOutletLength
    Coords(5) = BasicAsmWidth
    Coords(6) = BasicAsmHeight

    Dim MultiOutlet As IMSSymbolEntities.DSymbol
    Set MultiOutlet = oSymbolFactory.PlaceSymbol(boxDef, m_outputColl.ResourceManager)

    Set newEnumArg = New DEnumArgument
    Set IJEditJDArg = newEnumArg.IJDEditJDArgument

    For ii = 1 To 6
        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.UomValue = Coords(ii)
        PC.Type = igValue
        PC.UomType = 1
        ' Feed the Argument
        argument.Index = ii
        argument.Entity = PC
        ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
    Next
    
    MultiOutlet.IJDValuesArg.SetValues newEnumArg
    Set IJDInputsArg = MultiOutlet
    IJDInputsArg.Update
    Set IJDInputsArg = Nothing
    Set IJEditJDArg = Nothing
    Set newEnumArg = Nothing

    
    m_outputColl.AddOutput "MultiOutlet", MultiOutlet
    Set MultiOutlet = Nothing
        
    Set boxDef = Nothing
    Set defColl = Nothing

    Set oPart = Nothing
    Set oSymbolFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "VVADA definition"
        Debug.Assert False
End Sub

'=========================================================================
'CREATION OF MAINTENANCE ENVELOPE REPRESENTATION OF VVADA
'=========================================================================

Private Sub MaintenanceEnvelopeRep(ByRef arrayOfInputs())
    Dim X1 As Double, Y1 As Double, Z1 As Double
    Dim X2 As Double, Y2 As Double, Z2 As Double
    Dim BasicAsmWidth As Double, BasicAsmHeight As Double
    Dim BasicAsmLength As Double, HeatingCoilLength As Double
    Dim MultiOutletLength As Double
    Dim InletDia As Double
    Dim MultiOutletDia As Double
    Dim dNozLen1 As Double, dNozLen2 As Double, dNozLen3 As Double
    Dim Coords(1 To 6) As Double
    Dim ii As Integer
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
   
    On Error GoTo ErrorHandler
    
    'assign to meaningful variables from the input array
    BasicAsmWidth = arrayOfInputs(2)
    BasicAsmHeight = arrayOfInputs(3)
    BasicAsmLength = arrayOfInputs(4)
    HeatingCoilLength = arrayOfInputs(5)
    MultiOutletLength = arrayOfInputs(6)
    InletDia = arrayOfInputs(7)
    MultiOutletDia = arrayOfInputs(8)
    
'   Assumption: Nozzle length (Hvac and piping) is taken as EquipmentComponent length /1.7
    dNozLen1 = BasicAsmLength / 1.7
    dNozLen2 = BasicAsmLength / 1.7
    dNozLen3 = BasicAsmLength / 1.7
 
    ' Xmin, Ymin, Zmin and Xmax, Ymax, Zmax coordinates of the envelope box
    
    ' Box dimensions set to (0.1 * 2) times greater than the overall dimensions
    
    Dim totalLength As Double, totalWidth As Double, totalHeight As Double
    
    totalLength = BasicAsmLength + HeatingCoilLength + MultiOutletLength + _
                  dNozLen1 + dNozLen3
    totalWidth = BasicAsmWidth + dNozLen3 * 2#
    totalHeight = BasicAsmHeight
    
    Coords(1) = -(dNozLen1 + totalLength * 0.1)
    Coords(2) = -(dNozLen3 + totalWidth * 0.1)
    Coords(3) = -(totalHeight * 0.1)
    Coords(4) = totalLength - dNozLen1 + totalLength * 0.1
    Coords(5) = totalWidth - dNozLen3 + totalWidth * 0.1
    Coords(6) = totalHeight * 1.1

    ' Get or create the definition
    Dim defColl As IJDDefinitionCollection
    Set defColl = oSymbolFactory.DefinitionCollection(m_outputColl.ResourceManager)
    Dim boxDef As IJDSymbolDefinition
    Dim definitionParams As Variant
    definitionParams = ""
    Set boxDef = defColl.GetDefinitionByProgId(True, "Box.BoxServices", vbNullString, definitionParams)

    Dim env As IMSSymbolEntities.DSymbol
    Set env = oSymbolFactory.PlaceSymbol(boxDef, m_outputColl.ResourceManager)
    Set boxDef = Nothing
    Set defColl = Nothing

    Dim newEnumArg As IJDEnumArgument
    Dim IJEditJDArg As IJDEditJDArgument
    Set newEnumArg = New DEnumArgument
    Set IJEditJDArg = newEnumArg.IJDEditJDArgument

    Dim PC As IJDParameterContent
    Dim argument As IJDArgument
    For ii = 1 To 6
        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.UomValue = Coords(ii)
        PC.Type = igValue
        PC.UomType = 1
        ' Feed the Argument
        argument.Index = ii
        argument.Entity = PC
        ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
    Next

    env.IJDValuesArg.SetValues newEnumArg
    Dim IJDInputsArg As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg = env
    IJDInputsArg.Update
    Set IJDInputsArg = Nothing
    Set IJEditJDArg = Nothing
    Set newEnumArg = Nothing

    
    m_outputColl.AddOutput "VVADAEnvelope", env
    Set env = Nothing
    Set oSymbolFactory = Nothing

    Exit Sub
    
ErrorHandler:
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False

End Sub

