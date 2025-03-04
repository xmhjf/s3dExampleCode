VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMotorRoundDamper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'File
'    CMotorRoundDamper.cls
'
'Author
'    NN 17 March 2002
'
'Description
'    Definition of HVAC Motorized Round Damper.
'ProgID: SP3DMotorRoundDamper.CMotorRoundDamper
'Notes
'
'    <notes>
'
'History:
'
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   15.DEC.2004      MS  CR 67971 Corrected port indices and flow-directions in HVAC Symbols
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Const NEGLIGIBLE_THICKNESS = 0.0001

'Private Form As frmEditOccurrence

Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DMotorRoundDamper.CMotorRoundDamper"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DMotorRoundDamper.CMotorRoundDamper"
  oSymbolDefinition.CodeBase = CodeBase
  
  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId
  
  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  
  Set oSymbolDefinition = Nothing
  Set oSymbolFactory = Nothing
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

    ' Feed GSCADHRoundFlatFlange.HROUNDFF Definition
    ' Inputs:
    '          1. "Width"
    '
    '          2. "DamperLength"
    '          3. "Thickness"
    '          4. "FlangeThickness"
    
    '
    ' The representation "Physical" is set to the definition
    ' Physical :    (RepresentationId = 1)
  
    On Error GoTo ErrorHandler
  
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
   
    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition
        
    ' Create a new input by new operator
    Dim Inputs(1 To 4) As IMSSymbolEntities.DInput
    Dim index As Integer
  
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue

    Set Inputs(1) = New IMSSymbolEntities.DInput
    Inputs(1).Name = "Width"
    Inputs(1).Description = "Width of the Rect Duct"
    Inputs(1).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.1
    Inputs(1).DefaultParameterValue = PC
    
    
    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).Name = "DamperLength"
    Inputs(2).Description = "DamperLength between ducts"
    Inputs(2).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.1
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).Name = "Thickness"
    Inputs(3).Description = "Thickness of the Flat Flange of Damper"
    Inputs(3).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.01
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).Name = "FlangeWidth"
    Inputs(4).Description = "Width of the Flat Flange of Damper"
    Inputs(4).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.008
    Inputs(4).DefaultParameterValue = PC
       
    For index = 1 To 4
        InputsIf.SetInput Inputs(index), index + 1
        Set Inputs(index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 4) As IMSSymbolEntities.DOutput
    
    For index = 1 To 4
        Set O(index) = New IMSSymbolEntities.DOutput
        O(index).properties = 0
    Next

    O(1).Name = "Shaft"
    O(1).Description = "Shaft of Damper"
    
    O(2).Name = "Motor"
    O(2).Description = "Motor of Damper"

    O(3).Name = "AirIN"
    O(3).Description = "AirIN Hvac Port1 of Damper"
    
    O(4).Name = "AirOUT"
    O(4).Description = "AirOUT Hvac Port2 of Damper"
    
    'Define the representation "Physical"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation
  
    rep1.Name = "Physical"
    rep1.Description = "Physical Represntation of the Air Distrib Assembly"
    rep1.properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim oRepPhysicalOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepPhysicalOutputs = rep1
        
    For index = 1 To 4
        oRepPhysicalOutputs.SetOutput O(index)
        Set O(index) = Nothing
    Next

    'Set the representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1

   
    'Define evaluation for Physical representation
    Dim PhysicalRepEval As DRepresentationEvaluation
    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "script for the Physical representation"
    PhysicalRepEval.properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "SP3DMotorRoundDamper.CMotorRoundDamper"

    
    'Set the evaluations for the Physical representation on the definition
    Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    
    Set PhysicalRepEval = Nothing
    
    Set RepEvalsIf = Nothing

'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================

    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
    Set oRepPhysicalOutputs = Nothing
  Exit Sub

ErrorHandler:
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False

End Sub


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    'If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
        'End If

End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF HRECTFF
'=========================================================================


Private Sub Physical(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    
    
    Dim Width As Double, Depth As Double
    Dim DamperLength As Double, Thickness As Double
    Dim FlangeWidth As Double
    Dim HDamperLength As Double
    Dim NozzleLength As Double
    'Dim NozzleFlangeThickness As Double
    Dim CornerRadius As Double
        
    'assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    DamperLength = arrayOfInputs(3)
    Thickness = arrayOfInputs(4)
    FlangeWidth = arrayOfInputs(5)
    Depth = Width
    
'=====================================
'BUILD HVAC NOZZLE ON BASIC ASSEMBLY
'=====================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim nozzleIndex As Integer
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iLogicalDistPort As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim CptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim ObjShaft As Object
    Dim ObjMotor As Object
    Dim StPoint As IJDPosition
    Dim enPoint As IJDPosition
    Set StPoint = New DPosition
    Set enPoint = New DPosition
    Dim ShaftDia As Double
    Dim MotorDia As Double
    Dim iPortIndex As Integer

'Construct Shaft Output 1
    MotorDia = 0.085
    ShaftDia = 0.015
    StPoint.Set 0, Width / 3, 0
    enPoint.Set 0, Width / 2 + MotorDia * 0.6, 0
    Set ObjShaft = PlaceCylinder(m_outputColl, StPoint, enPoint, ShaftDia, False)
    m_outputColl.AddOutput "Shaft", ObjShaft
    Set ObjShaft = Nothing

'Construct Motor Output 2
    StPoint.Set 0, Width / 2 + MotorDia, DamperLength / 4
    enPoint.Set 0, Width / 2 + MotorDia, -DamperLength * 5 / 8
    Set ObjMotor = PlaceCylinder(m_outputColl, StPoint, enPoint, MotorDia, True)
    m_outputColl.AddOutput "Motor", ObjMotor
    Set ObjMotor = Nothing
    
'=====================================
'BUILD AirIN HVAC port 1 Of Damper
'=====================================
   'Set HVAC nozzle parameters
    iPortIndex = 1
    PortDepth = 0#
    CptOffset = 0#
    
    EndPrep = 11
    CornerRadius = 0#
    HDamperLength = DamperLength / 2
    'NozzleFlangeThickness = Thickness + HDamperLength
    NozzleLength = DamperLength * 1.05
    'NozzleLength Has to be GREATER than NozzleFlangeThickness
    

    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    
    
    FlowDir = DistribFlow_BOTH

    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "AirIN", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = -HDamperLength
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    'dir.Set 0, 1, 0
    'iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "AirIN", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
'=====================================
'BUILD AirOUT HVAC port 2 Of Damper
'=====================================
    iPortIndex = 2
    FlowDir = DistribFlow_BOTH
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "AirOUT", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    
    
    X = HDamperLength
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    'dir.Set 0, 1, 0
    'iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "AirOUT", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
    
    
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "CMotorRoundDamper definition"
        Debug.Assert False
End Sub

