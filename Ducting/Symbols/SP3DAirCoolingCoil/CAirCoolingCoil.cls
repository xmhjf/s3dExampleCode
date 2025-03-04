VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAirCoolingCoil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'File
'    CAirCoolingCoil.cls
'
'Author
'    NN 10 Mars 2002
'
'Description
'    Definition of HVAC AirCoolingCoil.
'
'Notes
'
'    <notes>
'
'History:
'   14 Sep, 1999 : PR
'       Fixed TR#7650.Removed the Usage of m_SymbolDefinition.
'   27th Sep, 1999: APS [APS]
'     Took care of P2R2 symbol impact.
'     For setting outputs on a rep, one
'     needs to query for outputs from rep and set them.
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
  IJDUserSymbolServices_GetDefinitionName = "SP3DAirCoolingCoil.CAirCoolingCoil"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DAirCoolingCoil.CAirCoolingCoil"
  oSymbolDefinition.CodeBase = CodeBase
  
  ' Give a unique name to the symbol definition
  oSymbolDefinition.name = oSymbolDefinition.ProgId
  
  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  
  Set oSymbolDefinition = Nothing
  Set oSymbolFactory = Nothing
  Exit Function

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.description
    Debug.Assert False
End Function


Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    ' Feed SP3DHFlatFlange.HFF Definition
    ' Inputs:
    '          1. "Width"
    '          2. "Depth"
    '          3. "CoilLength"
    '          4. "Thickness"
    '          5. "FlangeThickness"
    
    '
    ' The representation "Physical" is set to the definition
    ' Physical :    (RepresentationId = 1)
  
    On Error GoTo ErrorHandler
  
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
    
    ' Create a new input by new operator
    Dim Inputs(1 To 6) As IMSSymbolEntities.DInput
    Dim index As Integer
  
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition
    
    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).name = "Width"
    Inputs(2).description = "Width of the Rect Duct"
    Inputs(2).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.4
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).name = "Depth"
    Inputs(3).description = "Depth of the Rect Duct"
    Inputs(3).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.2
    Inputs(3).DefaultParameterValue = PC
    
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).name = "Length"
    Inputs(4).description = "CoilLength between ducts"
    Inputs(4).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.35
    Inputs(4).DefaultParameterValue = PC
    
    Set Inputs(5) = New IMSSymbolEntities.DInput
    Inputs(5).name = "Thickness"
    Inputs(5).description = "Thickness of the Flat Flange of AirCoolingCoil"
    Inputs(5).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.01
    Inputs(5).DefaultParameterValue = PC
    
    Set Inputs(6) = New IMSSymbolEntities.DInput
    Inputs(6).name = "FlangeWidth"
    Inputs(6).description = "Width of the Flat Flange of AirCoolingCoil"
    Inputs(6).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.02
    Inputs(6).DefaultParameterValue = PC
       
    For index = 2 To 6
        InputsIf.SetInput Inputs(index), index
        Set Inputs(index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 4) As IMSSymbolEntities.DOutput
    
    For index = 1 To 4
        Set O(index) = New IMSSymbolEntities.DOutput
        O(index).properties = 0
    Next
    O(1).name = "HvacNozzle1"
    O(1).description = "Hvac Port1 of AirCoolingCoil"
    
    O(2).name = "HvacNozzle2"
    O(2).description = "Hvac Port2 of AirCoolingCoil"

   
    O(3).name = "PipingNozzle3"
    O(3).description = "Piping Port3 of AirCoolingCoil"
    
    O(4).name = "PipingNozzle4"
    O(4).description = "Piping Port4 of AirCoolingCoil"
    
    'Define the representation "Physical"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation
  
    rep1.name = "Physical"
    rep1.description = "Physical Represntation of the Air Distrib Assembly"
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
    PhysicalRepEval.name = "Physical"
    PhysicalRepEval.description = "script for the Physical representation"
    PhysicalRepEval.properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "SP3DAirCoolingCoil.CAirCoolingCoil"

    
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
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.description
  Debug.Assert False

End Sub


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    'If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
        'End If

End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF HFF
'=========================================================================


Private Sub Physical(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    
    
    Dim Width As Double, Depth As Double
    Dim CoilLength As Double, Thickness As Double
    Dim FlangeWidth As Double
    Dim HCoilLength As Double
    Dim NozzleLength As Double
    Dim NozzleFlangeThickness As Double
    Dim CornerRadius As Double
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    'assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    CoilLength = arrayOfInputs(4)
    Thickness = arrayOfInputs(5)
    FlangeWidth = arrayOfInputs(6)
  
    
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
    Dim iPortIndex As Integer
    
    'Set HVAC nozzle parameters
    iPortIndex = 1
    PortDepth = 0#
    CptOffset = 0
    
    EndPrep = 11
    CornerRadius = 0#
    HCoilLength = CoilLength / 2
    'NozzleFlangeThickness = Thickness + HCoilLength
    NozzleLength = CoilLength - Thickness
    'NozzleLength Has to be GREATER than NozzleFlangeThickness
    

    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    
    
    FlowDir = DistribFlow_BOTH

    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "AirIN", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = -HCoilLength - CptOffset
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "HvacNozzle1", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
'=====================================
'BUILD HVACNOZZLE2 Of AirCoolingCoil
'=====================================
    iPortIndex = 2
    
    FlowDir = DistribFlow_BOTH
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "AirOUT", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    
    
    X = HCoilLength + CptOffset
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir
'    NozzleLength = Thickness * 1.001
    

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
     
    
    
    
    
    
    
    
 '=====================================
'BUILD Piping NOZZLE3 Of AirCoolingCoil
'=====================================


    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -CoilLength / 4, 0, Width / 2
    oDir.Set 0, 0, 1
    Set objNozzle = CreateNozzle(3, oPart, m_outputColl, oDir, oPlacePoint)
' Set the output
''''    iOutput = iOutput + 1
''''    m_outputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
''''    Set objNozzle = Nothing

    m_outputColl.AddOutput "PipingNozzle3", objNozzle
    Set objNozzle = Nothing





    '=====================================
'BUILD Piping NOZZLE4 Of AirCoolingCoil
'=====================================


'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set CoilLength / 4, 0, Width / 2
    oDir.Set 0, 0, 1
    Set objNozzle = CreateNozzle(4, oPart, m_outputColl, oDir, oPlacePoint)
' Set the output
''''    iOutput = iOutput + 1
''''    m_outputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
''''    Set objNozzle = Nothing

    m_outputColl.AddOutput "PipingNozzle4", objNozzle
    Set objNozzle = Nothing


    
    
    
    
    
    
    
    
    
    
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "HCoilHumidifier definition"
        Debug.Assert False
End Sub

