VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDivisionW2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'File
'    CDivisionW2.cls
'
'Author
'    NN 20011219
'
'Description
'    Definition of HVAC Division Width divided into two cells.
'
'Notes
'
'    <notes>
'
'History:
'   19th Dec 2001: NN Simple nozzle True is:
'   a solid plane without any flange thickness but not just a circle
'   Constructed a division: Width divided into two cells, cell 1 is closer
'   to the division reference (cylinder)
'   2nd June 2002: NN Raise the error: Err.Raise Err.Number
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
  IJDUserSymbolServices_GetDefinitionName = "SP3DDivisionW2.CDivisionW2"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DDivisionW2.CDivisionW2"
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

    ' Feed GSCADHRoundFlatFlange.HROUNDFF Definition
    ' Inputs:
    '          1. "Width"
    '          2. "Depth"
    '          3. "Cell"
    '          4. "CellDimension"
    
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
    Inputs(1).name = "Width"
    Inputs(1).description = "Width of the Rect Duct"
    Inputs(1).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.4
    Inputs(1).DefaultParameterValue = PC
    
    
    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).name = "Depth"
    Inputs(2).description = "Depth of the Rect Duct"
    Inputs(2).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.2
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).name = "Cell"
    Inputs(3).description = "Cell number"
    Inputs(3).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 1
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).name = "CellDimension"
    Inputs(4).description = "Width of the Cell"
    Inputs(4).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.1
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

   
    O(1).name = "GrandPort1"
    O(1).description = "Grand Port1 HvacPort 1 of Division"
    
    O(2).name = "Cell1ClosetoRef"
    O(2).description = "Close to Reference, at Cell = 1, HvacPort 2 of Division"
    
    O(3).name = "Cell2FarfromRef"
    O(3).description = "Far from Reference, at Cell = 2, HvacPort 3 of Division"
    
    O(4).name = "DivisionReference"
    O(4).description = "Reference of Division"

    
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
    PhysicalRepEval.ProgId = "SP3DDivisionW2.CDivisionW2"

    
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
'CREATION OF PHYSICAL REPRESENTATION OF HVAC Division 2
'=========================================================================


Private Sub Physical(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    
    Dim Width As Double
    Dim Depth As Double, Cell As Long
    Dim CellDimension As Double
    Dim NozzleLength As Double
    Dim CornerRadius As Double
    Dim Width2 As Double
    Dim Width3 As Double
    'assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    Cell = arrayOfInputs(4)
    CellDimension = arrayOfInputs(5)


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
    CptOffset = 0#
    
    EndPrep = 301
    CornerRadius = 0#

    NozzleLength = NEGLIGIBLE_THICKNESS * 2

    'NozzleLength Has to be GREATER than NozzleFlangeThickness

    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    
    
    FlowDir = DistribFlow_IN

    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    
    Dim NThickness As Double
    Dim NFlangeWidth As Double
    NThickness = NEGLIGIBLE_THICKNESS
    NFlangeWidth = NEGLIGIBLE_THICKNESS
        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            NThickness, NFlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", PortDepth, CptOffset, False, m_outputColl.ResourceManager)

   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = 0
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

    m_outputColl.AddOutput "GrandPort1", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
'=====================================
'BUILD HVACNOZZLE2 Of Division
'=====================================
    iPortIndex = 2

If Cell < 2 Then
        Width2 = CellDimension
        Width3 = Width - CellDimension
    Else
        Width3 = CellDimension
        Width2 = Width - CellDimension
End If


If Width2 >= Depth Then
FlowDir = DistribFlow_OUT
Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            NThickness, NFlangeWidth, FlowDir, Width2, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", PortDepth, CptOffset, False, m_outputColl.ResourceManager)

   
    'Position of the nozzle should be the conenct point of the nozzle
    
    
    X = NEGLIGIBLE_THICKNESS
    Y = 0
    Z = -(Width / 2 - Width2 / 2)

    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

 Else
    Dim NWidth2 As Double
    Dim NDepth2 As Double
    NWidth2 = Depth
    NDepth2 = Width2
    FlowDir = DistribFlow_OUT
        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            NThickness, NFlangeWidth, FlowDir, NWidth2, _
                                            NDepth2, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", PortDepth, CptOffset, False, m_outputColl.ResourceManager)


    'Position of the nozzle should be the conenct point of the nozzle
    
    
    X = NEGLIGIBLE_THICKNESS
    Y = 0
    Z = -(Width / 2 - Width2 / 2)

    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, 0, 1
    iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength
End If

    m_outputColl.AddOutput "Cell1ClosetoRef", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing

'=====================================
'BUILD HVACNOZZLE3 Of Division
'=====================================
iPortIndex = 3

If Width3 >= Depth Then
    FlowDir = DistribFlow_OUT
        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            NThickness, NFlangeWidth, FlowDir, Width3, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", PortDepth, CptOffset, False, m_outputColl.ResourceManager)

    'Position of the nozzle should be the conenct point of the nozzle
    
    
    X = NEGLIGIBLE_THICKNESS
    Y = 0
    Z = (Width / 2 - Width3 / 2)
    
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength
 Else
    Dim NWidth3 As Double
    Dim NDepth3 As Double
    NWidth3 = Depth
    NDepth3 = Width3
    FlowDir = DistribFlow_OUT
        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            NThickness, NFlangeWidth, FlowDir, NWidth3, _
                                            NDepth3, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", PortDepth, CptOffset, False, m_outputColl.ResourceManager)

   
    'Position of the nozzle should be the conenct point of the nozzle
    
    
    X = NEGLIGIBLE_THICKNESS
    Y = 0
    Z = (Width / 2 - Width3 / 2)

    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, 0, 1
    iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength
End If
    m_outputColl.AddOutput "Cell2FarfromRef", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing

Dim ObjDivisionRef As Object
 Dim stPoint   As New AutoMath.DPosition
 Dim enPoint   As New AutoMath.DPosition
 Dim ldiam     As Double
 
 stPoint.Set 0, -Depth / 2, -Width / 2
 enPoint.Set 0.001, -Depth / 2, -Width / 2
 ldiam = 0.02

Set ObjDivisionRef = PlaceCylinder(m_outputColl, stPoint, enPoint, ldiam, True)
    m_outputColl.AddOutput "DivisionReference", ObjDivisionRef
    Set ObjDivisionRef = Nothing
    Exit Sub
    
ErrorLabel:
Err.Raise Err.Number
End Sub

