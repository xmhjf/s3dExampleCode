VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HRndReducingHdTee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2005-07, Intergraph Corporation. All rights reserved.
'
'File
'    HRndReducingHdTee.cls
'    ProgID: SP3DHRndRedHeaderTee.HRndReducingHdTee
'
'Author
'    svsmylav
'    Creation Date:  Friday, Jun 3 2005
'
'Description
'    Definition of HVAC Fincantieri Tee with Port2 of reducing size (Reference: Fincantieri drawing).
'
'Notes
'
'    <notes>
'
'History:
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  20.Nov.2007     RRK  CR-123952 Updated the symbol to use new private RetrievePipePortParameters function to treat face projection correctly
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Const NEGLIGIBLE_THICKNESS = 0.0001

Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHRndRedHeaderTee.HRndReducingHdTee"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
  On Error GoTo ErrorHandler
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHRndRedHeaderTee.HRndReducingHdTee"
  oSymbolDefinition.CodeBase = CodeBase
  
  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId
  
  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  Set oSymbolDefinition = Nothing
  Set oSymbolFactory = Nothing
  
  Exit Function

ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Function


Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    ' Feed GSCADHRoundFlatFlange.HROUNDFF Definition
    ' Inputs:
    '          1. "HWidth"
    '          2. "Width2"
    '          3. "HLength"
    '          4. "BLength"
    '          5. "NozzleOffset"
    '
    ' The representation "Physical" is set to the definition
    ' Physical :    (RepresentationId = 1)
  
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
  
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    ' Create a new input by new operator
    Dim Inputs(1 To 8) As IMSSymbolEntities.DInput
    Dim index As Integer
  
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue

    Set Inputs(1) = New IMSSymbolEntities.DInput
    Inputs(1).Name = "Width"
    Inputs(1).Description = "Width of the Header"
    Inputs(1).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.35
    Inputs(1).DefaultParameterValue = PC

    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).Name = "Width2"
    Inputs(2).Description = "Branch Width"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.2
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).Name = "HLength"
    Inputs(3).Description = "Length of the Tee along the Header"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 1.4
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).Name = "BLength"
    Inputs(4).Description = "Length of the Tee along the Branch"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.5
    Inputs(4).DefaultParameterValue = PC
    
    Set Inputs(5) = New IMSSymbolEntities.DInput
    Inputs(5).Name = "NozzleOffset"
    Inputs(5).Description = "Nozzle Offset"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.068
    Inputs(5).DefaultParameterValue = PC
    
    Set Inputs(6) = New IMSSymbolEntities.DInput
    Inputs(6).Name = "FacetoCenter"
    Inputs(6).Description = "Faceto Center"
    Inputs(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.25
    Inputs(6).DefaultParameterValue = PC
    
    Set Inputs(7) = New IMSSymbolEntities.DInput
    Inputs(7).Name = "NozzleSpacing"
    Inputs(7).Description = "Nozzle Spacing"
    Inputs(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.2
    Inputs(7).DefaultParameterValue = PC
    
    Set Inputs(8) = New IMSSymbolEntities.DInput
    Inputs(8).Name = "InsulationThickness"
    Inputs(8).Description = "Insulation Thickness"
    Inputs(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(8).DefaultParameterValue = PC
       
    For index = 1 To 8
        InputsIf.SetInput Inputs(index), index + 1
        Set Inputs(index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 19) As IMSSymbolEntities.DOutput
    
    For index = 1 To 19
        Set O(index) = New IMSSymbolEntities.DOutput
        O(index).Properties = 0
    Next
   
    O(1).Name = "HvacNozzle1"
    O(1).Description = "Hvac Port1 of Tee"
    
    O(2).Name = "HvacNozzle2"
    O(2).Description = "HvacPort2 of Tee"
    
    O(3).Name = "HvacNozzle3"
    O(3).Description = "HvacPort3 of Tee"
    
    O(4).Name = "TangentPort1Side"
    O(4).Description = "Tangent Port 1 Side"
    
    O(5).Name = "BendPort1Side"
    O(5).Description = "Bend Port 1 Side"
    
    O(6).Name = "BendPort2Side"
    O(6).Description = "Bend Port 2 Side"
    
    O(7).Name = "TangentPort2Side"
    O(7).Description = "Tangent Port 2 Side"
    
    O(8).Name = "ReducerPort2Side"
    O(8).Description = "Reducer Port 2 Side"
    
    O(9).Name = "PNoz1"
    O(9).Description = "Piping Nozzle 1"
    
    O(10).Name = "PNoz2"
    O(10).Description = "Piping Nozzle 2"

    O(11).Name = "InsHvacPort1"
    O(11).Description = "Insulation of HvacPort1"

    O(12).Name = "InsBendPort1Side"
    O(12).Description = "Insulation of Bend Port 1 Side"

    O(13).Name = "InsHvacPort3"
    O(13).Description = "Insulation of HvacPort3"

    O(14).Name = "InsBendPort2Side"
    O(14).Description = "Insulation of Bend Port 1 Side"

    O(15).Name = "InsTangentPort2Side"
    O(15).Description = "Insulation of Tangent Port 2 Side"

    O(16).Name = "InsReducerPort2Side"
    O(16).Description = "Insulation of Reducer Port 2 Side"

    O(17).Name = "InsHvacPort2"
    O(17).Description = "Insulation of HvacPort1"

    O(18).Name = "InsPNoz1"
    O(18).Description = "Insulation of Piping Nozzle 1"

    O(19).Name = "InsPNoz2"
    O(19).Description = "Insulation of Piping Nozzle 2"

    'Define the representation "Physical"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation
  
    rep1.Name = "Physical"
    rep1.Description = "Physical Represntation of the Air Distrib Assembly"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim oRepPhysicalOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepPhysicalOutputs = rep1
        
    For index = 1 To 10
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
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "SP3DHRndRedHeaderTee.HRndReducingHdTee"

    'Set the evaluations for the Physical representation on the definition
    Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval

' Redefine outputs
   oRepPhysicalOutputs.RemoveAllOutput
   Set oRepPhysicalOutputs = Nothing
   
'  Insulation representation
   rep1.Name = "Insulation"
   rep1.Description = "Insulation of the Elbow"
   'Set the repID to Insulation.
   rep1.RepresentationId = Insulation
   rep1.Properties = igREPRESENTATION_ISVBFUNCTION
   
   Dim oRepInsulationOutputs As IMSSymbolEntities.IJDOutputs
   Set oRepInsulationOutputs = rep1
   
   'Set the output
    For index = 11 To 19
        oRepInsulationOutputs.SetOutput O(index)
        Set O(index) = Nothing
    Next
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1
    
' Define the evaluation associated to the Insulation representation
    Dim InsulationRepEval As DRepresentationEvaluation
    Set InsulationRepEval = New DRepresentationEvaluation
    InsulationRepEval.Name = "Insulation"
    InsulationRepEval.Description = "Insulation representation"
    InsulationRepEval.Properties = igREPRESENTATION_HIDDEN
    InsulationRepEval.Type = igREPRESENTATION_VBFUNCTION
    InsulationRepEval.ProgId = "SP3DHRndRedHeaderTee.HRndReducingHdTee"
    
    'Set the evaluations for the insulation representation on the definition
    RepEvalsIf.AddRepresentationEvaluation InsulationRepEval
    
    Set PhysicalRepEval = Nothing
    Set InsulationRepEval = Nothing
    Set RepEvalsIf = Nothing

'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================

    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
    Set oRepPhysicalOutputs = Nothing
  Exit Sub

ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
    ElseIf StrComp(repName, "Insulation") = 0 Then
        InsulationRep arrayOfInputs
    End If

End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF HRECTFF
'=========================================================================

Private Sub Physical(ByRef arrayOfInputs())
    Const METHOD = "Physical"
    On Error GoTo ErrorLabel
        
    Dim oPartFclt      As PartFacelets.IJDPart
    Dim Width As Double, depth As Double
    Dim BWidth As Double, BDepth As Double
    Dim Width2 As Double
    Dim HLength As Double
    Dim BLength As Double, FlangeWidth As Double
    Dim NozzleLength As Double, Thickness As Double
    Dim Angle As Double
    Dim NozzleOffset As Double
    Dim FacetoCenter As Double
    Dim NozzleSpacing As Double
    'Dim NozzleFlangeThickness As Double
    Dim CornerRadius As Double
        
    'assign to meaningful variables from the input array
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    BWidth = Width
    Width2 = arrayOfInputs(3)
    HLength = arrayOfInputs(4)
    BLength = arrayOfInputs(5)
    NozzleOffset = arrayOfInputs(6)
    FacetoCenter = arrayOfInputs(7)
    NozzleSpacing = arrayOfInputs(8)

    depth = Width
    BDepth = BWidth

    Dim PI       As Double
    PI = 4 * Atn(1)

'   Origin taken at the intersection of header and
'   Fincantieri part drawing shows an angle of 55 degrees measured from Z axis. Instead 35 degree
'   taken from positive Y-axis.
    Angle = 35 * PI / 180
    
'============================================================
'BUILD HVAC NOZZLE ON BASIC ASSEMBLY
'============================================================
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
    Dim minHL As Double
    Dim minBL As Double
    
    ' Define default minimum Tee or Reducing Tee header and branch lengths
    minHL = Width + 0.05
    minBL = minHL / 2
    If CmpDblLessThan(HLength, minHL) Then
        HLength = minHL
    End If
    
    If CmpDblLessThan(BLength, minBL) Then
        BLength = minBL
    End If
    
    'Set HVAC nozzle parameters
    iPortIndex = 1
    
    EndPrep = 301
    CornerRadius = 0#

'   Nozzle Length is 90 mm as per Fincantieri part drawing
    NozzleLength = 0.09
    'NozzleLength Has to be GREATER than NozzleFlangeThickness

    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    FlangeWidth = 0#
    Thickness = 0.0001
    PortDepth = 0#
    CptOffset = 0#
    FlowDir = DistribFlow_BOTH

    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    Dim dPort1toCenter As Double
    'dPort1toCenter is 602.5 mm as per Fincantieri part drawing
    dPort1toCenter = 0.6025
    X = -dPort1toCenter
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

    m_outputColl.AddOutput "HvacNozzle1", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
'==================================================
'BUILD HVACNOZZLE2 Of Tee
'==================================================
    iPortIndex = 2
    EndPrep = 11
    PortDepth = 0#
    CptOffset = 0#
    'NozzleLength Has to be GREATER than NozzleFlangeThickness

    ' Assumption: Flange diameter is taken as 0.3429m (Corresponding to 200 mm pipe size) and Flange thickness is assumed
    ' to be 0.022m. So Nozzle Length is taken as 0.025 m which is greater than the Flange thickness.

    FlangeWidth = 0.07145 '(0.3429 - 0.2)/2
    Thickness = 0.022
    
    FlangeWidth = 0#
    Thickness = 0.0001
    NozzleLength = 0.025
    FlowDir = DistribFlow_OUT
    PortStatus = DistribPortStatus_BASE

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width2, _
                                            depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    
    'Y coordinate of port 2 is 68 mm below the origin as per Fincantieri part drawing
    X = HLength - dPort1toCenter
    Y = -NozzleOffset
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

    m_outputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
'============================================
'BUILD HVACNOZZLE3
'============================================
    iPortIndex = 3
    EndPrep = 301
    CornerRadius = 0#

    'NozzleLength Has to be GREATER than NozzleFlangeThickness
    

    ' Assumption: Flange diameter is taken as 0.3429m (Corresponding to 200 mm pipe size) and
    ' Flange thickness assumed to be 0.022m.

    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    FlangeWidth = 0#
    Thickness = 0.0001
    PortDepth = 0#
    CptOffset = 0#
    FlowDir = DistribFlow_BOTH

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, BWidth, _
                                            BDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    'Port 3 is at 35 degrees angle from the Y axis as per Fincantieri part drawing
    X = 0#
    Y = BLength * Cos(Angle)
    Z = BLength * Sin(Angle)
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.

    dir.Set 0, Cos(Angle), Sin(Angle)
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, -Sin(Angle), Cos(Angle)
    'iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = BLength / Sqr(2) 'Sqr(2): Sin(45deg) is component of branch length upto elbow top

    m_outputColl.AddOutput "HvacNozzle3", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
'============================================
'BUILD Piping Nozzle 1
'============================================
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth_port      As Double

'   Note: Existing RetrieveParameters function in basGeom3D fails if the first port is a HVAC port. So we
'   made a local function RetrievePipePortParameters with a modified logic.
    RetrievePipePortParameters 4, oPartFclt, m_outputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth_port
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    NozzleLength = FacetoCenter
    If CmpDblLessThan(NozzleLength, flangeThick) Then NozzleLength = flangeThick
    oPlacePoint.Set NozzleSpacing / 2, 0, -FacetoCenter - sptOffset + depth_port
    oDir.Set 0, 0, -1
    Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_outputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    m_outputColl.AddOutput "PNoz1", objNozzle
    Set objNozzle = Nothing

'============================================
'BUILD Piping Nozzle 2
'============================================
    RetrievePipePortParameters 5, oPartFclt, m_outputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth_port

    If CmpDblLessThan(NozzleLength, flangeThick) Then NozzleLength = flangeThick
    
    oPlacePoint.Set -NozzleSpacing / 2, 0, -FacetoCenter - sptOffset + depth_port
    oDir.Set 0, 0, -1
    Set objNozzle = CreateNozzleWithLength(5, oPartFclt, m_outputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    m_outputColl.AddOutput "PNoz2", objNozzle
    Set objNozzle = Nothing

'   Prepare Tee body oututs
' Insert your code for output TangentPort1Side
    Dim ObjTangentPort1Side As Object


    Dim dTangentPort1SideLength As Double
    Dim dTotalLengthOfTwoBends As Double
    Dim dTangentPort2SideLength As Double
    Dim dReducerPort2SideLength As Double
    
    'Following dimensions are taken from Fincantieri part drawing
    dTangentPort1SideLength = 0.155
    dTotalLengthOfTwoBends = 0.715
    dTangentPort2SideLength = 0.1
    dReducerPort2SideLength = 0.34
    
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set -dTotalLengthOfTwoBends / 2 - dTangentPort1SideLength, 0, 0
    enPoint.Set stPoint.X + dTangentPort1SideLength, stPoint.Y, stPoint.Z

    Set ObjTangentPort1Side = PlaceCylinder(m_outputColl, stPoint, enPoint, Width, False)
    
'   Set the output
    m_outputColl.AddOutput "TangentPort1Side", ObjTangentPort1Side
    Set ObjTangentPort1Side = Nothing
    
' Insert your code for output BendPort1Side
    Dim ObjBendPort1Side As Object
    Dim circleCenter    As AutoMath.DPosition
    Dim circleNormal    As AutoMath.DVector
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set circleCenter = New AutoMath.DPosition
    circleCenter.Set -dTotalLengthOfTwoBends / 2, 0, 0
    Set circleNormal = New AutoMath.DVector
    circleNormal.Set 1, 0, 0
    
' Construct a circle that will be used to project the disc
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.X, circleCenter.Y, circleCenter.Z, _
                        circleNormal.X, circleNormal.Y, circleNormal.Z, _
                        Width / 2)
    
    Dim axisVector As AutoMath.DVector
    Dim centerPoint As AutoMath.DPosition
    Dim revAngle As Double
    
    Set axisVector = New AutoMath.DVector
    Set centerPoint = New AutoMath.DPosition
    
    axisVector.Set 0, -Sin(Angle), Cos(Angle)
'   Assumption: revAngle is taken as 45 deg and BLength
    revAngle = PI / 4
    centerPoint.Set -dTotalLengthOfTwoBends / 2, BLength * Cos(Angle), BLength * Sin(Angle)

    Set ObjBendPort1Side = geomFactory.Revolutions3d.CreateByCurve( _
                                                    m_outputColl.ResourceManager, _
                                                    objCircle, _
                                                    axisVector.X, axisVector.Y, axisVector.Z, _
                                                    centerPoint.X, centerPoint.Y, centerPoint.Z, _
                                                    revAngle, False)
                                                    
'   Set the output
    m_outputColl.AddOutput "BendPort1Side", ObjBendPort1Side
    Set ObjBendPort1Side = Nothing

' Insert your code for output BendPort2Side
    Dim ObjBendPort2Side As Object
    circleCenter.Set dTotalLengthOfTwoBends / 2, 0, 0
    circleNormal.Set 1, 0, 0
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.X, circleCenter.Y, circleCenter.Z, _
                        circleNormal.X, circleNormal.Y, circleNormal.Z, _
                        Width / 2)
    axisVector.Set 0, Sin(Angle), -Cos(Angle)
    centerPoint.Set dTotalLengthOfTwoBends / 2, BLength * Cos(Angle), BLength * Sin(Angle)
    Set ObjBendPort2Side = geomFactory.Revolutions3d.CreateByCurve( _
                                                    m_outputColl.ResourceManager, _
                                                    objCircle, _
                                                    axisVector.X, axisVector.Y, axisVector.Z, _
                                                    centerPoint.X, centerPoint.Y, centerPoint.Z, _
                                                    revAngle, False)
                                                    
'   Set the output
    m_outputColl.AddOutput "BendPort2Side", ObjBendPort2Side
    Set ObjBendPort2Side = Nothing

' Insert your code for output TangentPort2Side
    Dim ObjTangentPort2Side As Object

    stPoint.Set dTotalLengthOfTwoBends / 2, 0, 0
    enPoint.Set stPoint.X + dTangentPort2SideLength, stPoint.Y, stPoint.Z

    Set ObjTangentPort2Side = PlaceCylinder(m_outputColl, stPoint, enPoint, Width, False)
    
'   Set the output
    m_outputColl.AddOutput "TangentPort2Side", ObjTangentPort2Side
    Set ObjTangentPort2Side = Nothing
    
' Insert your code for output ReducerPort2Side
    Dim ObjReducerPort2Side As Object
    circleCenter.Set dTotalLengthOfTwoBends / 2 + dTangentPort2SideLength, 0, 0
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.X, circleCenter.Y, circleCenter.Z, _
                        circleNormal.X, circleNormal.Y, circleNormal.Z, _
                        Width / 2)
                        
    Dim objCircle2 As IngrGeom3D.Circle3d
    circleCenter.Set dTotalLengthOfTwoBends / 2 + dTangentPort2SideLength + dReducerPort2SideLength, -NozzleOffset, 0
    Set objCircle2 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.X, circleCenter.Y, circleCenter.Z, _
                        circleNormal.X, circleNormal.Y, circleNormal.Z, _
                        Width2 / 2)

    Set ObjReducerPort2Side = geomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                                                            objCircle, objCircle2, True)
'   Set the output
    m_outputColl.AddOutput "ReducerPort2Side", ObjReducerPort2Side
    Set ObjReducerPort2Side = Nothing

    Exit Sub
    
ErrorLabel:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

'=========================================================================
'CREATION OF INSULATION ASPECT
'=========================================================================

Private Sub InsulationRep(ByRef arrayOfInputs())
    Const METHOD = "InsulationRep"
    On Error GoTo ErrorLabel

    Dim ObjInsHvacPort1  As Object
    Dim ObjInsBendPort1Side  As Object
    Dim ObjInsHvacPort3  As Object
    Dim ObjInsBendPort2Side  As Object
    Dim ObjInsTangentPort2Side  As Object
    Dim ObjInsReducerPort2Side  As Object
    Dim ObjInsHvacPort2  As Object
    Dim ObjInsPNoz1  As Object
    Dim ObjInsPNoz2  As Object
    
    Dim oPartFclt      As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth_port      As Double

    Dim Width As Double, depth As Double
    Dim BWidth As Double, BDepth As Double
    Dim Width2 As Double
    Dim HLength As Double
    Dim BLength As Double, FlangeWidth As Double
    Dim NozzleLength As Double, Thickness As Double
    Dim Angle As Double
    Dim NozzleOffset As Double
    Dim FacetoCenter As Double
    Dim NozzleSpacing As Double
    Dim InsulationThickness As Double
        
    'assign to meaningful variables from the input array
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    BWidth = Width
    Width2 = arrayOfInputs(3)
    HLength = arrayOfInputs(4)
    BLength = arrayOfInputs(5)
    NozzleOffset = arrayOfInputs(6)
    FacetoCenter = arrayOfInputs(7)
    NozzleSpacing = arrayOfInputs(8)
    InsulationThickness = arrayOfInputs(9)
  

 ' Insert your code for output 4(Header insulation)
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
' Insert your code for output  InsHvacPort1
    Dim dPort1toCenter As Double
    
    Dim dTangentPort1SideLength As Double
    Dim dTotalLengthOfTwoBends As Double
    Dim dTangentPort2SideLength As Double
    Dim dReducerPort2SideLength As Double
    
    'Following dimensions are taken from Fincantieri part drawing
    dTangentPort1SideLength = 0.155
    dTotalLengthOfTwoBends = 0.715
    dTangentPort2SideLength = 0.1
    dReducerPort2SideLength = 0.34
    dPort1toCenter = 0.6025

    stPoint.Set -dPort1toCenter, 0, 0
    enPoint.Set -dTotalLengthOfTwoBends / 2, 0, 0

    Set ObjInsHvacPort1 = PlaceCylinder(m_outputColl, stPoint, enPoint, Width + 2 * InsulationThickness, True)

'   Set the output
    m_outputColl.AddOutput "InsHvacPort1", ObjInsHvacPort1
    Set ObjInsHvacPort1 = Nothing

' Insert your code for output  InsBendPort1Side
    Dim circleCenter    As AutoMath.DPosition
    Dim circleNormal    As AutoMath.DVector
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set circleCenter = New AutoMath.DPosition
    circleCenter.Set -dTotalLengthOfTwoBends / 2, 0, 0
    Set circleNormal = New AutoMath.DVector
    circleNormal.Set 1, 0, 0
    
' Construct a circle that will be used to project the disc
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.X, circleCenter.Y, circleCenter.Z, _
                        circleNormal.X, circleNormal.Y, circleNormal.Z, _
                        Width / 2 + InsulationThickness)
    
    Dim axisVector As AutoMath.DVector
    Dim centerPoint As AutoMath.DPosition
    Dim revAngle As Double
    
    Set axisVector = New AutoMath.DVector
    Set centerPoint = New AutoMath.DPosition

'   Assumption: revAngle is taken as 45 deg and BLength
    Dim PI        As Double

    PI = 4 * Atn(1)
    revAngle = PI / 4
    'Port 3 is at 35 degrees angle from the Y axis as per Fincantieri part drawing
    Angle = 35 * PI / 180

    axisVector.Set 0, -Sin(Angle), Cos(Angle)
    centerPoint.Set -dTotalLengthOfTwoBends / 2, BLength * Cos(Angle), BLength * Sin(Angle)

    Set ObjInsBendPort1Side = geomFactory.Revolutions3d.CreateByCurve( _
                                                    m_outputColl.ResourceManager, _
                                                    objCircle, _
                                                    axisVector.X, axisVector.Y, axisVector.Z, _
                                                    centerPoint.X, centerPoint.Y, centerPoint.Z, _
                                                    revAngle, False)
                                                    
'   Set the output
    m_outputColl.AddOutput "InsBendPort1Side", ObjInsBendPort1Side
    Set ObjInsBendPort1Side = Nothing

' Insert your code for output InsHvacPort3
    stPoint.Set 0, BLength * Cos(Angle), BLength * Sin(Angle)
'   Length of Nozzle 3 is BLength / Sqr(2): Sin(45deg) is component of branch length upto elbow top
'   So Origin to the start of the nozzle is BLength * (Sqr(2) -1) /Sqr(2)
    Dim dOriginToStartOfNozzle3 As Double
    dOriginToStartOfNozzle3 = BLength * (Sqr(2) - 1) / Sqr(2)
    enPoint.Set 0, dOriginToStartOfNozzle3 * Cos(Angle), dOriginToStartOfNozzle3 * Sin(Angle)

    Set ObjInsHvacPort3 = PlaceCylinder(m_outputColl, stPoint, enPoint, Width + 2 * InsulationThickness, True)

'   Set the output
    m_outputColl.AddOutput "InsHvacPort3", ObjInsHvacPort3
    Set ObjInsHvacPort3 = Nothing

' Insert your code for output  InsBendPort2Side
    circleCenter.Set dTotalLengthOfTwoBends / 2, 0, 0
    circleNormal.Set 1, 0, 0
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.X, circleCenter.Y, circleCenter.Z, _
                        circleNormal.X, circleNormal.Y, circleNormal.Z, _
                        Width / 2 + InsulationThickness)
    axisVector.Set 0, Sin(Angle), -Cos(Angle)
    centerPoint.Set dTotalLengthOfTwoBends / 2, BLength * Cos(Angle), BLength * Sin(Angle)
    Set ObjInsBendPort2Side = geomFactory.Revolutions3d.CreateByCurve( _
                                                    m_outputColl.ResourceManager, _
                                                    objCircle, _
                                                    axisVector.X, axisVector.Y, axisVector.Z, _
                                                    centerPoint.X, centerPoint.Y, centerPoint.Z, _
                                                    revAngle, False)
                                                    
'   Set the output
    m_outputColl.AddOutput "InsBendPort2Side", ObjInsBendPort2Side
    Set ObjInsBendPort2Side = Nothing

' Insert your code for output  InsTangentPort2Side
    stPoint.Set dTotalLengthOfTwoBends / 2, 0, 0
    enPoint.Set stPoint.X + dTangentPort2SideLength, stPoint.Y, stPoint.Z

    Set ObjInsTangentPort2Side = PlaceCylinder(m_outputColl, stPoint, enPoint, Width + 2 * InsulationThickness, True)
    
'   Set the output
    m_outputColl.AddOutput "InsTangentPort2Side", ObjInsTangentPort2Side
    Set ObjInsTangentPort2Side = Nothing

' Insert your code for output  InsReducerPort2Side
    circleCenter.Set dTotalLengthOfTwoBends / 2 + dTangentPort2SideLength, 0, 0
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.X, circleCenter.Y, circleCenter.Z, _
                        circleNormal.X, circleNormal.Y, circleNormal.Z, _
                        Width / 2 + InsulationThickness)
                        
    Dim objCircle2 As IngrGeom3D.Circle3d
    circleCenter.Set dTotalLengthOfTwoBends / 2 + dTangentPort2SideLength + dReducerPort2SideLength, -NozzleOffset, 0
    Set objCircle2 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.X, circleCenter.Y, circleCenter.Z, _
                        circleNormal.X, circleNormal.Y, circleNormal.Z, _
                        Width2 / 2 + InsulationThickness)

    Set ObjInsReducerPort2Side = geomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                                                            objCircle, objCircle2, True)
'   Set the output
    m_outputColl.AddOutput "InsReducerPort2Side", ObjInsReducerPort2Side
    Set ObjInsReducerPort2Side = Nothing

' Insert your code for output  InsHvacPort2
    Dim dInsulationDiameter As Double
    ' Assumption: Flange diameter is taken as 0.3429m (Corresponding to 200 mm pipe size) and Flange thickness is assumed
    ' to be 0.022m. So Nozzle Length is taken as 0.025 m which is greater than the Flange thickness.
    dInsulationDiameter = 0.3429 + 2 * InsulationThickness
    
    stPoint.Set dTotalLengthOfTwoBends / 2 + dTangentPort2SideLength + dReducerPort2SideLength, -NozzleOffset, 0
    enPoint.Set stPoint.X - 0.022 - InsulationThickness, stPoint.Y, stPoint.Z

    Set ObjInsHvacPort2 = PlaceCylinder(m_outputColl, stPoint, enPoint, Width + 2 * InsulationThickness, True)
    
'   Set the output
    m_outputColl.AddOutput "InsHvacPort2", ObjInsHvacPort2
    Set ObjInsHvacPort2 = Nothing

' Insert your code for output  InsPNoz1
'   Note: Existing RetrieveParameters function in basGeom3D fails if the first port is a HVAC port. So we
'   made a local function RetrievePipePortParameters with a modified logic.
    RetrievePipePortParameters 4, oPartFclt, m_outputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth_port

    NozzleLength = FacetoCenter
    If CmpDblLessThan(NozzleLength, flangeThick) Then NozzleLength = flangeThick

    stPoint.Set NozzleSpacing / 2, 0, -FacetoCenter
    enPoint.Set stPoint.X, stPoint.Y, stPoint.Z + NozzleLength
    dInsulationDiameter = flangeDiam + 2 * InsulationThickness
    If CmpDblLessThan(flangeDiam, pipeDiam) Then dInsulationDiameter = pipeDiam + 2 * InsulationThickness

    Set ObjInsPNoz1 = PlaceCylinder(m_outputColl, stPoint, enPoint, dInsulationDiameter, True)
    
'   Set the output
    m_outputColl.AddOutput "InsPNoz1", ObjInsPNoz1
    Set ObjInsPNoz1 = Nothing

' Insert your code for output  InsPNoz2
    RetrievePipePortParameters 5, oPartFclt, m_outputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth_port

    If CmpDblLessThan(NozzleLength, flangeThick) Then NozzleLength = flangeThick
    
    stPoint.Set -NozzleSpacing / 2, 0, -FacetoCenter
    enPoint.Set stPoint.X, stPoint.Y, stPoint.Z + NozzleLength
    dInsulationDiameter = flangeDiam + 2 * InsulationThickness
    If CmpDblLessThan(flangeDiam, pipeDiam) Then dInsulationDiameter = pipeDiam + 2 * InsulationThickness

    Set ObjInsPNoz2 = PlaceCylinder(m_outputColl, stPoint, enPoint, dInsulationDiameter, True)
    
'   Set the output
    m_outputColl.AddOutput "InsPNoz2", ObjInsPNoz2
    Set ObjInsPNoz2 = Nothing

    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Private Function RetrievePipePortParameters(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                ByRef lPipeDiam As Double, ByRef lFlangeThick As Double, _
                                ByRef lFlangeDiam As Double, ByRef lcptoffset As Double, _
                                ByRef lDepth As Double, _
                                Optional lBoltedPartDimIncludesFlgFaceProj = True)

    'Dim oPipePort       As GSCADNozzleEntities.IJDPipePort
    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    
    Set oCollection = partInput.GetNozzles()
    
    For pPortIndex = oCollection.Size To 1 Step -1
        Set oPipePort = oCollection.Item(pPortIndex)
            If oPipePort.PortIndex = index Then
                lPipeDiam = oPipePort.PipingOutsideDiameter
                lFlangeThick = oPipePort.FlangeOrHubThickness
                lFlangeDiam = oPipePort.FlangeOrHubOutsideDiameter
                lcptoffset = oPipePort.FlangeProjectionOrSocketOffset
                If lBoltedPartDimIncludesFlgFaceProj = True And _
                                        oPipePort.TerminationClass = 5 Then
                    lcptoffset = 0
                End If
                lDepth = oPipePort.SeatingOrGrooveOrSocketDepth
                Exit For
            End If
    Next pPortIndex
    
    Set oPipePort = Nothing
    
End Function

'''NOT USED
Public Function RetrieveParametersofPipePort(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                ByRef lPipeDiam As Double, ByRef lFlangeThick As Double, _
                                ByRef lFlangeDiam As Double, ByRef lcptoffset As Double, _
                                ByRef lDepth As Double)
'''
End Function
'''NOT USED
