VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HROUNDFFDWInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'File
'    HROUNDFFDWInc.cls
'
'   ProgID:         SP3DHRoundFlatFlangeDWInc.HROUNDFFDWInc
'   Author:         svsmylav
'   Creation Date:  Wednesday, Dec 18 2002
'
'Description
'    Definition of HVAC Round Flat Flange with Depth and Width Inclination.
'
'Notes
'
'   Reference - existing symbol 'RoundFlatFlange' is used.
'
'History:
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   15.DEC.2004      MS  CR 67971 Corrected port indices and flow-directions in HVAC Symbols
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   10.12.2007     PS      CR-132021  Modified the default value of Angle and Rotation
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_outputColl As IJDOutputCollection

Const NEGLIGIBLE_THICKNESS = 0.0001

Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHRoundFlatFlangeDWInc.HROUNDFFDWInc"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  Dim oSymbolFactory As New imssymbolentities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As imssymbolentities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHRoundFlatFlangeDWInc.HROUNDFFDWInc"
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


Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As imssymbolentities.IJDSymbolDefinition)
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
  
    On Error GoTo ErrorHandler
  
    ' Set the input to the definition
    Dim InputsIf As imssymbolentities.IJDInputs
    Set InputsIf = pSymbolDefinition

    'Following indented lines will be removed very soon. This has been
    'included to improve the performance. These lines will be removed as soon as
    'symbol subsytem allows this to be removed from Symbol Definition .- Raju 05/05/99.
        ' set the CMCacheForPart method in the definition
        
    ' Create a new input by new operator
    Dim Inputs(1 To 8) As imssymbolentities.DInput
    Dim Index As Integer
  
    ' Create a defaultValue
    Dim PC As imssymbolentities.DParameterContent
    Set PC = New imssymbolentities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
    
    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    Set Inputs(1) = New imssymbolentities.DInput
    Inputs(1).Name = "Width"
    Inputs(1).Description = "Width of the Rect Duct"
    Inputs(1).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.3
    Inputs(1).DefaultParameterValue = PC
        
    Set Inputs(2) = New imssymbolentities.DInput
    Inputs(2).Name = "Clearance"
    Inputs(2).Description = "Clearance between ducts"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.008
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New imssymbolentities.DInput
    Inputs(3).Name = "Thickness"
    Inputs(3).Description = "Thickness of the Flat Flange"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.06
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New imssymbolentities.DInput
    Inputs(4).Name = "FlangeWidth"
    Inputs(4).Description = "Width of the Flat Flange"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.04
    Inputs(4).DefaultParameterValue = PC
    
    Set Inputs(5) = New imssymbolentities.DInput
    Inputs(5).Name = "CompressedGasket"
    Inputs(5).Description = "Compressed Gasket Thickness"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.004
    Inputs(5).DefaultParameterValue = PC
    
    Set Inputs(6) = New imssymbolentities.DInput
    Inputs(6).Name = "Angle"
    Inputs(6).Description = "Angle"
    Inputs(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = (23 / 90) * Atn(1)
    Inputs(6).DefaultParameterValue = PC
    
    Set Inputs(7) = New imssymbolentities.DInput
    Inputs(7).Name = "Rotation"
    Inputs(7).Description = "Rotation"
    Inputs(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = (23 / 90) * Atn(1)
    Inputs(7).DefaultParameterValue = PC
    
    Set Inputs(8) = New imssymbolentities.DInput
    Inputs(8).Name = "InsulationThickness"
    Inputs(8).Description = "Insulation thickness of body"
    Inputs(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(8).DefaultParameterValue = PC
       
    For Index = 1 To 8
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 4) As imssymbolentities.DOutput
    
    For Index = 1 To 4
        Set O(Index) = New imssymbolentities.DOutput
        O(Index).Properties = 0
    Next
   
    O(1).Name = "HvacNozzle1"
    O(1).Description = "HvacPort1 of Flat Flange"
    
    O(2).Name = "HvacNozzle2"
    O(2).Description = "HvacPort2 of Flat Flange"
    
    O(3).Name = "DGFlange"
    O(3).Description = "Dumb Graphics Flange"
        
    O(4).Name = "InsDGFlange"
    O(4).Description = "Insulation for Dumb Graphics Flanges"
    
    'Define the representation "Physical"
    Dim rep1 As imssymbolentities.DRepresentation
    Set rep1 = New imssymbolentities.DRepresentation
  
    rep1.Name = "Physical"
    rep1.Description = "Physical Represntation of the HVAC Round Flat Flange With Depth Inclination"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim oRepPhysicalOutputs As imssymbolentities.IJDOutputs
    Set oRepPhysicalOutputs = rep1
        
    For Index = 1 To 3
        oRepPhysicalOutputs.SetOutput O(Index)
        Set O(Index) = Nothing
    Next

    'Set the representation to definition
    Dim RepsIf As imssymbolentities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1
   
    'Define evaluation for Physical representation
    Dim PhysicalRepEval As DRepresentationEvaluation
    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "script for the Physical representation"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "SP3DHRoundFlatFlangeDWInc.HROUNDFFDWInc"

    
'   Redefine outputs
    oRepPhysicalOutputs.RemoveAllOutput
    Set oRepPhysicalOutputs = Nothing

'  Insulation representation
    rep1.Name = "Insulation"
    rep1.Description = "Insulation of the Round Flat Flange with Depth Inclination"
    'Set the repID to Insulation. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = Insulation
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    
    Dim oRepInsulationOutputs As imssymbolentities.IJDOutputs
    Set oRepInsulationOutputs = rep1
    
'   Set the output
    oRepInsulationOutputs.SetOutput O(4)
    Set O(4) = Nothing
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1
   
' Define the evaluation associated to the Insulation representation
    Dim InsulationRepEval As DRepresentationEvaluation
    Set InsulationRepEval = New DRepresentationEvaluation
    InsulationRepEval.Name = "Insulation"
    InsulationRepEval.Description = "Insulation representation"
    InsulationRepEval.Properties = igREPRESENTATION_HIDDEN
    InsulationRepEval.Type = igREPRESENTATION_VBFUNCTION
    InsulationRepEval.ProgId = "SP3DHRoundFlatFlangeDWInc.HROUNDFFDWInc"

    'Set the evaluations for the Physical and Insulation representation on the definition
    Dim RepEvalsIf As imssymbolentities.IJDRepresentationEvaluations
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    RepEvalsIf.AddRepresentationEvaluation InsulationRepEval
    
    Set rep1 = Nothing
    Set oRepInsulationOutputs = Nothing
    Set RepsIf = Nothing
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
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False

End Sub


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
    ElseIf StrComp(repName, "Insulation") = 0 Then
        InsulationRep arrayOfInputs
    End If
End Sub

Private Sub Physical(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    
    
    Dim Width As Double, Depth As Double
    Dim Clearance As Double
    Dim Thickness As Double
    Dim FlangeWidth As Double
    Dim NozzleLength As Double
    Dim CornerRadius As Double
    Dim CompressedGasket As Double
    Dim Angle As Double
    Dim Rotation As Double
    Dim InsulationThickness As Double
    
    'assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    Clearance = arrayOfInputs(3)
    Thickness = arrayOfInputs(4)
    FlangeWidth = arrayOfInputs(5)
    CompressedGasket = arrayOfInputs(6)
    Angle = arrayOfInputs(7)        'Measured in anti-clockwise direction from positive Y-axis
    Rotation = arrayOfInputs(8)     'Measured in anti-clockwise direction from positive Z-axis
    InsulationThickness = arrayOfInputs(9)
    
'   Cross section is circular
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
    
    Dim iPortIndex As Integer

    'Set HVAC nozzle parameters
    iPortIndex = 1
    PortDepth = 0#
    CptOffset = CompressedGasket / 2
    EndPrep = 11
    CornerRadius = 0#
    Dim NozzleFlangeThickness As Double
    NozzleFlangeThickness = 0
    NozzleLength = NozzleFlangeThickness + NEGLIGIBLE_THICKNESS
    'NozzleLength Has to be GREATER than NozzleFlangeThickness
    
    FlowDir = DistribFlow_BOTH

    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Dim FlangeWidth1 As Double
    FlangeWidth1 = 0
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            NozzleFlangeThickness, FlangeWidth1, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac1", PortDepth, CptOffset, True, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = -CptOffset
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
'=====================================
'BUILD HVACNOZZLE2 Of Flat Flange
'=====================================
    iPortIndex = 2
    EndPrep = 301
    FlowDir = DistribFlow_BOTH
    FlangeWidth1 = 0
    CptOffset = 0
    NozzleFlangeThickness = 0
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            NozzleFlangeThickness, FlangeWidth1, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac2", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    
    
    X = Clearance
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
    iNozzle.Length = NEGLIGIBLE_THICKNESS

    m_outputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
'   Dumb graphics placement
'   Place DGFlange
    Dim circPoint   As New AutoMath.DPosition
    Dim NormalVector As New AutoMath.DVector
    Dim FlangeDia As Double
    Dim ObjDGFlange As Object
    
    CptOffset = CompressedGasket / 2
    
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    Dim objCircle       As IngrGeom3D.Circle3d
    
    circPoint.Set -CptOffset, 0, 0
    NormalVector.Set Cos(Angle) * Cos(Rotation), Sin(Angle), -Cos(Angle) * Sin(Rotation)
    
'   Note: Dot product between the NormalVector and unit vector along X-axis(i.e. 1,0,0) is
'   Cos(Angle) * Cos(Rotation, which is the Cosine of the effective angle between the
'   above two vectors. Flange diameter in the depth inclination case is
'   (Width / (Cos(Angle) + 2 * FlangeWidth), and in the current case it is
'   (Width / (Cos(Angle) * Cos(Rotation)) + 2 * FlangeWidth) respectively.
'   So Cosine of the effective angle is the divisor in the current case.

    FlangeDia = Width / (Cos(Angle) * Cos(Rotation)) + 2 * FlangeWidth
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circPoint.X, circPoint.Y, circPoint.Z, _
                        NormalVector.X, NormalVector.Y, NormalVector.Z, _
                        FlangeDia / 2)
'   Project along the normal vector of the flange face
    Set ObjDGFlange = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, _
                                                        objCircle, _
                                                        -NormalVector.X, -NormalVector.Y, -NormalVector.Z, _
                                                        Thickness, True)
'   Set the output
    m_outputColl.AddOutput "DGFlange", ObjDGFlange
    Set ObjDGFlange = Nothing
    
    Set circPoint = Nothing
    Set NormalVector = Nothing
    Set geomFactory = Nothing

'   remove curve
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCircle
    sp3dElem1.Remove
    Set objCircle = Nothing

    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "HROUNDFFDWInc definition"
        Debug.Assert False
End Sub
'=========================================================================
'CREATION OF INSULATION REPRESENTATION OF HRECTFF
'=========================================================================


Private Sub InsulationRep(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    
    Dim ObjInsDGFlange As Object
    
    Dim Width As Double
    Dim Clearance As Double
    Dim Thickness As Double
    Dim FlangeWidth As Double
    Dim CompressedGasket As Double
    Dim Angle As Double
    Dim Rotation As Double
    Dim InsulationThickness As Double

    Dim Depth As Double

    Dim NozzleLength As Double
    Dim CornerRadius As Double
    
    'assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    Clearance = arrayOfInputs(3)
    Thickness = arrayOfInputs(4)
    FlangeWidth = arrayOfInputs(5)
    CompressedGasket = arrayOfInputs(6)
    Angle = arrayOfInputs(7)        'Measured in anti-clockwise direction from positive Y-axis
    Rotation = arrayOfInputs(8)     'Measured in anti-clockwise direction from positive Z-axis
    InsulationThickness = arrayOfInputs(9)
    Depth = Width
    
    Dim CptOffset As Double
    CptOffset = CompressedGasket / 2

'  Prepare ObjInsDGFlange
    Dim parInsulationDiameter As Double
    Dim circPoint   As New AutoMath.DPosition
    Dim NormalVector As New AutoMath.DVector
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
        
    NormalVector.Set Cos(Angle) * Cos(Rotation), Sin(Angle), -Cos(Angle) * Sin(Rotation)
    circPoint.Set -CptOffset - (Thickness + InsulationThickness) * NormalVector.X, _
                        -(Thickness + InsulationThickness) * NormalVector.Y, _
                        -(Thickness + InsulationThickness) * NormalVector.Z

    parInsulationDiameter = Width / (Cos(Angle) * Cos(Rotation)) + _
                                2 * (FlangeWidth + InsulationThickness)
    
    Dim objCircle       As IngrGeom3D.Circle3d
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circPoint.X, circPoint.Y, circPoint.Z, _
                        NormalVector.X, NormalVector.Y, NormalVector.Z, _
                        parInsulationDiameter / 2)
                        
    Set ObjInsDGFlange = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, _
                                objCircle, NormalVector.X, NormalVector.Y, NormalVector.Z, _
                                2 * (CptOffset + Thickness + InsulationThickness), True)
                                                                                                        
'   Set the output
    m_outputColl.AddOutput "InsDGFlange", ObjInsDGFlange
    Set ObjInsDGFlange = Nothing
    Set circPoint = Nothing
    Set NormalVector = Nothing
    Set geomFactory = Nothing

'   remove curve
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCircle
    sp3dElem1.Remove
    Set objCircle = Nothing
    
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "HRECTFF definition"
        Debug.Assert False
End Sub

