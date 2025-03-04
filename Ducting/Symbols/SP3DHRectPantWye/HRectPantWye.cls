VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HRectPantWye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   HRectPantWye.cls
'   SP3DHRectPantWye.HRectPantWye
'   Author:         kkk
'   Creation Date:  Friday 17, March 2006
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       HVAC Rectangular Pant Wye. One Leg is fixed at 45 Degree and another at any available angle
'       The Symbol is taken from Sprial Catalog.  Part Number is Y-2.  This symbol code is prepared for Both Y-2 and Y-2-R
'       Both the legs of the Pant Wye can be of Reducing types and can be changed by the User.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC     DI-95670  Replace names with initials in all revision history sheets and symbols
'  10.12.2007      PS      CR-132021  Modified the default value of BranchAngle
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Physical:" 'Used for error messages

Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Dim PI As Double
    PI = 4 * Atn(1)
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHRectPantWye.HRectPantWye"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
  On Error GoTo ErrorHandler
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHRectPantWye.HRectPantWye"
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

    ' Feed SP3DHRectPantWye.HRectPantWye
    ' Inputs:
    '               1.Width
    '               2.Depth
    '               3.BWidth1
    '               4.BDepth1
    '               5.BWidth2
    '               6.BDepth2
    '               7.BLength
    '               8.BranchAngle
    '               9.InsulationThickness
    
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
    Dim Inputs(1 To 9) As IMSSymbolEntities.DInput
    Dim index As Integer
  
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue

    Set Inputs(1) = New IMSSymbolEntities.DInput
    Inputs(1).Name = "Width"
    Inputs(1).Description = "Width of the Header"
    Inputs(1).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Inputs(1).DefaultParameterValue = PC
    
    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).Name = "Depth"
    Inputs(2).Description = "Depth of the Header"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Inputs(2).DefaultParameterValue = PC

    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).Name = "BWidth1"
    Inputs(3).Description = "Branch 1 Width"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.08
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).Name = "BDepth1"
    Inputs(4).Description = "Branch 1 Depth"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.08
    Inputs(4).DefaultParameterValue = PC
    
    Set Inputs(5) = New IMSSymbolEntities.DInput
    Inputs(5).Name = "BWidth2"
    Inputs(5).Description = "Branch 1 Width"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.08
    Inputs(5).DefaultParameterValue = PC
    
    Set Inputs(6) = New IMSSymbolEntities.DInput
    Inputs(6).Name = "BDepth2"
    Inputs(6).Description = "Branch 1 Depth"
    Inputs(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.08
    Inputs(6).DefaultParameterValue = PC
    
    Set Inputs(7) = New IMSSymbolEntities.DInput
    Inputs(7).Name = "BLength"
    Inputs(7).Description = "Length of the Pant WYe along the Branch"
    Inputs(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(7).DefaultParameterValue = PC
    
    Set Inputs(8) = New IMSSymbolEntities.DInput
    Inputs(8).Name = "BranchAngle"
    Inputs(8).Description = "Angle of Branch 1"
    Inputs(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = Atn(1)
    Inputs(8).DefaultParameterValue = PC
    
    Set Inputs(9) = New IMSSymbolEntities.DInput
    Inputs(9).Name = "InsulationThickness"
    Inputs(9).Description = "Insulation Thickness"
    Inputs(9).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(9).DefaultParameterValue = PC
    
    
    For index = 1 To 9
        InputsIf.SetInput Inputs(index), index + 1
        Set Inputs(index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 9) As IMSSymbolEntities.DOutput
    
    For index = 1 To 9
        Set O(index) = New IMSSymbolEntities.DOutput
        O(index).Properties = 0
    Next
   
    O(1).Name = "HvacNozzle1"
    O(1).Description = "HvacPort1 of Rectangular Pant Wye"
    
    O(2).Name = "HvacNozzle2"
    O(2).Description = "HvacPort2 of Rectangular Pant Wye"
    
    O(3).Name = "HvacNozzle3"
    O(3).Description = "HvacPort2 of Rectangular Pant Wye"
    
    O(4).Name = "HeaderTakeoff"
    O(4).Description = "Take Off of Header portion of Rectangular Pant Wye"
    
    O(5).Name = "ObjLeg2"
    O(5).Description = "Leg Portion of Branch 1 of Rectangular Pant Wye"
    
    O(6).Name = "ObjLeg3"
    O(6).Description = "Leg Portion of Branch 2 of Rectangular Pant Wye"

'Symbol Definition Initialization of Insulation Aspect
    
    O(7).Name = "HeaderIns"
    O(7).Description = "Header Insulation"
    
    O(8).Name = "ObjLeg2Ins"
    O(8).Description = "Branch 1 Insulation"
    
    O(9).Name = "ObjLeg3Ins"
    O(9).Description = "Branch 2 Insulation"
    
       
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
        
    For index = 1 To 6
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
    PhysicalRepEval.ProgId = "SP3DHRectPantWye.HRectPantWye"

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
    For index = 7 To 9
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
    InsulationRepEval.ProgId = "SP3DHRectPantWye.HRectPantWye"
    
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
'CREATION OF PHYSICAL REPRESENTATION OF HVAC Rectangular Pant Wye
'=========================================================================
Private Sub Physical(ByRef arrayOfInputs())
    Const METHOD = "Physical"
    On Error GoTo ErrorLabel
   
    Dim oPartFclt As PartFacelets.IJDPart
    Dim Width As Double
    Dim Depth As Double
    
    Dim BWidth1 As Double
    Dim BDepth1 As Double
    
    Dim BWidth2 As Double
    Dim BDepth2 As Double
    
    Dim BLength As Double
    Dim FlangeWidth As Double
    Dim NozzleLength As Double
    Dim Thickness As Double
    Dim InsulationThickness As Double
    Dim CornerRadius As Double
    Dim BranchAngle As Double
    
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    
    Dim PI As Double
    PI = 4 * Atn(1)
    
    
    'assign to meaningful variables from the input array

    Set oPartFclt = arrayOfInputs(1)
    
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    
    BWidth1 = arrayOfInputs(4)
    BDepth1 = arrayOfInputs(5)
    
    BWidth2 = arrayOfInputs(6)
    BDepth2 = arrayOfInputs(7)
    
    BLength = arrayOfInputs(8)
    BranchAngle = arrayOfInputs(9)
    InsulationThickness = arrayOfInputs(10)
    
    
    Dim theta As Double
    theta = BranchAngle
    
     Width = Width * 1.01
    Depth = Depth * 1.01
'============================================================
'BUILD HVAC NOZZLE -1 ON BASIC ASSEMBLY - Header Pipe of Rectangular Pant Wye
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
    
    Dim stPoint As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim minHL As Double
    Dim minBL As Double
    Dim TakeOff As Double
' As per the Spiral Catalog the Take Off Length is considered as 2 Inches
    TakeOff = 2 * 25.4 / 1000 ' 2 Inches
    
    'Set HVAC nozzle parameters
    iPortIndex = 1
    EndPrep = 301
    CornerRadius = 0#
    
    NozzleLength = 25.4 / 1000 '1 Inch is given as Nozzle Length according to the catalog
    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    FlangeWidth = 0#
    Thickness = 0.0001
    PortDepth = 0#
    CptOffset = 0#
    FlowDir = DistribFlow_BOTH

    ' Width and Depth of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double

    X = 0#
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
'============================================
'BUILD Header Duct Take Off Portion
'============================================

' Insert your code for output 4(Hvac Header Insulation)
    
    Dim ObjHeaderTakeoff As Object
 
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    dX = -TakeOff
    dY = Depth / 2
    dZ = Width / 2

    Dim dlinepts(0 To 14) As Double

    'Point 1
    dlinepts(0) = dX
    dlinepts(1) = dY
    dlinepts(2) = dZ
    'Point 2
    dlinepts(3) = dX
    dlinepts(4) = dY
    dlinepts(5) = -dZ
    'Point 3
    dlinepts(6) = dX
    dlinepts(7) = -dY
    dlinepts(8) = -dZ
    'Point 4
    dlinepts(9) = dX
    dlinepts(10) = -dY
    dlinepts(11) = dZ
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)

    Dim oAxisVec As New AutoMath.DVector
    oAxisVec.Set 1, 0, 0
    
    Set ObjHeaderTakeoff = PlaceProjectedBox(m_outputColl, dlinepts, oAxisVec, TakeOff, True)

'   Set the output

    m_outputColl.AddOutput "HeaderTakeoff", ObjHeaderTakeoff
    Set ObjHeaderTakeoff = Nothing
        
'============================================
'BUILD HVACNOZZLE2 Of Rectangular Pant Wye
'============================================
    iPortIndex = 2
    FlowDir = DistribFlow_BOTH
         Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, BWidth1, _
                                            BDepth1, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    
    NozzleLength = 0.001
    X = (25.4 / 1000) + BLength * Cos(theta)
    Y = 0#
    Z = BLength * Sin(theta)
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    
    dir.Set Cos(theta), 0, Sin(theta)
    iDistribPort.SetDirectionVector dir
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    

'============================================
'BUILD Leg 2 Of HVAC Rectangular Pant Wye (Upper Leg)
'============================================
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    
    Dim objBaseCStr         As IngrGeom3D.ComplexString3d
    Dim objTopCStr1         As IngrGeom3D.ComplexString3d
    Dim objTopCStr2         As IngrGeom3D.ComplexString3d
    
    
    Dim ObjLeg2 As IngrGeom3D.RuledSurface3d
    
    Dim component1 As Double
    Dim alpha1 As Double
    
    Dim component2 As Double
    Dim alpha2 As Double
    
    X = (25.4 / 1000)
    Y = -Depth / 2
    Z = -Width / 2
    
    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, X, -Depth / 2, -Width / 2, X, -Depth / 2, Width / 2)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, X, -Depth / 2, Width / 2, X, Depth / 2, Width / 2)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, X, Depth / 2, Width / 2, X, Depth / 2, -Width / 2)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, X, Depth / 2, -Width / 2, X, -Depth / 2, -Width / 2)
    lines.Add oLine
    
    stPoint.Set X, Y, Z
    Set objBaseCStr = PlaceTrCString(stPoint, lines)
        
    component1 = Sqr(BLength * BLength + (BWidth1 / 2) * (BWidth1 / 2))
    alpha1 = Atn((BWidth1 / 2) / BLength)
    
    Dim p(1 To 4)   As New AutoMath.DPosition
    
    p(1).Set (25.4 / 1000) + component1 * Cos(alpha1 + theta), BDepth1 / 2, component1 * Sin(alpha1 + theta)
    p(2).Set (25.4 / 1000) + component1 * Cos(theta - alpha1), BDepth1 / 2, component1 * Sin(theta - alpha1)
    p(3).Set (25.4 / 1000) + component1 * Cos(theta - alpha1), -BDepth1 / 2, component1 * Sin(theta - alpha1)
    p(4).Set (25.4 / 1000) + component1 * Cos(alpha1 + theta), -BDepth1 / 2, component1 * Sin(alpha1 + theta)
    
    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(1).X, p(1).Y, p(1).Z, p(2).X, p(2).Y, p(2).Z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(2).X, p(2).Y, p(2).Z, p(3).X, p(3).Y, p(3).Z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(3).X, p(3).Y, p(3).Z, p(4).X, p(4).Y, p(4).Z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(4).X, p(4).Y, p(4).Z, p(1).X, p(1).Y, p(1).Z)
    lines.Add oLine
    
    stPoint.Set p(1).X, p(1).Y, p(1).Z
    Set objTopCStr1 = PlaceTrCString(stPoint, lines)
    
    Set ObjLeg2 = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                                           objBaseCStr, objTopCStr1, False)
'    Set the output
    m_outputColl.AddOutput "ObjLeg2", ObjLeg2
    
    Set objTopCStr1 = Nothing
    Set ObjLeg2 = Nothing
    
    
'=========================================================
'BUILD Leg 3 Of HVAC Rectangular Pant Wye (Lower Leg)
'==========================================================
    Dim ObjLeg3 As IngrGeom3D.RuledSurface3d
    
    component2 = Sqr(BLength * BLength + (BWidth2 / 2) * (BWidth2 / 2))
    alpha2 = Atn((BWidth2 / 2) / BLength)
    

    p(1).Set (25.4 / 1000) + component2 * Cos(alpha2 + PI / 4), -BDepth2 / 2, -component2 * Sin(alpha2 + PI / 4)
    p(2).Set (25.4 / 1000) + component2 * Cos(alpha2 + PI / 4), BDepth2 / 2, -component2 * Sin(alpha2 + PI / 4)
    p(3).Set (25.4 / 1000) + component2 * Cos(PI / 4 - alpha2), BDepth2 / 2, -component2 * Sin(PI / 4 - alpha2)
    p(4).Set (25.4 / 1000) + component2 * Cos(PI / 4 - alpha2), -BDepth2 / 2, -component2 * Sin(PI / 4 - alpha2)
    
    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(1).X, p(1).Y, p(1).Z, p(2).X, p(2).Y, p(2).Z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(2).X, p(2).Y, p(2).Z, p(3).X, p(3).Y, p(3).Z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(3).X, p(3).Y, p(3).Z, p(4).X, p(4).Y, p(4).Z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(4).X, p(4).Y, p(4).Z, p(1).X, p(1).Y, p(1).Z)
    lines.Add oLine
    
    stPoint.Set p(1).X, p(1).Y, p(1).Z
    Set objTopCStr2 = PlaceTrCString(stPoint, lines)
    
    Set ObjLeg3 = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                                           objBaseCStr, objTopCStr2, False)
'    Set the output
    m_outputColl.AddOutput "ObjLeg3", ObjLeg3
    
    Set objTopCStr2 = Nothing
    Set ObjLeg3 = Nothing
    Set oLine = Nothing
    Set objBaseCStr = Nothing
        
    Dim iCount As Integer
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set lines = Nothing
    
'============================================
'BUILD HVACNOZZLE3 Of Rectangular Pant Wye
'============================================
    iPortIndex = 3
    FlowDir = DistribFlow_BOTH
         Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, BWidth2, _
                                            BDepth2, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
'Position of the nozzle should be the conenct point of the nozzle
    NozzleLength = 0.001
    
    X = (25.4 / 1000) + (BLength * Cos(PI / 4))
    Y = 0#
    Z = -BLength * Sin(PI / 4)
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set Cos(PI / 4), 0, -Sin(PI / 4)
    iDistribPort.SetDirectionVector dir
    
    'dir.Set 0, 1, 0
    'iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "HvacNozzle3", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
    
    
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
    
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim X As Double, Y As Double, Z As Double


    Dim oPartFclt As PartFacelets.IJDPart
    Dim Width As Double
    Dim Depth As Double
    
    Dim BWidth1 As Double
    Dim BDepth1 As Double
    
    Dim BWidth2 As Double
    Dim BDepth2 As Double
    
    Dim BLength As Double
    Dim FlangeWidth As Double
    Dim NozzleLength As Double
    Dim Thickness As Double
    Dim InsulationThickness As Double
    Dim CornerRadius As Double
    Dim BranchAngle As Double
    Dim TakeOff As Double
    
    
'Assign to meaningful variables from the input array
    
    Set oPartFclt = arrayOfInputs(1)
    
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    
    BWidth1 = arrayOfInputs(4)
    BDepth1 = arrayOfInputs(5)
    
    BWidth2 = arrayOfInputs(6)
    BDepth2 = arrayOfInputs(7)
    
    BLength = arrayOfInputs(8)
    BranchAngle = arrayOfInputs(9)
    InsulationThickness = arrayOfInputs(10)
    
    Dim WidthIns As Double
    Dim DepthIns As Double
    
    Dim BWidth1Ins As Double
    Dim BDepth1Ins As Double
    
    Dim BWidth2Ins As Double
    Dim BDepth2Ins As Double
    
    WidthIns = WidthIns * 1.01
    DepthIns = DepthIns * 1.01
    
    WidthIns = Width + 2 * InsulationThickness
    DepthIns = Depth + 2 * InsulationThickness
    
    BWidth1Ins = BWidth1 + 2 * InsulationThickness
    BDepth1Ins = BDepth1 + 2 * InsulationThickness
    
    BWidth2Ins = BWidth2 + 2 * InsulationThickness
    BDepth2Ins = BDepth2 + 2 * InsulationThickness
    
    Dim PI As Double
    PI = 4 * Atn(1)
    
    
    Dim theta As Double
    theta = BranchAngle
    
    TakeOff = 2 * 25.4 / 1000 ' 2 Inches

'============================================
'BUILD Header Duct Insulation
'============================================

' Insert your code for output 4(Hvac Header Insulation)
 Dim ObjHeaderIns As Object
 
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    dX = -TakeOff
    dY = DepthIns / 2
    dZ = WidthIns / 2
    Dim dlinepts(0 To 14) As Double

    'Point 1
    dlinepts(0) = dX
    dlinepts(1) = dY
    dlinepts(2) = dZ
    'Point 2
    dlinepts(3) = dX
    dlinepts(4) = dY
    dlinepts(5) = -dZ
    'Point 3
    dlinepts(6) = dX
    dlinepts(7) = -dY
    dlinepts(8) = -dZ
    'Point 4
    dlinepts(9) = dX
    dlinepts(10) = -dY
    dlinepts(11) = dZ
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)

    Dim oAxisVec As New AutoMath.DVector
    oAxisVec.Set 1, 0, 0
    
    Set ObjHeaderIns = PlaceProjectedBox(m_outputColl, dlinepts, oAxisVec, TakeOff + (25.4 / 1000), True)

'   Set the output

    m_outputColl.AddOutput "HeaderIns", ObjHeaderIns
    Set ObjHeaderIns = Nothing
    
'============================================
'BUILD Insulation of Branch 2 (Upper Leg)
'============================================
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim objBaseCStr         As IngrGeom3D.ComplexString3d
    Dim objTopCStr1         As IngrGeom3D.ComplexString3d
    Dim objTopCStr2         As IngrGeom3D.ComplexString3d
    Dim stPoint As New AutoMath.DPosition
    
    Dim ObjLeg2Ins As IngrGeom3D.RuledSurface3d
    
    Dim component1 As Double
    Dim alpha1 As Double
    
    Dim component2 As Double
    Dim alpha2 As Double
    X = (25.4 / 1000)
    Y = -DepthIns / 2
    Z = -WidthIns / 2
    
        Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, X, -DepthIns / 2, -WidthIns / 2, X, -DepthIns / 2, WidthIns / 2)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, X, -DepthIns / 2, WidthIns / 2, X, DepthIns / 2, WidthIns / 2)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, X, DepthIns / 2, WidthIns / 2, X, DepthIns / 2, -WidthIns / 2)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, X, DepthIns / 2, -WidthIns / 2, X, -DepthIns / 2, -WidthIns / 2)
    lines.Add oLine
    
    stPoint.Set X, Y, Z
    Set objBaseCStr = PlaceTrCString(stPoint, lines)
        
    component1 = Sqr(BLength * BLength + BWidth1Ins / 2 * BWidth1Ins / 2)
    alpha1 = Atn((BWidth1Ins / 2) / BLength)
    
    component2 = Sqr(BLength * BLength + (BWidth2Ins / 2) * (BWidth2Ins / 2))
    alpha2 = Atn((BWidth2Ins / 2) / BLength)
    
    Dim p(1 To 4)   As New AutoMath.DPosition
    
    p(1).Set (25.4 / 1000) + component1 * Cos(alpha1 + theta), BDepth1Ins / 2, component1 * Sin(alpha1 + theta)
    p(2).Set (25.4 / 1000) + component1 * Cos(theta - alpha1), BDepth1Ins / 2, component1 * Sin(theta - alpha1)
    p(3).Set (25.4 / 1000) + component1 * Cos(theta - alpha1), -BDepth1Ins / 2, component1 * Sin(theta - alpha1)
    p(4).Set (25.4 / 1000) + component1 * Cos(alpha1 + theta), -BDepth1Ins / 2, component1 * Sin(alpha1 + theta)
    
    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(1).X, p(1).Y, p(1).Z, p(2).X, p(2).Y, p(2).Z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(2).X, p(2).Y, p(2).Z, p(3).X, p(3).Y, p(3).Z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(3).X, p(3).Y, p(3).Z, p(4).X, p(4).Y, p(4).Z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(4).X, p(4).Y, p(4).Z, p(1).X, p(1).Y, p(1).Z)
    lines.Add oLine
    
    stPoint.Set p(1).X, p(1).Y, p(1).Z
    Set objTopCStr1 = PlaceTrCString(stPoint, lines)
    
    Set ObjLeg2Ins = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                                           objBaseCStr, objTopCStr1, False)
'    Set the output
    m_outputColl.AddOutput "ObjLeg2Ins", ObjLeg2Ins
    
    Set objTopCStr1 = Nothing
    Set ObjLeg2Ins = Nothing
'============================================
'BUILD Insulation of Branch 3 (Lower Leg)
'============================================

    Dim ObjLeg3Ins As IngrGeom3D.RuledSurface3d

    p(1).Set (25.4 / 1000) + component2 * Cos(alpha2 + PI / 4), -BDepth2Ins / 2, -component2 * Sin(alpha2 + PI / 4)
    p(2).Set (25.4 / 1000) + component2 * Cos(alpha2 + PI / 4), BDepth2Ins / 2, -component2 * Sin(alpha2 + PI / 4)
    p(3).Set (25.4 / 1000) + component2 * Cos(PI / 4 - alpha2), BDepth2Ins / 2, -component2 * Sin(PI / 4 - alpha2)
    p(4).Set (25.4 / 1000) + component2 * Cos(PI / 4 - alpha2), -BDepth2Ins / 2, -component2 * Sin(PI / 4 - alpha2)
    
    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(1).X, p(1).Y, p(1).Z, p(2).X, p(2).Y, p(2).Z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(2).X, p(2).Y, p(2).Z, p(3).X, p(3).Y, p(3).Z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(3).X, p(3).Y, p(3).Z, p(4).X, p(4).Y, p(4).Z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, p(4).X, p(4).Y, p(4).Z, p(1).X, p(1).Y, p(1).Z)
    lines.Add oLine
    
    stPoint.Set p(1).X, p(1).Y, p(1).Z
    Set objTopCStr2 = PlaceTrCString(stPoint, lines)
    
    Set ObjLeg3Ins = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                                           objBaseCStr, objTopCStr2, False)
'    Set the output
    m_outputColl.AddOutput "ObjLeg3Ins", ObjLeg3Ins
    
    Set objTopCStr2 = Nothing
    Set ObjLeg3Ins = Nothing
    Set oLine = Nothing
    Set objBaseCStr = Nothing
        
    Dim iCount As Integer
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    Set lines = Nothing
       
    Exit Sub
    
ErrorLabel:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Private Function PlaceProjectedBox(ByVal objOutputColl As Object, _
                        dPoints() As Double, _
                        ByVal axisVector As AutoMath.DVector, _
                        height As Double, _
                        isCapped As Boolean) As Object

    Const METHOD = "PlaceProjectedBox:"
    On Error GoTo ErrorHandler

    Dim objProjection   As IngrGeom3D.Projection3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim oLineStr As IngrGeom3D.LineString3d

    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
    Set objProjection = geomFactory.Projections3d.CreateByCurve( _
                                                    objOutputColl.ResourceManager, _
                                                    oLineStr, _
                                                    axisVector.X, axisVector.Y, axisVector.Z, _
                                                    height, isCapped)

    Set PlaceProjectedBox = objProjection
    Set objProjection = Nothing
    Set oLineStr = Nothing
    Set geomFactory = Nothing
 
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Function


'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub
Public Sub ReportUnanticipatedError(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub
