VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HRndPantWye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   HRndPantWye.cls
'   SP3DHRndPantWye.HRndPantWye
'   Author:         kkk
'   Creation Date:  Friday 17, March 2006
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       HVAC Round Pant Wye. One Leg is fixed at 45 Degree and another at any available angle
'       The Symbol is taken from Sprial Catalog.  Part Number is Y-2.  This symbol code is prepared for Both Y-2 and Y-2-R
'       Both the legs of the Pant Wye can be of Reducing types and can be changed by the User.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   10.12.2007     PS      CR-132021  Modified the default value of BranchAngle
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Const NEGLIGIBLE_THICKNESS = 0.0001

Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Dim PI As Double
    PI = 4 * Atn(1)
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHRndPantWye.HRndPantWye"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
  On Error GoTo ErrorHandler
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHRndPantWye.HRndPantWye"
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

    ' Feed SP3DHRndPantWye.HRndPantWye
    ' Inputs:
    '          1. "HWidth"
    '          2. "BWidth1"
    '          3. "BWidth2"
    '          4. "BLength"
    '          5. "Branch Angle"
    
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
    Dim Inputs(1 To 6) As IMSSymbolEntities.DInput
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
    Inputs(2).Name = "BWidth1"
    Inputs(2).Description = "Branch 1 Width"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.08
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).Name = "BWidth2"
    Inputs(3).Description = "Branch 2 Width"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.08
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).Name = "BLength"
    Inputs(4).Description = "Length of the Pant WYe along the Branch"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(4).DefaultParameterValue = PC
    
    Set Inputs(5) = New IMSSymbolEntities.DInput
    Inputs(5).Name = "BranchAngle"
    Inputs(5).Description = "Angle of Branch 1"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = Atn(1)
    Inputs(5).DefaultParameterValue = PC
    
    Set Inputs(6) = New IMSSymbolEntities.DInput
    Inputs(6).Name = "InsulationThickness"
    Inputs(6).Description = "Insulation Thickness"
    Inputs(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(6).DefaultParameterValue = PC
    
    
    For index = 1 To 6
        InputsIf.SetInput Inputs(index), index + 1
        Set Inputs(index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 16) As IMSSymbolEntities.DOutput
    
    For index = 1 To 16
        Set O(index) = New IMSSymbolEntities.DOutput
        O(index).Properties = 0
    Next
   
    O(1).Name = "HvacNozzle1"
    O(1).Description = "HvacPort1 of Round Pant Wye"
    
    O(2).Name = "HvacNozzle2"
    O(2).Description = "HvacPort2 of Round Pant Wye"
    
    O(3).Name = "HvacNozzle3"
    O(3).Description = "HvacPort2 of Round Pant Wye"
    
    O(4).Name = "ObjBody1"
    O(4).Description = "Body 1 of Round Pant Wye"
    
    O(5).Name = "ObjBody2"
    O(5).Description = "Body 2 of Round Pant Wye"
    
    O(6).Name = "ObjBody3"
    O(6).Description = "Body 3 of Round Pant Wye"
    
    O(7).Name = "ObjTakeOff1"
    O(7).Description = "TakeOff 1 of Round Pant Wye"
    
    O(8).Name = "ObjTakeOff2"
    O(8).Description = "TakeOff 2 of Round Pant Wye"
    
    O(9).Name = "ObjTakeOff3"
    O(9).Description = "TakeOff 3 of Round Pant Wye"
    
    O(10).Name = "ObjLeg2"
    O(10).Description = "Leg 2 of Round Pant Wye"
    
    O(11).Name = "ObjLeg3"
    O(11).Description = "Leg 3 of Round Pant Wye"
    
       
'Symbol Definition Initialization of Insulation Aspect
            
    O(12).Name = "HeaderIns"
    O(12).Description = "Header Insulation"
    
    O(13).Name = "Branch1Ins"
    O(13).Description = "Branch1 Insulation"
    
    O(14).Name = "Branch2Ins"
    O(14).Description = "Branch2 Insulation"
    
    O(15).Name = "TakeOff2Ins"
    O(15).Description = "TakeOff2 Insulation"

    O(16).Name = "TakeOff3Ins"
    O(16).Description = "TakeOff3 Insulation"
    
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
        
    For index = 1 To 11
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
    PhysicalRepEval.ProgId = "SP3DHRndPantWye.HRndPantWye"

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
    For index = 12 To 16
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
    InsulationRepEval.ProgId = "SP3DHRndPantWye.HRndPantWye"
    
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
'CREATION OF PHYSICAL REPRESENTATION OF HVAC Round Pant Wye
'=========================================================================
Private Sub Physical(ByRef arrayOfInputs())
    Const METHOD = "Physical"
    On Error GoTo ErrorLabel
   
    Dim oPartFclt As PartFacelets.IJDPart
    Dim Width As Double
    Dim BWidth1 As Double
    Dim BWidth2 As Double
    Dim BLength As Double
    Dim FlangeWidth As Double
    Dim NozzleLength As Double
    Dim Thickness As Double
    Dim InsulationThickness As Double
    Dim CornerRadius As Double
    Dim BranchAngle As Double
    
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oBaseCircle As IngrGeom3D.Circle3d
    Dim oTopCircle1 As IngrGeom3D.Circle3d
    Dim oTopCircle2 As IngrGeom3D.Circle3d
    
        
    Dim PI As Double
    PI = 4 * Atn(1)
    
    
    'assign to meaningful variables from the input array

    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    BWidth1 = arrayOfInputs(3)
    BWidth2 = arrayOfInputs(4)
    BLength = arrayOfInputs(5)
    BranchAngle = arrayOfInputs(6)
    InsulationThickness = arrayOfInputs(7)
    
    
    Dim theta As Double
    theta = BranchAngle
    
    
'============================================================
'BUILD HVAC NOZZLE -1 ON BASIC ASSEMBLY - Header Pipe of Round Pant Wye
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
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Width, CornerRadius, DimBaseOuter, PortStatus, _
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
'BUILD TakeOff 1 Of Round Pant Wye
'============================================
    Dim ObjTakeOff1 As Object
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    

    stPoint.Set 0, 0, 0
    enPoint.Set -TakeOff, 0, 0
    
    Set ObjTakeOff1 = PlaceCylinder(m_outputColl, stPoint, enPoint, Width * 1.01, False)
'   Set the output
    m_outputColl.AddOutput "ObjTakeOff1", ObjTakeOff1
    
    Set ObjTakeOff1 = Nothing
    
'============================================
'BUILD Body 1 Of Round Pant Wye
'============================================

    Dim ObjBody1 As Object
    stPoint.Set -4 / 1000, 0, 0
    enPoint.Set 4 / 1000, 0, 0
    
    Set ObjBody1 = PlaceCylinder(m_outputColl, stPoint, enPoint, Width * 1.06, True)
    
    
'   Set the output
    m_outputColl.AddOutput "ObjBody1", ObjBody1
    Set ObjBody1 = Nothing
    
    
'============================================
'BUILD HVACNOZZLE2 Of Round Pant Wye
'============================================
    iPortIndex = 2
    FlowDir = DistribFlow_BOTH
         Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, BWidth2, _
                                            BWidth2, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    
    NozzleLength = 0.001
    X = (25.4 / 1000) + BLength * Cos(theta)
    Y = BLength * Sin(theta)
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    
    dir.Set Cos(theta), Sin(theta), 0
    iDistribPort.SetDirectionVector dir
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    

'============================================
'BUILD Leg 2 Of HVAC Round Pant Wye (Upper Leg)
'============================================
    Dim ObjLeg2 As IngrGeom3D.RuledSurface3d
    
        X = (25.4 / 1000)
        Y = 0#
        Z = 0#
    
    Set oBaseCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, X, 0, 0, 1, 0, 0, Width / 2)
    Set oTopCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, X + BLength * Cos(theta), Y + BLength * Sin(theta), Z, Cos(theta), Sin(theta), 0, BWidth2 / 2)
    
    Set ObjLeg2 = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                                                oBaseCircle, oTopCircle1, False)

'   Set the output
    m_outputColl.AddOutput "ObjLeg2", ObjLeg2
    Set oTopCircle1 = Nothing
    Set ObjLeg2 = Nothing
    
'==================================================
'BUILD Body 2 Of Round Pant Wye (Upward Leg's Ring)
'==================================================
    Dim ObjBody2 As Object
    
    X = (25.4 / 1000) + BLength * Cos(theta)
    Y = BLength * Sin(theta)
    Z = 0#
    
    stPoint.Set X - (4 / 1000) * Cos(theta), Y - (4 / 1000) * Sin(theta), Z
    enPoint.Set X + (4 / 1000) * Cos(theta), Y + (4 / 1000) * Sin(theta), Z
    Set ObjBody2 = PlaceCylinder(m_outputColl, stPoint, enPoint, BWidth2 * 1.06, True)
    
'   Set the output
    m_outputColl.AddOutput "ObjBody2", ObjBody2
    Set ObjBody2 = Nothing
    
'============================================
'BUILD TakeOff 2 Of Round Pant Wye
'============================================
    Dim ObjTakeOff2 As Object

    X = (25.4 / 1000) + BLength * Cos(theta)
    Y = BLength * Sin(theta)
    Z = 0#
    
    stPoint.Set X, Y, Z
    enPoint.Set X + TakeOff * Cos(theta), Y + TakeOff * Sin(theta), Z
    Set ObjTakeOff2 = PlaceCylinder(m_outputColl, stPoint, enPoint, BWidth2 * 1.01, False)
'   Set the output
    m_outputColl.AddOutput "ObjTakeOff2", ObjTakeOff2
    
    Set ObjTakeOff2 = Nothing
    
    
'============================================
'BUILD Leg 3 Of HVAC Round Pant Wye (Lower Leg)
'============================================
    Dim ObjLeg3 As IngrGeom3D.RuledSurface3d
    
        X = (25.4 / 1000)
        Y = 0#
        Z = 0#
    Set oTopCircle2 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, X + BLength * Cos(PI / 4), -(Y + BLength * Sin(PI / 4)), Z, Cos(PI / 4), -Sin(PI / 4), 0, BWidth1 / 2)
    
    Set ObjLeg3 = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                                                oBaseCircle, oTopCircle2, False)

'   Set the output
    m_outputColl.AddOutput "ObjLeg3", ObjLeg3
    Set oTopCircle2 = Nothing
    Set oBaseCircle = Nothing
    Set ObjLeg3 = Nothing
    
    
'============================================
'BUILD Ring 3 Of Round Pant Wye
'============================================
    Dim ObjBody3 As Object
    
    X = (25.4 / 1000) + (BLength * Cos(PI / 4))
    Y = -(BLength * Sin(PI / 4))
    Z = 0#
    
    stPoint.Set X - (4 / 1000) * Cos(PI / 4), Y + (4 / 1000) * Sin(PI / 4), Z
    enPoint.Set X + (4 / 1000) * Cos(PI / 4), Y - (4 / 1000) * Sin(PI / 4), Z
    Set ObjBody3 = PlaceCylinder(m_outputColl, stPoint, enPoint, BWidth1 * 1.06, True)
    
'   Set the output
    m_outputColl.AddOutput "ObjBody3", ObjBody3
    Set ObjBody3 = Nothing
    
    
'============================================
'BUILD TakeOff 3 Of Round Pant Wye
'============================================
    Dim ObjTakeOff3 As Object
    
    X = (25.4 / 1000) + (BLength * Cos(PI / 4))
    Y = -(BLength * Sin(PI / 4))
    Z = 0#
    
    stPoint.Set X, Y, Z
    enPoint.Set X + TakeOff * Cos(PI / 4), Y - TakeOff * Sin(PI / 4), Z
    Set ObjTakeOff3 = PlaceCylinder(m_outputColl, stPoint, enPoint, BWidth1 * 1.01, False)
'   Set the output
    m_outputColl.AddOutput "ObjTakeOff3", ObjTakeOff3
    
    Set ObjTakeOff3 = Nothing
    
'============================================
'BUILD HVACNOZZLE3 Of Round Pant Wye
'============================================
    iPortIndex = 3
    FlowDir = DistribFlow_BOTH
         Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, BWidth1, _
                                            BWidth1, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
'Position of the nozzle should be the conenct point of the nozzle
    NozzleLength = 0.001
    
    X = (25.4 / 1000) + (BLength * Cos(PI / 4))
    Y = -BLength * Sin(PI / 4)
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set Cos(PI / 4), -Sin(PI / 4), 0
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
    Dim oBaseCircle As IngrGeom3D.Circle3d
    Dim oTopCircle1 As IngrGeom3D.Circle3d
    Dim oTopCircle2 As IngrGeom3D.Circle3d
    

Dim oPartFclt As PartFacelets.IJDPart
    Dim Width As Double
    Dim BWidth1 As Double
    Dim BWidth2 As Double
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
    BWidth1 = arrayOfInputs(3)
    BWidth2 = arrayOfInputs(4)
    BLength = arrayOfInputs(5)
    BranchAngle = arrayOfInputs(6)
    InsulationThickness = arrayOfInputs(7)
    
    
    Dim PI As Double
    PI = 4 * Atn(1)
    
    
    Dim theta As Double
    theta = BranchAngle
    TakeOff = 2 * 25.4 / 1000 ' 2 Inches

'============================================
'BUILD Header Duct Insulation
'============================================
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition


    stPoint.Set -TakeOff, 0, 0
    enPoint.Set 25.4 / 1000, 0, 0

    Dim ObjHeaderIns As Object

    Set ObjHeaderIns = PlaceCylinder(m_outputColl, stPoint, enPoint, _
                 Width * 1.01 + 2 * InsulationThickness, True)

    m_outputColl.AddOutput "HeaderIns", ObjHeaderIns
    Set ObjHeaderIns = Nothing

'============================================
'BUILD Insulation of Branch 2 (Upper Leg)
'============================================
    Dim ObjBranch1Ins As IngrGeom3D.RuledSurface3d
    
        X = (25.4 / 1000)
        Y = 0#
        Z = 0#
    
    Set oBaseCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, X, 0, 0, 1, 0, 0, (Width / 2) + InsulationThickness)
    Set oTopCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, X + BLength * Cos(theta), Y + BLength * Sin(theta), Z, Cos(theta), Sin(theta), 0, (BWidth2 / 2) + InsulationThickness)
    
    Set ObjBranch1Ins = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                                                oBaseCircle, oTopCircle1, False)

'   Set the output
    m_outputColl.AddOutput "Branch1Ins", ObjBranch1Ins
    Set oTopCircle1 = Nothing
    Set ObjBranch1Ins = Nothing
    
'============================================
'BUILD Insulation of Branch 3 (Lower Leg)
'============================================

    Dim ObjBranch2Ins As IngrGeom3D.RuledSurface3d
    
        X = (25.4 / 1000)
        Y = 0#
        Z = 0#
    
    Set oTopCircle2 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, X + BLength * Cos(PI / 4), -(Y + BLength * Sin(PI / 4)), Z, Cos(PI / 4), -Sin(PI / 4), 0, (BWidth1 / 2) + InsulationThickness)
    
    Set ObjBranch2Ins = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                                                oBaseCircle, oTopCircle2, False)

'   Set the output
    m_outputColl.AddOutput "Branch2Ins", ObjBranch2Ins
    Set oTopCircle2 = Nothing
    Set ObjBranch2Ins = Nothing
    Set oBaseCircle = Nothing
    
' Insert your code for output 7(TakeOff-2 insulation)

    Dim ObjTakeOff2Ins As Object

    X = (25.4 / 1000) + BLength * Cos(theta)
    Y = BLength * Sin(theta)
    Z = 0#
    
    stPoint.Set X, Y, Z
    enPoint.Set X + TakeOff * Cos(theta), Y + TakeOff * Sin(theta), Z
    Set ObjTakeOff2Ins = PlaceCylinder(m_outputColl, stPoint, enPoint, (BWidth2 * 1.01) + 2 * InsulationThickness, False)
'   Set the output
    m_outputColl.AddOutput "TakeOff2Ins", ObjTakeOff2Ins
    
    Set ObjTakeOff2Ins = Nothing
    
' Insert your code for output 8(TakeOff-3 insulation)

    Dim ObjTakeOff3Ins As Object

    X = (25.4 / 1000) + BLength * Cos(PI / 4)
    Y = -BLength * Sin(PI / 4)
    Z = 0#
    
    stPoint.Set X, Y, Z
    enPoint.Set X + TakeOff * Cos(PI / 4), -(Y + TakeOff * Sin(PI / 4)), Z
    Set ObjTakeOff3Ins = PlaceCylinder(m_outputColl, stPoint, enPoint, (BWidth1 * 1.01) + 2 * InsulationThickness, False)
'   Set the output
    m_outputColl.AddOutput "TakeOff3Ins", ObjTakeOff3Ins
    
    Set ObjTakeOff3Ins = Nothing
    
    Exit Sub
    
ErrorLabel:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
