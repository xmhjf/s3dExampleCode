VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HRndCross"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   HRndCross.cls
'   SP3DHRndCross.HRndCross
'   Author:         kkk
'   Creation Date:  Friday 17, March 2006
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      HVAC Round Cross. The Symbol is taken from Lindab Catalog.  Part Number is XCU and this component is called in the Catalog
'      as 'X-Piece'.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  29.Jan.2008     VRK  TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Const NEGLIGIBLE_THICKNESS = 0.0001
Const LINEAR_TOLERANCE = 0.0000001

Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = "SP3DHRndCross.HRndCross"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler
    Dim oSymbolFactory As New imssymbolentities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As imssymbolentities.DSymbolDefinition

    Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = "SP3DHRndCross.HRndCross"
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


Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As imssymbolentities.IJDSymbolDefinition)
' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    ' Feed GSCADHRoundFlatFlange.HROUNDFF Definition
    ' Inputs:
    '          1. "HWidth"
    '          2. "BWidth"
    '          3. "HLength"
    '          4. "BLength"

    '
    ' The representation "Physical" is set to the definition
    ' Physical :    (RepresentationId = 1)

    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler

    ' Set the input to the definition
    Dim InputsIf As imssymbolentities.IJDInputs
    Set InputsIf = pSymbolDefinition

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    ' Create a new input by new operator
    Dim Inputs(1 To 5) As imssymbolentities.DInput
    Dim Index As Integer

    ' Create a defaultValue
    Dim PC As imssymbolentities.DParameterContent
    Set PC = New imssymbolentities.DParameterContent    'not persistent PC

    PC.Type = igValue

    Set Inputs(1) = New imssymbolentities.DInput
    Inputs(1).Name = "Width"
    Inputs(1).Description = "Width of the Header"
    Inputs(1).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Inputs(1).DefaultParameterValue = PC

    Set Inputs(2) = New imssymbolentities.DInput
    Inputs(2).Name = "BWidth"
    Inputs(2).Description = "Branch Width"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.08
    Inputs(2).DefaultParameterValue = PC

    Set Inputs(3) = New imssymbolentities.DInput
    Inputs(3).Name = "HLength"
    Inputs(3).Description = "Length of the Tee along the Header"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.15
    Inputs(3).DefaultParameterValue = PC

    Set Inputs(4) = New imssymbolentities.DInput
    Inputs(4).Name = "BLength"
    Inputs(4).Description = "Length of the Tee along the Branch"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.08
    Inputs(4).DefaultParameterValue = PC

    Set Inputs(5) = New imssymbolentities.DInput
    Inputs(5).Name = "InsulationThickness"
    Inputs(5).Description = "Insulation Thickness"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(5).DefaultParameterValue = PC

    For Index = 1 To 5
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 14) As imssymbolentities.DOutput

    For Index = 1 To 14
        Set O(Index) = New imssymbolentities.DOutput
        O(Index).Properties = 0
    Next

    O(1).Name = "HvacNozzle1"
    O(1).Description = "HvacPort 1 of Round Cross"

    O(2).Name = "HvacNozzle2"
    O(2).Description = "HvacPort 2 of Round Cross"

    O(3).Name = "HvacNozzle3"
    O(3).Description = "HvacPort 3 of Round Cross"

    O(4).Name = "HvacNozzle4"
    O(4).Description = "HvacPort 4 of Round Cross"

    O(5).Name = "ObjTakeOff1"
    O(5).Description = "Take-Off 1 of Round Cross"

    O(6).Name = "ObjTakeOff2"
    O(6).Description = "Take-Off 2 of Round Cross"

    O(7).Name = "ObjTakeOff3"
    O(7).Description = "Take-Off 3 of Round Cross"

    O(8).Name = "ObjTakeOff4"
    O(8).Description = "Take-Off 4 of Round Cross"

    O(9).Name = "ObjRing1"
    O(9).Description = "Outer Ring 1 of Round Cross"

    O(10).Name = "ObjRing2"
    O(10).Description = "Outer Ring 2 of Round Cross"

    O(11).Name = "ObjRing3"
    O(11).Description = "Outer Ring 3 of Round Cross"

    O(12).Name = "ObjRing4"
    O(12).Description = "Outer Ring 4 of Round Cross"

    O(13).Name = "HeaderIns"
    O(13).Description = "Header Insulation"

    O(14).Name = "BranchIns"
    O(14).Description = "Branch Insulation"

    'Define the representation "Physical"
    Dim rep1 As imssymbolentities.DRepresentation
    Set rep1 = New imssymbolentities.DRepresentation

    rep1.Name = "Physical"
    rep1.Description = "Physical Represntation of the Air Distrib Assembly"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim oRepPhysicalOutputs As imssymbolentities.IJDOutputs
    Set oRepPhysicalOutputs = rep1

    For Index = 1 To 12
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
    PhysicalRepEval.ProgId = "SP3DHRndCross.HRndCross"

    'Set the evaluations for the Physical representation on the definition
    Dim RepEvalsIf As imssymbolentities.IJDRepresentationEvaluations
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

    Dim oRepInsulationOutputs As imssymbolentities.IJDOutputs
    Set oRepInsulationOutputs = rep1

    'Set the output
    For Index = 13 To 14
        oRepInsulationOutputs.SetOutput O(Index)
        Set O(Index) = Nothing
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
    InsulationRepEval.ProgId = "SP3DHRndCross.HRndCross"

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
'CREATION OF PHYSICAL REPRESENTATION OF HVAC Round Cross
'=========================================================================

Private Sub Physical(ByRef arrayOfInputs())
    Const METHOD = "Physical"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim Width As Double, Depth As Double
    Dim BWidth As Double, BDepth As Double
    Dim HLength As Double, halfHLength As Double
    Dim BLength As Double, FlangeWidth As Double
    Dim NozzleLength As Double, Thickness As Double
    Dim TakeOff As Double
    Dim CornerRadius As Double

    'assign to meaningful variables from the input array
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    BWidth = arrayOfInputs(3)
    HLength = arrayOfInputs(4)
    BLength = arrayOfInputs(5)
    Depth = Width
    BDepth = BWidth
    TakeOff = 2 * 25.4 / 1000    'Considering Take-Off as 2 Inches on all the Four sides of the Cross

    '============================================
    'BUILD TakeOff 1 Of Round Cross
    '============================================
    Dim ObjTakeOff1 As Object
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    stPoint.Set -HLength / 2 - TakeOff, 0, 0
    enPoint.Set -HLength / 2, 0, 0

    Set ObjTakeOff1 = PlaceCylinder(m_outputColl, stPoint, enPoint, Width * 1.01, False)
    '   Set the output
    m_outputColl.AddOutput "ObjTakeOff1", ObjTakeOff1

    Set ObjTakeOff1 = Nothing
    '============================================
    'BUILD OuterRing 1 Of Round Cross
    '============================================
    Dim ObjRing1 As Object

    stPoint.Set -HLength / 2 - 4 / 1000, 0, 0
    enPoint.Set -HLength / 2 + 4 / 1000, 0, 0

    Set ObjRing1 = PlaceCylinder(m_outputColl, stPoint, enPoint, Width * 1.06, True)

    '   Set the output
    m_outputColl.AddOutput "ObjRing1", ObjRing1

    Set ObjRing1 = Nothing
    '============================================
    'BUILD Ring 2 Of Round Cross
    '============================================
    Dim ObjRing2 As Object

    stPoint.Set HLength / 2 - 4 / 1000, 0, 0
    enPoint.Set HLength / 2 + 4 / 1000, 0, 0
    Set ObjRing2 = PlaceCylinder(m_outputColl, stPoint, enPoint, Width * 1.06, True)

    '   Set the output
    m_outputColl.AddOutput "ObjRing2", ObjRing2
    Set ObjRing2 = Nothing

    '============================================
    'BUILD TakeOff 2 Of Lateral
    '============================================
    Dim ObjTakeOff2 As Object

    stPoint.Set HLength / 2, 0, 0
    enPoint.Set HLength / 2 + TakeOff, 0, 0
    Set ObjTakeOff2 = PlaceCylinder(m_outputColl, stPoint, enPoint, Width * 1.01, False)
    '   Set the output
    m_outputColl.AddOutput "ObjTakeOff2", ObjTakeOff2

    Set ObjTakeOff2 = Nothing
    '============================================
    'BUILD Ring 3 Of Round Cross
    '============================================
    Dim ObjRing3 As Object

    stPoint.Set 0, BLength - (4 / 1000), 0
    enPoint.Set 0, BLength + (4 / 1000), 0
    Set ObjRing3 = PlaceCylinder(m_outputColl, stPoint, enPoint, BWidth * 1.06, True)

    '   Set the output
    m_outputColl.AddOutput "ObjRing3", ObjRing3
    Set ObjRing3 = Nothing

    '============================================
    'BUILD TakeOff 3 Of Lateral
    '============================================
    Dim ObjTakeOff3 As Object

    stPoint.Set 0, BLength, 0
    enPoint.Set 0, BLength + TakeOff, 0
    Set ObjTakeOff3 = PlaceCylinder(m_outputColl, stPoint, enPoint, BWidth * 1.01, False)
    '   Set the output
    m_outputColl.AddOutput "ObjTakeOff3", ObjTakeOff3

    Set ObjTakeOff3 = Nothing
    '============================================
    'BUILD Ring 4 Of Round Cross
    '============================================
    Dim ObjRing4 As Object

    stPoint.Set 0, -(BLength - 4 / 1000), 0
    enPoint.Set 0, -(BLength + 4 / 1000), 0
    Set ObjRing4 = PlaceCylinder(m_outputColl, stPoint, enPoint, BWidth * 1.06, True)

    '   Set the output
    m_outputColl.AddOutput "ObjRing4", ObjRing4
    Set ObjRing4 = Nothing

    '============================================
    'BUILD TakeOff 3 Of Lateral
    '============================================
    Dim ObjTakeOff4 As Object

    stPoint.Set 0, -BLength, 0
    enPoint.Set 0, -(BLength + TakeOff), 0
    Set ObjTakeOff4 = PlaceCylinder(m_outputColl, stPoint, enPoint, BWidth * 1.01, False)
    '   Set the output
    m_outputColl.AddOutput "ObjTakeOff4", ObjTakeOff4

    Set ObjTakeOff4 = Nothing

    '============================================================
    'BUILD HVAC NOZZLE -1 ON BASIC ASSEMBLY Round Cross
    '============================================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim nozzleIndex As Integer
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iLogicalDistPort As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort

    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    
    'HVAC parameters for Nozzle
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim iPortIndex As Integer
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 4) As Long
    Dim dThickness(1 To 4) As Double
    Dim dFlangeWidth(1 To 4) As Double
    Dim lFlowDir(1 To 4) As Long
    Dim dPortDepth(1 To 4) As Double
    Dim dCptOffSet(1 To 4) As Double
    
    CornerRadius = 0#
    
    'Set HVAC nozzle parameters from catalog
    Set oHvacColl = oPartFclt.GetNozzles()
    For iPortIndex = 1 To oHvacColl.Size
        Set oHvacPort = oHvacColl.Item(iPortIndex)
        lEndPrep(iPortIndex) = oHvacPort.EndPrep
        dThickness(iPortIndex) = oHvacPort.Thickness
        dFlangeWidth(iPortIndex) = oHvacPort.FlangeWidth
        lFlowDir(iPortIndex) = oHvacPort.FlowDirection
        dPortDepth(iPortIndex) = oHvacPort.PortDepth
        dCptOffSet(iPortIndex) = oHvacPort.CptOffset
    Next iPortIndex
        
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    
    halfHLength = HLength / 2
    NozzleLength = HLength

    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia
    iPortIndex = 1
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                            GSCADNozzleEntities.Round, lEndPrep(1), _
                            dThickness(1), dFlangeWidth(1), lFlowDir(1), Width, _
                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                            "HNoz1", dPortDepth(1), dCptOffSet(1), False, _
                            m_outputColl.ResourceManager)

    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = -halfHLength
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
    'BUILD HVACNOZZLE2 Of Round Cross
    '==================================================
    iPortIndex = 2
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                            GSCADNozzleEntities.Round, lEndPrep(2), _
                            dThickness(2), dFlangeWidth(2), lFlowDir(2), Width, _
                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                            "HNoz2", dPortDepth(2), dCptOffSet(2), False, _
                            m_outputColl.ResourceManager)

    'Position of the nozzle should be the conenct point of the nozzle
    NozzleLength = 0
    X = halfHLength
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

    m_outputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    '============================================
    'BUILD HVACNOZZLE3 Of Round Cross
    '============================================
    iPortIndex = 3
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                            GSCADNozzleEntities.Round, lEndPrep(3), _
                            dThickness(3), dFlangeWidth(3), lFlowDir(3), BWidth, _
                            BDepth, CornerRadius, DimBaseOuter, PortStatus, _
                            "HNoz3", dPortDepth(3), dCptOffSet(3), False, _
                            m_outputColl.ResourceManager)

    'Position of the nozzle should be the conenct point of the nozzle
    NozzleLength = BLength
    X = 0#
    Y = BLength
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 0, 1, 0
    iDistribPort.SetDirectionVector dir

    'dir.Set 0, 1, 0
    'iDistribPort.SetRadialOrient dir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "HvacNozzle3", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    '============================================
    'BUILD HVACNOZZLE 4 Of Round Cross
    '============================================
    iPortIndex = 4
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                            GSCADNozzleEntities.Round, lEndPrep(4), _
                            dThickness(4), dFlangeWidth(4), lFlowDir(4), BWidth, _
                            BDepth, CornerRadius, DimBaseOuter, PortStatus, _
                            "HNoz4", dPortDepth(4), dCptOffSet(4), False, _
                            m_outputColl.ResourceManager)
                            
    'Position of the nozzle should be the conenct point of the nozzle
    NozzleLength = BLength
    X = 0#
    Y = -BLength
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 0, -1, 0
    iDistribPort.SetDirectionVector dir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "HvacNozzle4", oHvacNozzle
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

    Dim Width As Double
    Dim BWidth As Double
    Dim HLength As Double
    Dim BLength As Double
    Dim InsulationThickness As Double

    'assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    BWidth = arrayOfInputs(3)
    HLength = arrayOfInputs(4)
    BLength = arrayOfInputs(5)
    InsulationThickness = arrayOfInputs(6)

    ' Insert your code for output 4(Header insulation)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set -HLength / 2, 0, 0
    enPoint.Set HLength / 2, 0, 0

    Dim ObjHeaderIns As Object

    Set ObjHeaderIns = PlaceCylinder(m_outputColl, stPoint, enPoint, _
                                     Width + 2 * InsulationThickness, True)

    m_outputColl.AddOutput "HeaderIns", ObjHeaderIns
    Set ObjHeaderIns = Nothing

    ' Insert your code for output 5(Branch insulation)
    stPoint.Set 0, BLength, 0
    enPoint.Set 0, -BLength, 0

    Dim ObjBranchIns As Object

    Set ObjBranchIns = PlaceCylinder(m_outputColl, stPoint, enPoint, _
                                     BWidth + 2 * InsulationThickness, True)

    m_outputColl.AddOutput "BranchIns", ObjBranchIns
    Set ObjBranchIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub

Private Function PlaceCylinder(ByVal objOutputColl As Object, _
                               lStartPoint As AutoMath.DPosition, _
                               lEndPoint As AutoMath.DPosition, _
                               lDiameter As Double, _
                               isCapped As Boolean) As Object
    Const METHOD = "PlaceCylinder"
    On Error GoTo ErrorHandler

    Dim circleCenter As AutoMath.DPosition
    Dim circleNormal As AutoMath.DVector
    Dim objCircle As IngrGeom3D.Circle3d
    Dim dblCylWidth As Double
    Dim objProjection As IngrGeom3D.Projection3d
    Dim geomFactory As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set circleCenter = New AutoMath.DPosition
    circleCenter.Set lStartPoint.X, lStartPoint.Y, lStartPoint.Z
    Set circleNormal = New AutoMath.DVector
    circleNormal.Set lEndPoint.X - lStartPoint.X, _
                     lEndPoint.Y - lStartPoint.Y, _
                     lEndPoint.Z - lStartPoint.Z
    dblCylWidth = circleNormal.Length
    circleNormal.Length = 1

    ' Construct a circle that will be used to project the disc
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                     circleCenter.X, circleCenter.Y, circleCenter.Z, _
                                                                     circleNormal.X, circleNormal.Y, circleNormal.Z, _
                                                                     lDiameter / 2)

    ' Project the disc of body
    Set objProjection = geomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                                objCircle, _
                                                                circleNormal.X, circleNormal.Y, circleNormal.Z, _
                                                                dblCylWidth, isCapped)

    Set objCircle = Nothing

    Set PlaceCylinder = objProjection
    Set objProjection = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext

End Function

Public Function CmpDblLessThanOrEqualTo(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
                                        Optional ByVal Tolerance As Double = LINEAR_TOLERANCE) As Boolean

'DblLessThanOrEqualTo = IIf(LeftVariable <= (RightVariable + Tolerance), true, False)
    If (LeftVariable <= (RightVariable + Tolerance)) Then
        CmpDblLessThanOrEqualTo = True

    Else
        CmpDblLessThanOrEqualTo = False
    End If

End Function


