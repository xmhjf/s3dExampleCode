VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HRoundBranAngRTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'File
'    HRoundBranAngRTE.cls
'
'ProgId:    SP3DHRoundBranchAngRTE.HRoundBranAngRTE
'
'Author:    svsmylav
'
'Creation Date:  Friday, Jan 10 2003
'
'Description
'    Definition of HVAC Round Tee and Reducing Tee having branch angle as input.
'    This symbol is based on GSCADHRoundRTE delivered symbol for 90 degrees reducing Tee.
'
'Notes
'
'    <notes>
'
'History:
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   15.DEC.2004      MS  CR 67971 Corrected port indices and flow-directions in HVAC Symbols
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   10.12.2007     PS      CR-132021  Modified the default value of Angle
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHRoundBranchAngRTE.HRoundBranAngRTE"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHRoundBranchAngRTE.HRoundBranAngRTE"
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
  
    On Error GoTo ErrorHandler
  
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    ' Create a new input by new operator
    Dim Inputs(1 To 6) As IMSSymbolEntities.DInput
    Dim Index As Integer
  
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
    Inputs(2).Name = "BWidth"
    Inputs(2).Description = "Branch Width"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.08
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).Name = "HLength"
    Inputs(3).Description = "Length of the Tee along the Header"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.15
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).Name = "BLength"
    Inputs(4).Description = "Length of the Tee along the Branch"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.08
    Inputs(4).DefaultParameterValue = PC
    
    Set Inputs(5) = New IMSSymbolEntities.DInput
    Inputs(5).Name = "Angle"
    Inputs(5).Description = "Angle of the Branch"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 2 * Atn(1)
    Inputs(5).DefaultParameterValue = PC
    
    Set Inputs(6) = New IMSSymbolEntities.DInput
    Inputs(6).Name = "InsulationThickness"
    Inputs(6).Description = "Insulation Thickness"
    Inputs(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(6).DefaultParameterValue = PC
       
    For Index = 1 To 6
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 5) As IMSSymbolEntities.DOutput
    
    For Index = 1 To 5
        Set O(Index) = New IMSSymbolEntities.DOutput
        O(Index).Properties = 0
    Next
   
    O(1).Name = "HvacNozzle1"
    O(1).Description = "HvacPort1 of Tee and Reducing Tee"
    
    O(2).Name = "HvacNozzle2"
    O(2).Description = "HvacPort2 of Tee and Reducing Tee"
    
    O(3).Name = "HvacNozzle3"
    O(3).Description = "HvacPort2 of Tee and Reducing Tee"
    
    O(4).Name = "HvacHeaderIns"
    O(4).Description = "Hvac Header Insulation"
    
    O(5).Name = "HvacBranchIns"
    O(5).Description = "Hvac Branch Insulation"
    
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
        
    For Index = 1 To 3
        oRepPhysicalOutputs.SetOutput O(Index)
        Set O(Index) = Nothing
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
    PhysicalRepEval.ProgId = "SP3DHRoundBranchAngRTE.HRoundBranAngRTE"

    'Set the evaluations for the Physical representation on the definition
    Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    
    Set PhysicalRepEval = Nothing
    Set RepEvalsIf = Nothing
    Set oRepPhysicalOutputs = Nothing
    
'   Define the representation "Insulation"
    rep1.Name = "Insulation"
    rep1.Description = "Insulation Represntation of the Round Tee with Angle input"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to Insulation. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = Insulation

    Dim oRepInsulationOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepInsulationOutputs = rep1
    
    For Index = 4 To 5
        oRepInsulationOutputs.SetOutput O(Index)
        Set O(Index) = Nothing
    Next Index
    
    'Set the representation to definition
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1

    'Define evaluation for Insulation representation
    Dim InsulationRepEval As DRepresentationEvaluation
    Set InsulationRepEval = New DRepresentationEvaluation
    InsulationRepEval.Name = "Insulation"
    InsulationRepEval.Description = "script for the Insulation representation"
    InsulationRepEval.Properties = igREPRESENTATION_HIDDEN
    InsulationRepEval.Type = igREPRESENTATION_VBFUNCTION
    InsulationRepEval.ProgId = "SP3DHRoundBranchAngRTE.HRoundBranAngRTE"
    
    'Set the evaluations for the Physical representation on the definition
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation InsulationRepEval
    
    Set InsulationRepEval = Nothing
    
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
    ElseIf StrComp(repName, "Insulation") = 0 Then
        InsulationRep arrayOfInputs
    End If
End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF HRoundBranAngRTE
'=========================================================================
Private Sub Physical(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
        
    Dim Width As Double
    Dim Depth As Double
    Dim BWidth As Double
    Dim BDepth As Double
    Dim HLength As Double
    Dim BLength As Double
'   Angle is measured from positive X-axis in anti-clockwise direction
    Dim Angle As Double
    Dim InsulationThickness As Double
    
    Dim halfHLength As Double
    Dim FlangeWidth As Double
    Dim NozzleLength As Double
    Dim Thickness As Double
    'Dim NozzleFlangeThickness As Double
    Dim CornerRadius As Double
        
'   Assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    BWidth = arrayOfInputs(3)
    HLength = arrayOfInputs(4)
    BLength = arrayOfInputs(5)
    Angle = arrayOfInputs(6)
    InsulationThickness = arrayOfInputs(7)
    Depth = Width
    BDepth = BWidth
'============================================================
'BUILD HVAC NOZZLE ON BASIC ASSEMBLY Tee and Reducing Tee
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
    If HLength < minHL Then
    HLength = minHL
    End If
    If BLength < minBL Then
    BLength = minBL
    End If
    halfHLength = HLength / 2
    
    'Set HVAC nozzle parameters
    iPortIndex = 1
    
    EndPrep = 301
    CornerRadius = 0#
    NozzleLength = HLength
    
'   To construct nozzle as crosssection only, use FlangeWidth of 0, and
'   some non-zero value for flange thickness
'   Flange thickness assumed to be negigible thickness
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
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
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
'BUILD HVACNOZZLE2 Of Tee and Reducing Tee
'==================================================
    iPortIndex = 2
    FlowDir = DistribFlow_BOTH
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    NozzleLength = 0.0002

    
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
'BUILD HVACNOZZLE3 Of Tee and Reducing Tee
'============================================
    iPortIndex = 3
    FlowDir = DistribFlow_BOTH
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                    Thickness, FlangeWidth, FlowDir, BWidth, _
                                    BDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                    "HNoz3", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    NozzleLength = BLength
    
    X = BLength * Cos(Angle)
    Y = BLength * Sin(Angle)
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set Cos(Angle), Sin(Angle), 0
    iDistribPort.SetDirectionVector dir
    
    'dir.Set 0, 1, 0
    'iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "HvacNozzle3", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing
    
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "HRoundBranAngRTE definition"
        Debug.Assert False
End Sub

'=========================================================================
'CREATION OF INSULATION REPRESENTATION OF HRoundBranAngRTE
'=========================================================================

Private Sub InsulationRep(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    
    Dim Width As Double
    Dim BWidth As Double
    Dim HLength As Double
    Dim BLength As Double
'   Angle is measured from positive X-axis in anti-clockwise direction
    Dim Angle As Double
    Dim InsulationThickness As Double
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
        
'   Assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    BWidth = arrayOfInputs(3)
    HLength = arrayOfInputs(4)
    BLength = arrayOfInputs(5)
    Angle = arrayOfInputs(6)
    InsulationThickness = arrayOfInputs(7)
    
'   Define default minimum Tee or Reducing Tee header and branch lengths
    Dim minHL As Double
    Dim minBL As Double
    
    minHL = Width + 0.05
    minBL = minHL / 2
    If HLength < minHL Then
        HLength = minHL
    End If
    If BLength < minBL Then
        BLength = minBL
    End If
        
' Code for Output 4(Hvac Header Insulation)
    Dim objHvacHeaderIns As Object
    
    stPoint.Set -HLength / 2, 0, 0
    enPoint.Set HLength / 2, 0, 0

    Set objHvacHeaderIns = PlaceCylinder(m_outputColl, stPoint, enPoint, _
                                                Width + 2 * InsulationThickness, True)
   
'   Add the output into the Collection
    m_outputColl.AddOutput "HvacHeaderIns", objHvacHeaderIns
    Set objHvacHeaderIns = Nothing
    
' Code for Output 5(Hvac Branch Insulation)
    Dim objHvacBranchIns As Object
    
    stPoint.Set 0, 0, 0
    enPoint.Set BLength * Cos(Angle), BLength * Sin(Angle), 0
    Set objHvacBranchIns = PlaceCylinder(m_outputColl, stPoint, enPoint, _
                                                BWidth + 2 * InsulationThickness, True)
    
'   Add the output into the Collection
    m_outputColl.AddOutput "HvacBranchIns", objHvacBranchIns
    Set objHvacBranchIns = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Insulation representation", vbExclamation, "HRoundBranAngRTE definition"
        Debug.Assert False
End Sub

Private Function PlaceCylinder(ByVal objOutputColl As Object, _
                                lStartPoint As AutoMath.DPosition, _
                                lEndPoint As AutoMath.DPosition, _
                                lDiameter As Double, _
                                isCapped As Boolean) As Object

    On Error GoTo ErrorHandler
    
    Dim circleCenter    As AutoMath.DPosition
    Dim circleNormal    As AutoMath.DVector
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim dblCylWidth     As Double
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory

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
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
        
End Function


