VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "H45DegRndLateral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   H45DegRndLateral.cls
'   SP3DH45DegRndLateral.H45DegRndLateral
'   Author:         kkk
'   Creation Date:  Friday 10, March 2006
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       HVAC 45 Degree Round Lateral
'       The Symbol is taken from Sprial Catalog.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   26.04.2006      kkc     DI-97251 Added BLength Calculation to the InsulationRep and
'                           BLength is removed from input parameters
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
  IJDUserSymbolServices_GetDefinitionName = "SP3DH45DegRndLateral.H45DegRndLateral"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
  On Error GoTo ErrorHandler
  Dim oSymbolFactory As New imssymbolentities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As imssymbolentities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DH45DegRndLateral.H45DegRndLateral"
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

    ' Feed SP3DH45DegRndLateral.H45DegRndLateral
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
    Dim Inputs(1 To 4) As imssymbolentities.DInput
    Dim Index As Integer
  
    ' Create a defaultValue
    Dim PC As imssymbolentities.DParameterContent
    Set PC = New imssymbolentities.DParameterContent 'not persistent PC
  
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
    Inputs(3).Description = "Length of the Lateral along the Header"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.15
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New imssymbolentities.DInput
    Inputs(4).Name = "InsulationThickness"
    Inputs(4).Description = "Insulation Thickness"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(4).DefaultParameterValue = PC
    
    
    For Index = 1 To 4
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 11) As imssymbolentities.DOutput
    
    For Index = 1 To 11
        Set O(Index) = New imssymbolentities.DOutput
        O(Index).Properties = 0
    Next
   
    O(1).Name = "HvacNozzle1"
    O(1).Description = "HvacPort1 of Lateral"
    
    O(2).Name = "HvacNozzle2"
    O(2).Description = "HvacPort2 of Lateral"
    
    O(3).Name = "HvacNozzle3"
    O(3).Description = "HvacPort2 of Lateral"
    
    O(4).Name = "ObjBody1"
    O(4).Description = "Body 1 of Lateral"
    
    O(5).Name = "ObjBody2"
    O(5).Description = "Body 2 of Lateral"
    
    O(6).Name = "ObjBody3"
    O(6).Description = "Body 3 of Lateral"
    
    O(7).Name = "ObjTakeOff1"
    O(7).Description = "TakeOff 1 of Lateral"
    
    O(8).Name = "ObjTakeOff2"
    O(8).Description = "TakeOff 2 of Lateral"
    
    O(9).Name = "ObjTakeOff3"
    O(9).Description = "TakeOff 3 of Lateral"
    
            
    O(10).Name = "HeaderIns"
    O(10).Description = "Header Insulation"
    
    O(11).Name = "BranchIns"
    O(11).Description = "Branch Insulation"

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
        
    For Index = 1 To 9
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
    PhysicalRepEval.ProgId = "SP3DH45DegRndLateral.H45DegRndLateral"

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
    For Index = 10 To 11
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
    InsulationRepEval.ProgId = "SP3DH45DegRndLateral.H45DegRndLateral"
    
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
'CREATION OF PHYSICAL REPRESENTATION OF Lateral
'=========================================================================


Private Sub Physical(ByRef arrayOfInputs())
    Const METHOD = "Physical"
    On Error GoTo ErrorLabel
    
    
    Dim Width As Double, Depth As Double
    Dim BWidth As Double, BDepth As Double
    Dim HLength As Double, halfHLength As Double
    Dim BLength As Double, FlangeWidth As Double
    Dim NozzleLength As Double, Thickness As Double
    'Dim NozzleFlangeThickness As Double
    Dim CornerRadius As Double
        
    'assign to meaningful variables from the input array
    
    Dim PI As Double
    PI = 4 * Atn(1)
    
    'assign to meaningful variables from the input array
    
    Width = arrayOfInputs(2)
    BWidth = arrayOfInputs(3)
    HLength = arrayOfInputs(4)
' Branch Length is computed as given in the Spiral Catalog.
' Branch Length is computed as BL = BWidth*Cot(theta)+ 2 Inches

    BLength = (BWidth / Tan(PI / 4)) + 2 * 25.4 / 1000
    
    Depth = Width
    BDepth = BWidth
'============================================================
'BUILD HVAC NOZZLE ON BASIC ASSEMBLY Lateral
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
    TakeOff = 2 * 25.4 / 1000 ' 2 Inches
    
    ' Define default minimum Lateral or Reducing Lateral header and branch lengths
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
    'NozzleFlangeThickness = Thickness + HDamperLength
    NozzleLength = HLength
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
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
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
'============================================
'BUILD Body 1 Of Lateral
'============================================

Dim ObjBody1 As Object
Dim stPoint   As New AutoMath.DPosition
Dim enPoint   As New AutoMath.DPosition
    
    
    stPoint.Set -HLength / 2 - 4 / 1000, 0, 0
    enPoint.Set -HLength / 2 + 4 / 1000, 0, 0
    
    Set ObjBody1 = PlaceCylinder(m_outputColl, stPoint, enPoint, Width * 1.06, True)
    
    
'   Set the output
    m_outputColl.AddOutput "ObjBody1", ObjBody1
    
    Set ObjBody1 = Nothing
'============================================
'BUILD TakeOff 1 Of Lateral
'============================================
    Dim ObjTakeOff1 As Object
    
    stPoint.Set -HLength / 2 - TakeOff, 0, 0
    enPoint.Set -HLength / 2, 0, 0
    
    Set ObjTakeOff1 = PlaceCylinder(m_outputColl, stPoint, enPoint, Width * 1.01, False)
'   Set the output
    m_outputColl.AddOutput "ObjTakeOff1", ObjTakeOff1
    
    Set ObjTakeOff1 = Nothing
    
'============================================
'BUILD Body 2 Of Lateral
'============================================
Dim ObjBody2 As Object
   
    stPoint.Set HLength / 2 - 4 / 1000, 0, 0
    enPoint.Set HLength / 2 + 4 / 1000, 0, 0
    Set ObjBody2 = PlaceCylinder(m_outputColl, stPoint, enPoint, Width * 1.06, True)
    
'   Set the output
    m_outputColl.AddOutput "ObjBody2", ObjBody2
    Set ObjBody2 = Nothing
    
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
'BUILD Body 3 Of Lateral
'============================================
Dim ObjBody3 As Object
    
    X = -halfHLength + (25.4 / 1000) + BWidth / 2 * Sin(PI / 4) + (BLength * Cos(PI / 4))
    Y = -BLength * Sin(PI / 4) - Depth / 2 + (BWidth / 2) * Cos(PI / 4)
    Z = 0#
    
    stPoint.Set X - (4 / 1000) * Cos(PI / 4), Y + (4 / 1000) * Sin(PI / 4), Z
    enPoint.Set X + (4 / 1000) * Cos(PI / 4), Y - (4 / 1000) * Sin(PI / 4), Z
    Set ObjBody3 = PlaceCylinder(m_outputColl, stPoint, enPoint, BWidth * 1.06, True)
    
'   Set the output
    m_outputColl.AddOutput "ObjBody3", ObjBody3
    Set ObjBody3 = Nothing
    
'============================================
'BUILD TakeOff 3 Of Lateral
'============================================
    Dim ObjTakeOff3 As Object
    
    X = -halfHLength + (25.4 / 1000) + BWidth / 2 * Sin(PI / 4) + (BLength * Cos(PI / 4))
    Y = -BLength * Sin(PI / 4) - Depth / 2 + (BWidth / 2) * Cos(PI / 4)
    Z = 0#
    
    stPoint.Set X, Y, Z
    enPoint.Set X + TakeOff * Cos(PI / 4), Y - TakeOff * Sin(PI / 4), Z
    Set ObjTakeOff3 = PlaceCylinder(m_outputColl, stPoint, enPoint, BWidth * 1.01, False)
'   Set the output
    m_outputColl.AddOutput "ObjTakeOff3", ObjTakeOff3
    
    Set ObjTakeOff3 = Nothing
    
'==================================================
'BUILD HVACNOZZLE2 Of Lateral
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
'BUILD HVACNOZZLE3 Of Lateral
'============================================
    iPortIndex = 3
    FlowDir = DistribFlow_BOTH
         Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, BWidth, _
                                            BDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz3", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
        NozzleLength = BLength

    
    X = -halfHLength + (25.4 / 1000) + BWidth / 2 * Sin(PI / 4) + (BLength * Cos(PI / 4))
    Y = -BLength * Sin(PI / 4) - Depth / 2 + (BWidth / 2) * Cos(PI / 4)
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

    Dim Width As Double
    Dim BWidth As Double
    Dim HLength As Double
    Dim BLength As Double
    Dim InsulationThickness As Double
    Dim TakeOff As Double
    Dim PI As Double
    PI = 4 * Atn(1)

    TakeOff = 2 * 25.4 / 1000 ' 2 Inches

    'assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    BWidth = arrayOfInputs(3)
    HLength = arrayOfInputs(4)
    'BLength = arrayOfInputs(5)
    InsulationThickness = arrayOfInputs(5)
    
    BLength = (BWidth / Tan(PI / 4)) + 2 * 25.4 / 1000

 ' Insert your code for output 4(Header insulation)
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    Dim minHL As Double
    Dim minBL As Double

    ' Define default minimum Lateral or Reducing Lateral header and branch lengths
    minHL = Width + 0.05
    minBL = minHL / 2
    If HLength < minHL Then
        HLength = minHL
    End If
    If BLength < minBL Then
        BLength = minBL
    End If

    stPoint.Set -HLength / 2 - TakeOff, 0, 0
    enPoint.Set HLength / 2 + TakeOff, 0, 0

    Dim ObjHeaderIns As Object

    Set ObjHeaderIns = PlaceCylinder(m_outputColl, stPoint, enPoint, _
                 Width * 1.01 + 2 * InsulationThickness, True)

    m_outputColl.AddOutput "HeaderIns", ObjHeaderIns
    Set ObjHeaderIns = Nothing

 ' Insert your code for output 5(Branch insulation)
 
 Dim X As Double, Y As Double, Z As Double
 
    X = -HLength / 2 + (25.4 / 1000) + BWidth / 2 * Sin(PI / 4) + ((BLength + TakeOff) * Cos(PI / 4))
    Y = -(BLength + TakeOff) * Sin(PI / 4) - Width / 2 + (BWidth / 2) * Cos(PI / 4)
    Z = 0#
    
    stPoint.Set X, Y, Z
    enPoint.Set X - BLength * Cos(PI / 4), Y + BLength * Sin(PI / 4), Z

    Dim ObjBranchIns As Object

    Set ObjBranchIns = PlaceCylinder(m_outputColl, stPoint, enPoint, _
                  BWidth * 1.01 + 2 * InsulationThickness, True)
                
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
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
        
End Function

