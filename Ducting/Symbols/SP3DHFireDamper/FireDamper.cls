VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FireDamper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'File
'   FireDamper.cls
'
'Author
'    NN 10 Mars 2002
'
'Description
'    Definition of HVAC AirFilterR.
'
'Notes
'
'    <notes>
'
'History:
'   dd.mmm.yyyy     who                 change description
'   -----------     ---                 ------------------
'   03.Nov.2005     kkk     CR-87366  Create TROX HVAC parts and symbols
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  06.DEC.2007     MA   CR-123921 Implemented Part Data Basis for the value 116.
'  08.SEP.2008     VKC  CR-149512 Implemented Part Data Basis for the value 151,152,153,154,155.
'  12.SEP.2008     VKC  CR-149512 Also added the Insulation Aspect for the Part Data Basis.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Const NEGLIGIBLE_THICKNESS = 0.0001

Private Const nOutputs = 50

Private m_GeomFactory As IngrGeom3D.GeometryFactory

'Private Form As frmEditOccurrence

Implements IJDUserSymbolServices

Private Sub Class_Initialize()

Set m_GeomFactory = New IngrGeom3D.GeometryFactory

End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHFireDamper.FireDamper"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHFireDamper.FireDamper"
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

    ' Feed SP3DHFlatFlange.HFF Definition
    ' Inputs:
    '          1. "Width"
    '          2. "Depth"
    '          3. "FilterLength"
    '          4. "Thickness"
    '          5. "FlangeThickness"
    '          6.  "HVACShape"
    '          7.  "TopOffset"
    '          8.  "BottomOffset"
    '          9.  "MotorWidth"
    '         10.  "DamperWidth"
    '         11.  "DamperDepth"
    '         12.  "TransitionLength"
    '         13.  "OffsetDepth"
    '         14.  "PlateSpacing"
    '         15.  "Length1"
    '         16.  "Length"
    '         17.  "PlateWidth"
    '         18.  "PlateDepth"
    
    '
    ' The representation "Physical"  and "Insulation" is set to the definition
    ' Physical :    (RepresentationId = 1)
    ' Insulation :    (RepresentationId = 32)
      
    On Error GoTo ErrorHandler
  
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
    
    ' Create a new input by new operator
    Dim Inputs(1 To 24) As IMSSymbolEntities.DInput
    Dim index As Integer
  
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition
    
    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).Name = "Width"
    Inputs(2).Description = "Width of the Rect Duct"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.4
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).Name = "Depth"
    Inputs(3).Description = "Depth of the Rect Duct"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(3).DefaultParameterValue = PC
    
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).Name = "Length"
    Inputs(4).Description = "FilterLength between ducts"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(4).DefaultParameterValue = PC
    
    Set Inputs(5) = New IMSSymbolEntities.DInput
    Inputs(5).Name = "Thickness"
    Inputs(5).Description = "Thickness of the Flat Flange of AirFilterR"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(5).DefaultParameterValue = PC
    
    Set Inputs(6) = New IMSSymbolEntities.DInput
    Inputs(6).Name = "FlangeWidth"
    Inputs(6).Description = "Width of the Flat Flange of AirFilterR"
    Inputs(6).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(6).DefaultParameterValue = PC
    
    ' GW. Add inputs
    
    Set Inputs(7) = New IMSSymbolEntities.DInput
    Inputs(7).Name = "CasingWidth"
    Inputs(7).Description = "CasingWidth"
    Inputs(7).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(7).DefaultParameterValue = PC
    
    Set Inputs(8) = New IMSSymbolEntities.DInput
    Inputs(8).Name = "CasingThickness"
    Inputs(8).Description = "CasingThickness"
    Inputs(8).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(8).DefaultParameterValue = PC
    
    Set Inputs(9) = New IMSSymbolEntities.DInput
    Inputs(9).Name = "DamperLength"
    Inputs(9).Description = "DamperLength"
    Inputs(9).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(9).DefaultParameterValue = PC
    
    Set Inputs(10) = New IMSSymbolEntities.DInput
    Inputs(10).Name = "ActuatorWidth"
    Inputs(10).Description = "Actuator Width"
    Inputs(10).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(10).DefaultParameterValue = PC
    
    Set Inputs(11) = New IMSSymbolEntities.DInput
    Inputs(11).Name = "TopOffset"
    Inputs(11).Description = "Offset of the Motor from the Damper top"
    Inputs(11).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(11).DefaultParameterValue = PC
    
    Set Inputs(12) = New IMSSymbolEntities.DInput
    Inputs(12).Name = "BottomOffset"
    Inputs(12).Description = "Offset of the Motor from the Damper bottom"
    Inputs(12).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(12).DefaultParameterValue = PC
    
    Set Inputs(13) = New IMSSymbolEntities.DInput
    Inputs(13).Name = "MotorWidth"
    Inputs(13).Description = "Motor Width"
    Inputs(13).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(13).DefaultParameterValue = PC
     
    Set Inputs(14) = New IMSSymbolEntities.DInput
    Inputs(14).Name = "DamperWidth"
    Inputs(14).Description = "Damper Width"
    Inputs(14).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(14).DefaultParameterValue = PC
    
    Set Inputs(15) = New IMSSymbolEntities.DInput
    Inputs(15).Name = "DamperDepth"
    Inputs(15).Description = "Damper Height"
    Inputs(15).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(15).DefaultParameterValue = PC
    
    Set Inputs(16) = New IMSSymbolEntities.DInput
    Inputs(16).Name = "HVACShape"
    Inputs(16).Description = "HVAC Shape"
    Inputs(16).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 1
    Inputs(16).DefaultParameterValue = PC
    
    Set Inputs(17) = New IMSSymbolEntities.DInput
    Inputs(17).Name = "PlateSpacing"
    Inputs(17).Description = "Distance Between The Plates"
    Inputs(17).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(17).DefaultParameterValue = PC
       
    Set Inputs(18) = New IMSSymbolEntities.DInput
    Inputs(18).Name = "Length1"
    Inputs(18).Description = "Distance From One End"
    Inputs(18).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(18).DefaultParameterValue = PC
    
    Set Inputs(19) = New IMSSymbolEntities.DInput
    Inputs(19).Name = "Length"
    Inputs(19).Description = "Length Of The Cylinder"
    Inputs(19).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(19).DefaultParameterValue = PC
    
    Set Inputs(20) = New IMSSymbolEntities.DInput
    Inputs(20).Name = "PlateWidth"
    Inputs(20).Description = "Plate Width"
    Inputs(20).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(20).DefaultParameterValue = PC
    
    Set Inputs(21) = New IMSSymbolEntities.DInput
    Inputs(21).Name = "PlateDepth"
    Inputs(21).Description = "Plate Depth"
    Inputs(21).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(21).DefaultParameterValue = PC
    
    Set Inputs(22) = New IMSSymbolEntities.DInput
    Inputs(22).Name = "InsulationThickness"
    Inputs(22).Description = "Insulation Thickness"
    Inputs(22).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(22).DefaultParameterValue = PC
    
    Set Inputs(23) = New IMSSymbolEntities.DInput
    Inputs(23).Name = "OffsetDepth"
    Inputs(23).Description = "Offset Depth"
    Inputs(23).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(23).DefaultParameterValue = PC
    
    Set Inputs(24) = New IMSSymbolEntities.DInput
    Inputs(24).Name = "TransitionLength"
    Inputs(24).Description = "Transition Length"
    Inputs(24).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(24).DefaultParameterValue = PC
    
    For index = 2 To 24
        InputsIf.SetInput Inputs(index), index
        Set Inputs(index) = Nothing
    Next

    'Define the outputs
    Dim o(1 To nOutputs) As IMSSymbolEntities.DOutput
    
    For index = 1 To nOutputs
        Set o(index) = New IMSSymbolEntities.DOutput
        o(index).Properties = 0
    Next
   
    o(1).Name = "HvacNozzle1"
    o(1).Description = "HvacPort1 of Fire Damper"
    
    o(2).Name = "HvacNozzle2"
    o(2).Description = "HvacPort2 of Fire Damper"
    
    o(3).Name = "Hebel"
    o(3).Description = "Hebel"
    
    For index = 4 To nOutputs
      o(index).Name = "Feature" & (index - 3)
      o(index).Description = "Feature" & (index - 3)
    Next index
    
    
' Description:  This defines 1 representation which is Physical.
'               This defines 32 representation which is Insulation.
'********************************************************************
   'Set the representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Dim rep32 As IMSSymbolEntities.DRepresentation
    
    On Error GoTo ErrorHandler
    Set rep1 = New IMSSymbolEntities.DRepresentation
    Set rep32 = New IMSSymbolEntities.DRepresentation
    
    'Define the representation "Physical"
    rep1.Name = "Physical"
    rep1.Description = "Physical Represntation of the Air Distrib Assembly"
    rep1.Properties = igCOLLECTION_VARIABLE
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical
    
    Dim oRepPhysicalOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepPhysicalOutputs = rep1
        
    'Set the outputs
    For index = 1 To 3
        oRepPhysicalOutputs.SetOutput o(index)
        Set o(index) = Nothing
    Next
    
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1
    
    'Redefine outputs
    oRepPhysicalOutputs.RemoveAllOutput
    Set oRepPhysicalOutputs = Nothing
    
    'Define the representation "Insulation"
       
    rep32.Name = "Insulation"
    rep32.Description = "Insulation of the HVAC Fire Damper"
    'Set the repID to Insulation. See GSCADSymbolServices library to see
    'different repIDs available.
    rep32.RepresentationId = Insulation
    rep32.Properties = igCOLLECTION_VARIABLE
    Dim oRepInsulationOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepInsulationOutputs = rep32
        
    'Set the outputs
    oRepInsulationOutputs.SetOutput o(4)
    Set o(4) = Nothing
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep32

    Set rep32 = Nothing
    Set oRepInsulationOutputs = Nothing
    Set RepsIf = Nothing
    
    ' Description: This defines representaion evaluation.
'********************************************************************
    'Define evaluation for Physical representation
    Dim PhysicalRepEval As DRepresentationEvaluation
    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "script for the Physical representation"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "SP3DHFireDamper.FireDamper"

    
    'Set the evaluations for the Physical representation on the definition
    Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    
       
    'Define evaluation for Insulation
    Dim InsulationRepEval As DRepresentationEvaluation
    Set InsulationRepEval = New DRepresentationEvaluation
    InsulationRepEval.Name = "Insulation"
    InsulationRepEval.Description = "Script for the Insulation"
    InsulationRepEval.Properties = igREPRESENTATION_HIDDEN
    InsulationRepEval.Type = igREPRESENTATION_VBFUNCTION
    InsulationRepEval.ProgId = "SP3DHFireDamper.FireDamper"
    
    'Set the evaluations for the Insulation on the definition
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation InsulationRepEval
    
    Set InsulationRepEval = Nothing
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
  Exit Sub
  Resume
End Sub


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler

    Set m_outputColl = outputcoll
    If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
    ElseIf StrComp(repName, "Insulation") = 0 Then
        InsulationOf arrayOfInputs
    End If
    
Exit Sub

ErrorHandler:
        Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF HFF
'=========================================================================


Private Sub Physical(ByRef arrayOfInputs())
    
    Const METHOD = "FireDamper-Physical"
    
    On Error GoTo ErrorLabel
    
    Dim Width                    As Double
    Dim Depth                    As Double
    Dim FilterLength             As Double
    Dim Thickness                As Double
    Dim FlangeWidth              As Double
    
    Dim CasingWidth              As Double
    Dim CasingThickness          As Double
    Dim DamperLength             As Double
    Dim ActuatorWidth            As Double
    
    Dim HVACShape                As Double
    Dim TopOffset                As Double
    Dim BottomOffset             As Double
    Dim MotorWidth               As Double
    Dim DamperWidth              As Double
    Dim DamperDepth              As Double
    Dim TransitionLength         As Double
    Dim OffsetDepth              As Double
    Dim PlateSpacing             As Double
    Dim Length1                  As Double
    Dim Length                   As Double
    Dim PlateWidth               As Double
    Dim PlateDepth               As Double
    Dim arcs                     As Collection
    Dim oLine                    As Object
    Dim oProjection              As Object
    Dim oArc                     As IngrGeom3D.Arc3d
    Dim oAxis                    As AutoMath.DVector
    Dim oCString                 As IngrGeom3D.ComplexString3d
    Dim iCount                   As Integer
    Dim oGeomFactory             As IngrGeom3D.GeometryFactory
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oAxis = New AutoMath.DVector
    Set arcs = New Collection
    
       
    Dim HFilterLength            As Double
    Dim NozzleLength             As Double
    Dim NozzleFlangeThickness    As Double
    Dim CornerRadius             As Double
    Dim lHvacShape               As Long
    
    Dim NozzleFactory            As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle              As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle                  As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort             As GSCADNozzleEntities.IJDistribPort
    Dim PortStatus               As DistribPortStatus
    Dim DimBaseOuter             As Boolean
    Dim iPortIndex               As Integer
    Dim lngIndex                 As Long
    Dim obox                     As Object
    Dim ObjCylinder              As Object
    Dim oStPoint                 As AutoMath.DPosition
    Dim oEnPoint                 As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    
    'assign to meaningful variables from the input array
'    Width = arrayOfInputs(2)
'    Depth = arrayOfInputs(3)
'    FilterLength = arrayOfInputs(4)
'    Thickness = arrayOfInputs(5)
'    FlangeWidth = arrayOfInputs(6)
'    CasingWidth = arrayOfInputs(7)
'    CasingThickness = arrayOfInputs(8)
'    DamperLength = arrayOfInputs(9)
'    ActuatorWidth = arrayOfInputs(10)
'    TopOffset = arrayOfInputs(11)
'    BottomOffset = arrayOfInputs(12)
'    MotorWidth= arrayOfInputs(13)
'    DamperWidth = arrayOfInputs(14)
'    DamperDepth= arrayOfInputs(15)
'    HVACSShape=arrayOfInputs(16)
'    PlateSpacing = arrayOfInputs(17)
'    Length1 = arrayOfInputs(18)
'    Length = arrayOfInputs(19)
'    PlateWidth = arrayOfInputs(20)
'    PlateDepth = arrayOfInputs(21)
'    Insulationthickness=arrayOfInputs(22)
'    OffsetDepth=arrayOfInputs(23)
'    TransitionLength=arrayOfInputs(24)

    ' This symbol is based on the following Part data basis values that govern its geometry
    '   (Part data Basis value -115): Fire Damper
    '   (Part data Basis value -116): Fire and Gas Damper
    '   (Part data Basis value -151): Fire Damper Without a Casing.
    '   (Part data Basis value -152):Fire Damper With Retaining Plates.
    '   (Part data Basis value -153):Fire Damper With an actuator.
    '   (Part data Basis value -154):Fire Damper with Transition And An Actuator.
    '   (Part data Basis value -155):Fire Damper With Transition.
    
    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPart
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing

    Select Case lPartDataBasis

        Case Is <= 1, 115
        
            Width = arrayOfInputs(2)
            Depth = arrayOfInputs(3)
            FilterLength = arrayOfInputs(4)
            Thickness = arrayOfInputs(5)
            FlangeWidth = arrayOfInputs(6)
            CasingWidth = arrayOfInputs(7)
            CasingThickness = arrayOfInputs(8)
            DamperLength = arrayOfInputs(9)
              
        Case 116
          
            Width = arrayOfInputs(2)
            Depth = arrayOfInputs(3)
            FlangeWidth = arrayOfInputs(6)
            DamperLength = arrayOfInputs(9)
            ActuatorWidth = arrayOfInputs(10)
            lHvacShape = 1
        Case 151
            
            Width = arrayOfInputs(2)
            Depth = arrayOfInputs(3)
            DamperLength = arrayOfInputs(9)
            HVACShape = arrayOfInputs(16)
            lHvacShape = HVACShape
        
        Case 152
        
            Width = arrayOfInputs(2)
            Thickness = arrayOfInputs(5)
            PlateSpacing = arrayOfInputs(17)
            Length1 = arrayOfInputs(18)
            Length = arrayOfInputs(19)
            PlateWidth = arrayOfInputs(20)
            PlateDepth = arrayOfInputs(21)
            HVACShape = arrayOfInputs(16)
            lHvacShape = HVACShape
            'As it is a round duct this assumption is taken.
            If CmpDblEqual(Depth, 0) Then
            Depth = Width
            End If
            
        Case 153
            
            Width = arrayOfInputs(2)
            Depth = arrayOfInputs(3)
            DamperLength = arrayOfInputs(9)
            TopOffset = arrayOfInputs(11)
            BottomOffset = arrayOfInputs(12)
            MotorWidth = arrayOfInputs(13)
            HVACShape = arrayOfInputs(16)
            lHvacShape = HVACShape
            
        Case 154
            
            Width = arrayOfInputs(2)
            Depth = arrayOfInputs(3)
            DamperLength = arrayOfInputs(9)
            TopOffset = arrayOfInputs(11)
            BottomOffset = arrayOfInputs(12)
            MotorWidth = arrayOfInputs(13)
            DamperWidth = arrayOfInputs(14)
            DamperDepth = arrayOfInputs(15)
            HVACShape = arrayOfInputs(16)
            OffsetDepth = arrayOfInputs(23)
            TransitionLength = arrayOfInputs(24)
            lHvacShape = HVACShape
            
        Case 155
            
            Width = arrayOfInputs(2)
            Depth = arrayOfInputs(3)
            DamperLength = arrayOfInputs(9)
            DamperWidth = arrayOfInputs(14)
            DamperDepth = arrayOfInputs(15)
            HVACShape = arrayOfInputs(16)
            OffsetDepth = arrayOfInputs(23)
            TransitionLength = arrayOfInputs(24)
            lHvacShape = HVACShape
            
    End Select
'------------------------------------------------------------------------------------------------

    If (lPartDataBasis <= 1 Or lPartDataBasis = 115) Then
            '=====================================
            'BUILD HVAC NOZZLE ON BASIC ASSEMBLY
            '=====================================
            Dim nozzleIndex As Integer
            Dim iLogicalDistPort As GSCADNozzleEntities.IJLogicalDistPort
            
            Dim EndPrep             As Long
            Dim FlowDir             As DistribFlow
            Dim PortDepth           As Double
            Dim CptOffset           As Double
            Dim pos                 As New AutoMath.DPosition
            Dim dir                 As New AutoMath.DVector
            
            'Set HVAC nozzle parameters
            iPortIndex = 1
            PortDepth = 0#
            CptOffset = 0
            
            EndPrep = 11
            CornerRadius = 0#
            HFilterLength = FilterLength / 2
            'NozzleFlangeThickness = Thickness + HFilterLength
            NozzleLength = FilterLength - Thickness
            'NozzleLength Has to be GREATER than NozzleFlangeThickness
         
            
        
            ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
            ' some non-zero value for flange thickness
            ' Flange thickness assumed to be negigible thickness
            
            
            FlowDir = DistribFlow_IN
        
            ' Depth and Width of crosssection will be the same as diameter of the nozzle which
            ' is the inletDia
            
            DimBaseOuter = True
            PortStatus = DistribPortStatus_BASE
        
            
            Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                                    Thickness, FlangeWidth, FlowDir, Width, _
                                                    Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                                    "hvac1", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
           
            'Position of the nozzle should be the conenct point of the nozzle
            x = -DamperLength * 0.5 - CptOffset '  -HFilterLength - CptOffset
            y = 0#
            z = 0#
            pos.Set x, y, z
            Set iDistribPort = oHvacNozzle
            iDistribPort.SetPortLocation pos
        
            'Direction specified here of the nozzle should be the direction in which pipe will be routed.
            'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
            dir.Set -1, 0, 0
            iDistribPort.SetDirectionVector dir
            
            dir.Set 0, 1, 0
            iDistribPort.SetRadialOrient dir
            
            Set iNozzle = oHvacNozzle
            iNozzle.Length = 0.26
        
            m_outputColl.AddOutput "HvacNozzle1", oHvacNozzle
            Set oHvacNozzle = Nothing
            Set iNozzle = Nothing
            Set iDistribPort = Nothing
        
            '=====================================
            'BUILD HVACNOZZLE2 Of AirFilterR
            '=====================================
            iPortIndex = 2
            Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                                    Thickness, FlangeWidth, FlowDir, Width, _
                                                    Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                                    "hvac2", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
           
            'Position of the nozzle should be the conenct point of the nozzle
            x = DamperLength * 0.5 + CptOffset ' HFilterLength + CptOffset
            y = 0#
            z = 0#
            pos.Set x, y, z
            Set iDistribPort = oHvacNozzle
            iDistribPort.SetPortLocation pos
        
            'Direction specified here of the nozzle should be the direction in which pipe will be routed.
            'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
            dir.Set 1, 0, 0
            iDistribPort.SetDirectionVector dir
            
            dir.Set 0, 1, 0
            iDistribPort.SetRadialOrient dir
         
            Set iNozzle = oHvacNozzle
            iNozzle.Length = DamperLength - 0.26
        
            m_outputColl.AddOutput "HvacNozzle2", oHvacNozzle
            Set oHvacNozzle = Nothing
            Set iNozzle = Nothing
            Set iDistribPort = Nothing
Else
'==================================================
        'BUILD HVACNOZZLE 1
 '==================================================
        Dim oPos As New AutoMath.DPosition
        Dim oDir As New AutoMath.DVector
        
        Dim oHvacPort As IJDHvacPort
        Dim oHvacColl As IJDCollection
        
        Dim lEndPrep(1 To 2)        As Long
        Dim dThickness(1 To 2)      As Double
        Dim dFlangeWidth(1 To 2)    As Double
        Dim lFlowDir(1 To 2)        As Long
        Dim dPortDepth(1 To 2)      As Double
        Dim dCptOffSet(1 To 2)      As Double
        Dim dNozzLength(1 To 2)     As Double
        
        CornerRadius = 0#
        dNozzLength(1) = 0.001
        dNozzLength(2) = 0.001
        
        iPortIndex = 1
        'Set HVAC nozzle parameters
        Set oHvacColl = oPart.GetNozzles()
        For iPortIndex = 1 To oHvacColl.Size
            Set oHvacPort = oHvacColl.Item(iPortIndex)
            lEndPrep(iPortIndex) = oHvacPort.EndPrep
            dThickness(iPortIndex) = oHvacPort.Thickness
            If CmpDblEqual(FlangeWidth, 0) Then
                dFlangeWidth(iPortIndex) = oHvacPort.FlangeWidth
            Else
                dFlangeWidth(iPortIndex) = FlangeWidth
            End If
            lFlowDir(iPortIndex) = oHvacPort.FlowDirection
            dPortDepth(iPortIndex) = oHvacPort.PortDepth
            dCptOffSet(iPortIndex) = oHvacPort.CptOffset
            'NozzleLength Has to be GREATER than NozzleFlangeThickness
            If CmpDblLessThanOrEqualTo(dThickness(iPortIndex), LINEAR_TOLERANCE) Then
                dThickness(iPortIndex) = NEGLIGIBLE_THICKNESS
            End If
            If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
                dNozzLength(iPortIndex) = dThickness(iPortIndex) + NEGLIGIBLE_THICKNESS
            End If
        Next iPortIndex
            
        Set oHvacPort = Nothing
        oHvacColl.Clear
        Set oHvacColl = Nothing
        
        iPortIndex = 1
        DimBaseOuter = True
        PortStatus = DistribPortStatus_BASE
        
        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                    lHvacShape, lEndPrep(1), _
                                    dThickness(1), dFlangeWidth(1), lFlowDir(1), Width, _
                                    Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                    "HNoz1", dPortDepth(1), dCptOffSet(1), False, _
                                    m_outputColl.ResourceManager)
    
        'Position of the Nozzle
        Select Case lPartDataBasis
        
            Case 116
             x = -DamperLength / 2
             y = 0#
             z = 0#
             oPos.Set x, y, z
             Set iDistribPort = oHvacNozzle
             iDistribPort.SetPortLocation oPos
                
        
            Case Is = 151, 153
                x = -DamperLength / 2
                y = 0#
                z = 0#
                oPos.Set x, y, z
                Set iDistribPort = oHvacNozzle
                iDistribPort.SetPortLocation oPos
            
            Case 152
                x = -Length / 2
                y = 0#
                z = 0#
                oPos.Set x, y, z
                Set iDistribPort = oHvacNozzle
                iDistribPort.SetPortLocation oPos
            
            Case Is = 154, 155
                x = -DamperLength / 2 - TransitionLength
                y = 0#
                z = 0#
                oPos.Set x, y, z
                Set iDistribPort = oHvacNozzle
                iDistribPort.SetPortLocation oPos
        End Select
        
        'Direction specified here of the nozzle should be the direction in which pipe will be routed.
        'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
        oDir.Set -1, 0, 0
        iDistribPort.SetDirectionVector oDir
        
        oDir.Set 0, 1, 0
        iDistribPort.SetRadialOrient oDir
        
        Set iNozzle = oHvacNozzle
        iNozzle.Length = dNozzLength(1)
 
        'Set the Output
        m_outputColl.AddOutput "HvacNozzle1", oHvacNozzle
        Set oHvacNozzle = Nothing
        Set iNozzle = Nothing
        Set iDistribPort = Nothing

        '==================================================
        'BUILD HVACNOZZLE 2
        '==================================================
        iPortIndex = 2
        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                    lHvacShape, lEndPrep(2), _
                                    dThickness(2), dFlangeWidth(2), lFlowDir(2), Width, _
                                    Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                    "HNoz2", dPortDepth(2), dCptOffSet(2), False, _
                                    m_outputColl.ResourceManager)
    
        'Position of the Nozzle
        Select Case lPartDataBasis
        
            Case 116
                x = DamperLength / 2
                y = 0#
                z = 0#
                oPos.Set x, y, z
                Set iDistribPort = oHvacNozzle
                iDistribPort.SetPortLocation oPos
        
            Case Is = 151, 153
                x = DamperLength / 2
                y = 0#
                z = 0#
                oPos.Set x, y, z
                Set iDistribPort = oHvacNozzle
                iDistribPort.SetPortLocation oPos
                
            Case 152
                x = Length / 2
                y = 0#
                z = 0#
                oPos.Set x, y, z
                Set iDistribPort = oHvacNozzle
                iDistribPort.SetPortLocation oPos
                
            Case Is = 154, 155
                x = DamperLength / 2 + TransitionLength
                y = 0#
                z = 0#
                oPos.Set x, y, z
                Set iDistribPort = oHvacNozzle
                iDistribPort.SetPortLocation oPos
            
        End Select
        
        'Direction specified here of the nozzle should be the direction in which pipe will be routed.
        'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
        oDir.Set 1, 0, 0
        iDistribPort.SetDirectionVector oDir
        
        oDir.Set 0, 1, 0
        iDistribPort.SetRadialOrient oDir
        
        Set iNozzle = oHvacNozzle
        iNozzle.Length = dNozzLength(2)

        'Set the Output
        m_outputColl.AddOutput "HvacNozzle2", oHvacNozzle
        Set oHvacNozzle = Nothing
        Set iNozzle = Nothing
        Set iDistribPort = Nothing

        Set NozzleFactory = Nothing
        Set oPos = Nothing
        Set oDir = Nothing
        
    End If
    
    Select Case lPartDataBasis
    
            Case Is <= 1, 115
            
            ' GW: Hebel:
            Dim objStart As New AutoMath.DPosition
            Dim objP1 As New AutoMath.DPosition
            Dim objP2 As New AutoMath.DPosition
            Dim objP3 As New AutoMath.DPosition
            Dim objP4 As New AutoMath.DPosition
            Dim objDir As New AutoMath.DVector
            Dim objUp As New AutoMath.DVector
            Dim objQuer As New AutoMath.DVector
        
            'GW: create center flansche
            Dim dirDown As New AutoMath.DVector
            
            objP1.Set -DamperLength * 0.5 + 0.26, Depth * 0.5 + CasingWidth, Width * 0.5 + CasingWidth
            objP2.Set -DamperLength * 0.5 + 0.26, -Depth * 0.5 - CasingWidth, Width * 0.5 + CasingWidth
            objP3.Set -DamperLength * 0.5 + 0.26, -Depth * 0.5 - CasingWidth, -Width * 0.5 - CasingWidth
            objP4.Set -DamperLength * 0.5 + 0.26, Depth * 0.5 + CasingWidth, -Width * 0.5 - CasingWidth
            
            dirDown.Set -CasingThickness, 0, 0
            
            Call createRect(m_outputColl, "Feature", lngIndex, objP1, objP2, objP3, objP4, CasingWidth, dirDown)
            
            ' GW: add Inspectionsöffnungen.
            Dim ObjBodyA As Object
            
            objP1.Set -DamperLength * 0.5 + 0.25 - 0.01, 0.1, Width * 0.5
            objP3.Set -DamperLength * 0.5 + FlangeWidth + 0.01, -0.06, Width * 0.5 + 0.03
            
            Set ObjBodyA = PlaceBox(m_outputColl, objP1, objP3)
            
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, ObjBodyA
            Set ObjBodyA = Nothing
            
            ' GW: add Hebel für Inspectionsöffnungen.
            
            objP1.Set -DamperLength * 0.5 + 0.25 - 0.01, 0.08, Width * 0.5 + 0.03
            objP3.Set -DamperLength * 0.5 + 0.25 - 0.01 - 0.16, 0.06, Width * 0.5 + 0.05
            
            Set ObjBodyA = PlaceBox(m_outputColl, objP1, objP3)
            
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, ObjBodyA
            Set ObjBodyA = Nothing
        
        
        Case 116
          
            Width = arrayOfInputs(2)
            Depth = arrayOfInputs(3)
            FlangeWidth = arrayOfInputs(6)
            DamperLength = arrayOfInputs(9)
            ActuatorWidth = arrayOfInputs(10)
            
                 
            'Create Rectangular Projection
            Dim dLineStrPoints(0 To 14) As Double
            Dim oLineString As IngrGeom3D.LineString3d
            
            dLineStrPoints(0) = -DamperLength / 2
            dLineStrPoints(1) = Depth / 2
            dLineStrPoints(2) = Width / 2
    
            dLineStrPoints(3) = -DamperLength / 2
            dLineStrPoints(4) = -Depth / 2
            dLineStrPoints(5) = Width / 2
    
            dLineStrPoints(6) = -DamperLength / 2
            dLineStrPoints(7) = -Depth / 2
            dLineStrPoints(8) = -Width / 2
    
            dLineStrPoints(9) = -DamperLength / 2
            dLineStrPoints(10) = Depth / 2
            dLineStrPoints(11) = -Width / 2
    
            dLineStrPoints(12) = -DamperLength / 2
            dLineStrPoints(13) = Depth / 2
            dLineStrPoints(14) = Width / 2
    
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
            Dim oAxisVec As AutoMath.DVector
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
    
            Dim oRectangle As Object
            Set oRectangle = PlaceProjection(m_outputColl, oLineString, oAxisVec, _
                                                            DamperLength, False)
    
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, oRectangle
    
            Set oRectangle = Nothing
            Set oLineString = Nothing
            
            
            oStPoint.Set DamperLength / 2, 0.3 * Depth, 0.5 * Width + FlangeWidth
            oEnPoint.Set -DamperLength / 2, -0.1 * Depth, 0.5 * Width + FlangeWidth + ActuatorWidth
          
            Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
          
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, obox
            Set obox = Nothing
      
            'Create Box 2
            oStPoint.Set 0.35 * DamperLength, 0.4 * Depth, _
                                        0.5 * Width + FlangeWidth + 0.1 * ActuatorWidth
            oEnPoint.Set -0.1 * DamperLength, 0.3 * Depth, _
                                        0.5 * Width + FlangeWidth + 0.9 * ActuatorWidth
         
            Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
          
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, obox
            Set obox = Nothing
      
            'Create Box 3
            oStPoint.Set 0.35 * DamperLength, -0.1 * Depth, _
                                        0.5 * Width + FlangeWidth + 0.1 * ActuatorWidth
            oEnPoint.Set -0.1 * DamperLength, -0.2 * Depth, _
                                        0.5 * Width + FlangeWidth + 0.9 * ActuatorWidth
    
            Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
          
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, obox
            Set obox = Nothing
            
            'Create Box 4
            oStPoint.Set DamperLength / 2, 0.3 * Depth, 0.5 * Width
            oEnPoint.Set -DamperLength / 2, -0.1 * Depth, _
                                        0.5 * Width + FlangeWidth
    
            Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
          
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, obox
            Set obox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

    Case 151
          
        'Create Box
        oStPoint.Set DamperLength / 2, 0.5 * Depth, 0.5 * Width
        oEnPoint.Set -DamperLength / 2, -0.5 * Depth, -0.5 * Width
        Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
      
        'Set the output
        lngIndex = lngIndex + 1
        m_outputColl.AddOutput "Feature" & lngIndex, obox
        
        Set obox = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        
    Case 152
         
        If HVACShape = 4 Then 'Round
        
            'Create Cylinder(For the cylinder)
            oStPoint.Set -Length / 2, 0, 0
            oEnPoint.Set Length / 2, 0, 0
            Set ObjCylinder = PlaceCylinder(m_outputColl, oStPoint, oEnPoint, Width, True)
    
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, ObjCylinder
            Set ObjCylinder = Nothing
            
            'Create Box (For the plate1)
            oStPoint.Set -Length / 2 + Length1, -PlateDepth / 2, -PlateWidth / 2
            oEnPoint.Set -Length / 2 + Length1 + Thickness, PlateDepth / 2, PlateWidth / 2
            Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
          
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, obox
            Set obox = Nothing
            
            'Create Box (For the plate2)
            oStPoint.Set -Length + Length1 + PlateSpacing, -PlateDepth / 2, -PlateWidth / 2
            oEnPoint.Set -Length + Length1 + PlateSpacing + Thickness, PlateDepth / 2, PlateWidth / 2
            Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
          
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, obox
            
            Set obox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
        End If
        
    Case 153
  
        If HVACShape = 1 Then ' Rectangular
            
            'Create Box (For the Damper)
            oStPoint.Set DamperLength / 2, 0.5 * Depth, 0.5 * Width
            oEnPoint.Set -DamperLength / 2, -0.5 * Depth, -0.5 * Width
            Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
          
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, obox
            Set obox = Nothing
           
            'Create Box(For the actuator)
            If CmpDblEqual(BottomOffset, 0) And CmpDblGreaterthan(TopOffset, LINEAR_TOLERANCE) Then
              oStPoint.Set 0.3 * DamperLength, 0.5 * Depth + TopOffset, 0.5 * Width
              oEnPoint.Set -0.3 * DamperLength, 0, 0.5 * Width + MotorWidth
              
            ElseIf CmpDblEqual(TopOffset, 0) And CmpDblGreaterthan(BottomOffset, LINEAR_TOLERANCE) Then
              oStPoint.Set 0.3 * DamperLength, 0, 0.5 * Width
              oEnPoint.Set -0.3 * DamperLength, -0.5 * Depth - BottomOffset, 0.5 * Width + MotorWidth
              
            ElseIf CmpDblGreaterthan(TopOffset, LINEAR_TOLERANCE) And CmpDblGreaterthan(BottomOffset, LINEAR_TOLERANCE) Then
              oStPoint.Set 0.3 * DamperLength, 0.5 * Depth + TopOffset, 0.5 * Width
              oEnPoint.Set -0.3 * DamperLength, -0.5 * Depth - BottomOffset, 0.5 * Width + MotorWidth
              
            Else
              oStPoint.Set 0.3 * DamperLength, 0.5 * Depth, 0.5 * Width
              oEnPoint.Set -0.3 * DamperLength, -0.5 * Depth, 0.5 * Width + MotorWidth
            End If
              
            Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
            
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, obox
            
            Set obox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
        End If
            
  Case 154
                
        If CmpDblEqual(OffsetDepth, 0) Then
            'Create Box (For the Damper)
            oStPoint.Set DamperLength / 2, 0.5 * DamperDepth, 0.5 * DamperWidth
            oEnPoint.Set -DamperLength / 2, -0.5 * DamperDepth, -0.5 * DamperWidth
        Else
            oStPoint.Set DamperLength / 2, 0.5 * DamperDepth - 0.5 * Depth - OffsetDepth, 0.5 * DamperWidth
            oEnPoint.Set -DamperLength / 2, -0.5 * Depth - OffsetDepth, -0.5 * DamperWidth
        End If
        
        Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
      
        'Set the output
        lngIndex = lngIndex + 1
        m_outputColl.AddOutput "Feature" & lngIndex, obox
        Set obox = Nothing
   
        If HVACShape = 4 Then 'Round
             
            'Create Cylinder(For the Transition)
            oStPoint.Set -DamperLength / 2 - TransitionLength, 0, 0
            oEnPoint.Set DamperLength / 2 + TransitionLength, 0, 0
            Set ObjCylinder = PlaceCylinder(m_outputColl, oStPoint, oEnPoint, Width, True)
            
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, ObjCylinder
            Set ObjCylinder = Nothing
            
        End If
            
        If HVACShape = 1 Then 'Rectangular
             
            'Create Rectangular box(For the Transition)
            oStPoint.Set -DamperLength / 2 - TransitionLength, -Depth / 2, -Width / 2
            oEnPoint.Set DamperLength / 2 + TransitionLength, Depth / 2, Width / 2
            Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
          
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, obox
            Set obox = Nothing
            
        End If
            
        If HVACShape = 3 Then 'Oval
            If (Width = Depth) Then
             Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -DamperLength / 2 - TransitionLength, Depth / 2, (Width - Depth) / 2, _
                            -DamperLength / 2 - TransitionLength, 0, Width / 2, _
                               -DamperLength / 2 - TransitionLength, -Depth / 2, (Width - Depth) / 2)
            arcs.Add oArc
            
            
            Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -DamperLength / 2 - TransitionLength, -Depth / 2, -(Width - Depth) / 2, _
                            -DamperLength / 2 - TransitionLength, 0, -Width / 2, _
                              -DamperLength / 2 - TransitionLength, Depth / 2, -(Width - Depth) / 2)
            arcs.Add oArc
            
            Else
            'Create arcs and lines(for the transition)
            Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -DamperLength / 2 - TransitionLength, Depth / 2, (Width - Depth) / 2, _
                            -DamperLength / 2 - TransitionLength, 0, Width / 2, _
                               -DamperLength / 2 - TransitionLength, -Depth / 2, (Width - Depth) / 2)
            arcs.Add oArc
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -DamperLength / 2 - TransitionLength, -Depth / 2, (Width - Depth) / 2, _
                                       -DamperLength / 2 - TransitionLength, -Depth / 2, -(Width - Depth) / 2)
            arcs.Add oLine
            
            Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -DamperLength / 2 - TransitionLength, -Depth / 2, -(Width - Depth) / 2, _
                            -DamperLength / 2 - TransitionLength, 0, -Width / 2, _
                              -DamperLength / 2 - TransitionLength, Depth / 2, -(Width - Depth) / 2)
            arcs.Add oArc
            End If
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -DamperLength / 2 - TransitionLength, Depth / 2, -(Width - Depth) / 2, _
                     -DamperLength / 2 - TransitionLength, Depth / 2, (Width - Depth) / 2)
            arcs.Add oLine
            
            oAxis.Set 1, 0, 0
            
            oStPoint.Set -DamperLength / 2 - TransitionLength, Depth / 2, (Width - Depth) / 2
            Set oCString = PlaceTrCString(oStPoint, arcs)
            Set oProjection = PlaceProjection(m_outputColl, oCString, oAxis, DamperLength + 2 * TransitionLength, True)
                 
            'Set Output(Projection)
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, oProjection
            
            Set oAxis = Nothing
            Set oCString = Nothing
            Set oArc = Nothing
            Set arcs = Nothing
            Set oLine = Nothing
        
        End If
                
        'Create Box(For the actuator)
   
        If CmpDblEqual(BottomOffset, 0) And CmpDblGreaterthan(TopOffset, LINEAR_TOLERANCE) Then
            oStPoint.Set 0.3 * DamperLength, 0.5 * DamperDepth + TopOffset, 0.5 * DamperWidth
            oEnPoint.Set -0.3 * DamperLength, 0, 0.5 * DamperWidth + MotorWidth
          
        ElseIf CmpDblEqual(TopOffset, 0) And CmpDblGreaterthan(BottomOffset, LINEAR_TOLERANCE) Then
            oStPoint.Set 0.3 * DamperLength, 0, 0.5 * DamperWidth
            oEnPoint.Set -0.3 * DamperLength, -0.5 * DamperDepth - BottomOffset, 0.5 * DamperWidth + MotorWidth
                   
        ElseIf CmpDblGreaterthan(TopOffset, LINEAR_TOLERANCE) And CmpDblGreaterthan(BottomOffset, LINEAR_TOLERANCE) Then
            oStPoint.Set 0.3 * DamperLength, 0.5 * DamperDepth + TopOffset, 0.5 * DamperWidth
            oEnPoint.Set -0.3 * DamperLength, -0.5 * DamperDepth - BottomOffset, 0.5 * DamperWidth + MotorWidth
          
        Else
            oStPoint.Set 0.3 * DamperLength, 0.5 * DamperDepth, 0.5 * DamperWidth
            oEnPoint.Set -0.3 * DamperLength, -0.5 * DamperDepth, 0.5 * DamperWidth + MotorWidth
        End If
          
        Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
        
        'Set the output
        lngIndex = lngIndex + 1
        m_outputColl.AddOutput "Feature" & lngIndex, obox
        
        Set obox = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        
    Case 155
              
        If CmpDblEqual(OffsetDepth, 0) Then
            'Create Box (For the Damper)
            oStPoint.Set DamperLength / 2, 0.5 * DamperDepth, 0.5 * DamperWidth
            oEnPoint.Set -DamperLength / 2, -0.5 * DamperDepth, -0.5 * DamperWidth
        Else
            oStPoint.Set DamperLength / 2, 0.5 * DamperDepth - 0.5 * Depth - OffsetDepth, 0.5 * DamperWidth
            oEnPoint.Set -DamperLength / 2, -0.5 * Depth - OffsetDepth, -0.5 * DamperWidth
        End If
        
        Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
          
        'Set the output
        lngIndex = lngIndex + 1
        m_outputColl.AddOutput "Feature" & lngIndex, obox
        Set obox = Nothing
         
        If HVACShape = 4 Then 'Round
             
            'Create Cylinder(For the Transition)
            oStPoint.Set -DamperLength / 2 - TransitionLength, 0, 0
            oEnPoint.Set DamperLength / 2 + TransitionLength, 0, 0
            Set ObjCylinder = PlaceCylinder(m_outputColl, oStPoint, oEnPoint, Width, True)
            
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, ObjCylinder
            Set ObjCylinder = Nothing
            
        End If
            
        If HVACShape = 1 Then 'Rectangular
        
            'Create Rectangular box(For the Transition)
            oStPoint.Set -DamperLength / 2 - TransitionLength, -Depth / 2, -Width / 2
            oEnPoint.Set DamperLength / 2 + TransitionLength, Depth / 2, Width / 2
            Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
          
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, obox
            Set obox = Nothing
            
        End If
            
        If HVACShape = 3 Then 'Oval
        
            'Create arcs and lines(for the transition)
            Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -DamperLength / 2 - TransitionLength, Depth / 2, (Width - Depth) / 2, _
                            -DamperLength / 2 - TransitionLength, 0, Width / 2, _
                               -DamperLength / 2 - TransitionLength, -Depth / 2, (Width - Depth) / 2)
            arcs.Add oArc
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -DamperLength / 2 - TransitionLength, -Depth / 2, (Width - Depth) / 2, _
                                       -DamperLength / 2 - TransitionLength, -Depth / 2, -(Width - Depth) / 2)
            arcs.Add oLine
            
            Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -DamperLength / 2 - TransitionLength, -Depth / 2, -(Width - Depth) / 2, _
                            -DamperLength / 2 - TransitionLength, 0, -Width / 2, _
                              -DamperLength / 2 - TransitionLength, Depth / 2, -(Width - Depth) / 2)
            arcs.Add oArc
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -DamperLength / 2 - TransitionLength, Depth / 2, -(Width - Depth) / 2, _
                     -DamperLength / 2 - TransitionLength, Depth / 2, (Width - Depth) / 2)
            arcs.Add oLine
            
            oAxis.Set 1, 0, 0
            
            oStPoint.Set -DamperLength / 2 - TransitionLength, Depth / 2, (Width - Depth) / 2
            Set oCString = PlaceTrCString(oStPoint, arcs)
            Set oProjection = PlaceProjection(m_outputColl, oCString, oAxis, DamperLength + 2 * TransitionLength, True)
                 
            'Set Output(Projection)
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, oProjection
            
            Set oAxis = Nothing
            Set oCString = Nothing
            Set oArc = Nothing
            Set arcs = Nothing
            Set oLine = Nothing
        
        End If
       
    Case Else
        GoTo ErrorLabel:

    End Select
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
  Exit Sub
      
ErrorLabel:
    LogErrorGer METHOD & ": " & Err.Description
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "FireDamper definition"
    Debug.Assert False
    Exit Sub
    Resume
    
End Sub
 
'=========================================================================
'CREATION OF INSULATION REPRESENTATION OF HVAC FIRE DAMPER
'=========================================================================
Private Sub InsulationOf(ByRef arrayOfInputs())
   
    Const METHOD = "FireDamper-Insulation"
    
    On Error GoTo ErrorLabel

    Dim Width                    As Double
    Dim Depth                    As Double
    Dim FilterLength             As Double
    Dim Thickness                As Double
    Dim FlangeWidth              As Double
    
    Dim CasingWidth              As Double
    Dim CasingThickness          As Double
    Dim DamperLength             As Double
    Dim ActuatorWidth            As Double
    Dim lHvacShape               As Long
    
    Dim HVACShape                As Double
    Dim TopOffset                As Double
    Dim BottomOffset             As Double
    Dim MotorWidth               As Double
    Dim DamperWidth              As Double
    Dim DamperDepth              As Double
    Dim TransitionLength         As Double
    Dim OffsetDepth              As Double
    Dim PlateSpacing             As Double
    Dim Length1                  As Double
    Dim Length                   As Double
    Dim PlateWidth               As Double
    Dim PlateDepth               As Double
    Dim InsulationThickness      As Double
    Dim arcs                     As Collection
    Dim oLine                    As Object
    Dim oProjection              As Object
    Dim oArc                     As IngrGeom3D.Arc3d
    Dim oAxis                    As AutoMath.DVector
    Dim oCString                 As IngrGeom3D.ComplexString3d
    Dim iCount                   As Integer
    Dim oGeomFactory             As IngrGeom3D.GeometryFactory
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oAxis = New AutoMath.DVector
    Set arcs = New Collection

    Dim lngIndex                 As Long
    Dim obox                     As Object
    Dim ObjCylinder              As Object
    Dim oStPoint                 As AutoMath.DPosition
    Dim oEnPoint                 As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim x As Double, y As Double, z As Double
    
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    
    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPart
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
    
    
    If lPartDataBasis = 151 Or lPartDataBasis = 152 Or lPartDataBasis = 153 Or lPartDataBasis = 154 Or _
        lPartDataBasis = 155 Then
        
        'assign to meaningful variables from the input array
    '    Width = arrayOfInputs(2)
    '    Depth = arrayOfInputs(3)
    '    FilterLength = arrayOfInputs(4)
    '    Thickness = arrayOfInputs(5)
    '    FlangeWidth = arrayOfInputs(6)
    '    CasingWidth = arrayOfInputs(7)
    '    CasingThickness = arrayOfInputs(8)
    '    DamperLength = arrayOfInputs(9)
    '    ActuatorWidth = arrayOfInputs(10)
    '    TopOffset = arrayOfInputs(11)
    '    BottomOffset = arrayOfInputs(12)
    '    MotorWidth= arrayOfInputs(13)
    '    DamperWidth = arrayOfInputs(14)
    '    DamperDepth= arrayOfInputs(15)
    '    HVACSShape=arrayOfInputs(16)
    '    PlateSpacing = arrayOfInputs(17)
    '    Length1 = arrayOfInputs(18)
    '    Length = arrayOfInputs(19)
    '    PlateWidth = arrayOfInputs(20)
    '    PlateDepth = arrayOfInputs(21)
    '    Insulationthickness=arrayOfInputs(22)
    '    OffsetDepth=arrayOfInputs(23)
    '    TransitionLength=arrayOfInputs(24)
    '    InsulationThickness = arrayOfInputs(22)
            
      
        Select Case lPartDataBasis
    
            Case 151
                Width = arrayOfInputs(2)
                Depth = arrayOfInputs(3)
                DamperLength = arrayOfInputs(9)
                HVACShape = arrayOfInputs(16)
                InsulationThickness = arrayOfInputs(22)
            
            Case 152
                Width = arrayOfInputs(2)
                Thickness = arrayOfInputs(5)
                PlateSpacing = arrayOfInputs(17)
                Length1 = arrayOfInputs(18)
                Length = arrayOfInputs(19)
                PlateWidth = arrayOfInputs(20)
                PlateDepth = arrayOfInputs(21)
                HVACShape = arrayOfInputs(16)
                InsulationThickness = arrayOfInputs(22)
                lHvacShape = HVACShape
                'As it is a round duct this assumption is taken.
                If CmpDblEqual(Depth, 0) Then
                Depth = Width
                End If
                
            Case 153
                Width = arrayOfInputs(2)
                Depth = arrayOfInputs(3)
                DamperLength = arrayOfInputs(9)
                TopOffset = arrayOfInputs(11)
                BottomOffset = arrayOfInputs(12)
                MotorWidth = arrayOfInputs(13)
                HVACShape = arrayOfInputs(16)
                InsulationThickness = arrayOfInputs(22)
                lHvacShape = HVACShape
                
            Case 154
                Width = arrayOfInputs(2)
                Depth = arrayOfInputs(3)
                DamperLength = arrayOfInputs(9)
                TopOffset = arrayOfInputs(11)
                BottomOffset = arrayOfInputs(12)
                MotorWidth = arrayOfInputs(13)
                DamperWidth = arrayOfInputs(14)
                DamperDepth = arrayOfInputs(15)
                HVACShape = arrayOfInputs(16)
                InsulationThickness = arrayOfInputs(22)
                OffsetDepth = arrayOfInputs(23)
                TransitionLength = arrayOfInputs(24)
                lHvacShape = HVACShape
                
            Case 155
                Width = arrayOfInputs(2)
                Depth = arrayOfInputs(3)
                DamperLength = arrayOfInputs(9)
                DamperWidth = arrayOfInputs(14)
                DamperDepth = arrayOfInputs(15)
                HVACShape = arrayOfInputs(16)
                InsulationThickness = arrayOfInputs(22)
                OffsetDepth = arrayOfInputs(23)
                TransitionLength = arrayOfInputs(24)
                lHvacShape = HVACShape
                 
        End Select
            
        Select Case lPartDataBasis
        
        Case 151
        
            'Create Box
            oStPoint.Set DamperLength / 2, 0.5 * Depth + InsulationThickness, 0.5 * Width + InsulationThickness
            oEnPoint.Set -DamperLength / 2, -0.5 * Depth - InsulationThickness, -0.5 * Width - InsulationThickness
            Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
          
            'Set the output
            lngIndex = lngIndex + 1
            m_outputColl.AddOutput "Feature" & lngIndex, obox
            
            Set obox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
        Case 152
        
            If HVACShape = 4 Then 'Round
        
              'Create Cylinder(For the cylinder)
              oStPoint.Set -Length / 2, 0, 0
              oEnPoint.Set Length / 2, 0, 0
              Set ObjCylinder = PlaceCylinder(m_outputColl, oStPoint, oEnPoint, Width + 2 * InsulationThickness, True)
        
              'Set the output
              lngIndex = lngIndex + 1
              m_outputColl.AddOutput "Feature" & lngIndex, ObjCylinder
              Set ObjCylinder = Nothing
              
              'Create Box (For the plate1)
              oStPoint.Set -Length / 2 + Length1 - InsulationThickness, -PlateDepth / 2 - InsulationThickness, -PlateWidth / 2 - InsulationThickness
              oEnPoint.Set -Length / 2 + Length1 + Thickness + InsulationThickness, PlateDepth / 2 + InsulationThickness, PlateWidth / 2 + InsulationThickness
              Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
            
              'Set the output
              lngIndex = lngIndex + 1
              m_outputColl.AddOutput "Feature" & lngIndex, obox
              Set obox = Nothing
              
              'Create Box (For the plate2)
              oStPoint.Set -Length + Length1 + PlateSpacing - InsulationThickness, -PlateDepth / 2 - InsulationThickness, -PlateWidth / 2 - InsulationThickness
              oEnPoint.Set -Length + Length1 + PlateSpacing + Thickness + InsulationThickness, PlateDepth / 2 + InsulationThickness, PlateWidth / 2 + InsulationThickness
              Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
            
              'Set the output
              lngIndex = lngIndex + 1
              m_outputColl.AddOutput "Feature" & lngIndex, obox
              
              Set obox = Nothing
              Set oStPoint = Nothing
              Set oEnPoint = Nothing
              
            End If
            
        Case 153
            
            If HVACShape = 1 Then ' Rectangular
        
                'Create Box (For the Damper)
                oStPoint.Set DamperLength / 2, 0.5 * Depth + InsulationThickness, 0.5 * Width + InsulationThickness
                oEnPoint.Set -DamperLength / 2, -0.5 * Depth - InsulationThickness, -0.5 * Width - InsulationThickness
                Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
            
                'Set the output
                lngIndex = lngIndex + 1
                m_outputColl.AddOutput "Feature" & lngIndex, obox
                Set obox = Nothing
             
                Set oStPoint = Nothing
                Set oEnPoint = Nothing
              
            End If
            
        Case Is = 154, 155
        
            If CmpDblEqual(OffsetDepth, 0) Then
                
                'Create Box (For the Damper)
                oStPoint.Set DamperLength / 2 + InsulationThickness, 0.5 * DamperDepth + InsulationThickness, 0.5 * DamperWidth + InsulationThickness
                oEnPoint.Set -DamperLength / 2 - InsulationThickness, -0.5 * DamperDepth - InsulationThickness, -0.5 * DamperWidth - InsulationThickness
            Else
                oStPoint.Set DamperLength / 2 + InsulationThickness, 0.5 * DamperDepth + InsulationThickness - 0.5 * Depth - OffsetDepth, 0.5 * DamperWidth + InsulationThickness
                oEnPoint.Set -DamperLength / 2 - InsulationThickness, -0.5 * Depth - InsulationThickness - OffsetDepth, -0.5 * DamperWidth - InsulationThickness
            End If
                
                Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
          
            'Set the output
                lngIndex = lngIndex + 1
                m_outputColl.AddOutput "Feature" & lngIndex, obox
                Set obox = Nothing
             
            If HVACShape = 4 Then 'Round
                 
                'Create Cylinder(For the Transition)
                oStPoint.Set -DamperLength / 2 - TransitionLength, 0, 0
                oEnPoint.Set DamperLength / 2 + TransitionLength, 0, 0
                Set ObjCylinder = PlaceCylinder(m_outputColl, oStPoint, oEnPoint, Width + 2 * InsulationThickness, True)
                
                'Set the output
                lngIndex = lngIndex + 1
                m_outputColl.AddOutput "Feature" & lngIndex, ObjCylinder
                Set ObjCylinder = Nothing
                
            End If
                
            If HVACShape = 1 Then 'Rectangular
                 
                'Create Rectangular box(For the Transition)
                oStPoint.Set -DamperLength / 2 - TransitionLength, -Depth / 2 - InsulationThickness, -Width / 2 - InsulationThickness
                oEnPoint.Set DamperLength / 2 + TransitionLength, Depth / 2 + InsulationThickness, Width / 2 + InsulationThickness
              
                Set obox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
              
                'Set the output
                lngIndex = lngIndex + 1
                m_outputColl.AddOutput "Feature" & lngIndex, obox
                Set obox = Nothing
                
            End If
                
            If HVACShape = 3 Then 'Oval
              If (Width = Depth) Then
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -DamperLength / 2 - TransitionLength, Depth / 2 + InsulationThickness, (Width - Depth) / 2, _
                                  -DamperLength / 2 - TransitionLength, InsulationThickness, Width / 2 + InsulationThickness, _
                                  -DamperLength / 2 - TransitionLength, -Depth / 2 - InsulationThickness, (Width - Depth) / 2)
                arcs.Add oArc
                
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -DamperLength / 2 - TransitionLength, -Depth / 2 - InsulationThickness, -(Width - Depth) / 2, _
                        -DamperLength / 2 - TransitionLength, InsulationThickness, -Width / 2 - InsulationThickness, _
                          -DamperLength / 2 - TransitionLength, Depth / 2 + InsulationThickness, -(Width - Depth) / 2)
                arcs.Add oArc
                
               
               Else
               
               'Create arcs and lines(For the insulation of Transition at front end)
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -DamperLength / 2 - TransitionLength, Depth / 2 + InsulationThickness, (Width - Depth) / 2, _
                                        -DamperLength / 2 - TransitionLength, InsulationThickness, Width / 2 + InsulationThickness, _
                                        -DamperLength / 2 - TransitionLength, -Depth / 2 - InsulationThickness, (Width - Depth) / 2)
                arcs.Add oArc
            
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -DamperLength / 2 - TransitionLength, -Depth / 2 - InsulationThickness, (Width - Depth) / 2, _
                                           -DamperLength / 2 - TransitionLength, -Depth / 2 - InsulationThickness, -(Width - Depth) / 2)
                arcs.Add oLine
            
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -DamperLength / 2 - TransitionLength, -Depth / 2 - InsulationThickness, -(Width - Depth) / 2, _
                                -DamperLength / 2 - TransitionLength, InsulationThickness, -Width / 2 - InsulationThickness, _
                                  -DamperLength / 2 - TransitionLength, Depth / 2 + InsulationThickness, -(Width - Depth) / 2)
                arcs.Add oArc
                            
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -DamperLength / 2 - TransitionLength, Depth / 2 + InsulationThickness, -(Width - Depth) / 2, _
                         -DamperLength / 2 - TransitionLength, Depth / 2 + InsulationThickness, (Width - Depth) / 2)
                arcs.Add oLine
                
            End If
                oAxis.Set 1, 0, 0
            
                oStPoint.Set -DamperLength / 2 - TransitionLength, Depth / 2 + InsulationThickness, (Width - Depth) / 2
                Set oCString = PlaceTrCString(oStPoint, arcs)
                Set oProjection = PlaceProjection(m_outputColl, oCString, oAxis, TransitionLength, True)
                 
                'Set Output(Projection)
                lngIndex = lngIndex + 1
                m_outputColl.AddOutput "Feature" & lngIndex, oProjection
                
                 If (Width = Depth) Then
                 Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, DamperLength / 2 + TransitionLength, Depth / 2 + InsulationThickness, InsulationThickness, _
                        DamperLength / 2 + TransitionLength, InsulationThickness, Width / 2 + InsulationThickness, _
                          DamperLength / 2 + TransitionLength, -Depth / 2 - InsulationThickness, InsulationThickness)
                arcs.Add oArc
                
                
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, DamperLength / 2 + TransitionLength, -Depth / 2 - InsulationThickness, -InsulationThickness, _
                        DamperLength / 2 + TransitionLength, InsulationThickness, -Width / 2 - InsulationThickness, _
                          DamperLength / 2 + TransitionLength, Depth / 2 + InsulationThickness, -InsulationThickness)
                arcs.Add oArc
                
                Else
                'Create arcs and lines(For the insulation of Transition at rear end)
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, DamperLength / 2 + TransitionLength, Depth / 2 + InsulationThickness, (Width - Depth) / 2, _
                                DamperLength / 2 + TransitionLength, InsulationThickness, Width / 2 + InsulationThickness, _
                                   DamperLength / 2 + TransitionLength, -Depth / 2 - InsulationThickness, (Width - Depth) / 2)
                arcs.Add oArc
            
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, DamperLength / 2 + TransitionLength, -Depth / 2 - InsulationThickness, (Width - Depth) / 2, _
                                           DamperLength / 2 + TransitionLength, -Depth / 2 - InsulationThickness, -(Width - Depth) / 2)
                arcs.Add oLine
            
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, DamperLength / 2 + TransitionLength, -Depth / 2 - InsulationThickness, -(Width - Depth) / 2, _
                                DamperLength / 2 + TransitionLength, InsulationThickness, -Width / 2 - InsulationThickness, _
                                  DamperLength / 2 + TransitionLength, Depth / 2 + InsulationThickness, -(Width - Depth) / 2)
                arcs.Add oArc
            
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, DamperLength / 2 + TransitionLength, Depth / 2 + InsulationThickness, -(Width - Depth) / 2, _
                         DamperLength / 2 + TransitionLength, Depth / 2 + InsulationThickness, (Width - Depth) / 2)
                arcs.Add oLine
              
                End If
                oAxis.Set -1, 0, 0
            
                oStPoint.Set DamperLength / 2 + TransitionLength, Depth / 2 + InsulationThickness, (Width - Depth) / 2
                Set oCString = PlaceTrCString(oStPoint, arcs)
                Set oProjection = PlaceProjection(m_outputColl, oCString, oAxis, TransitionLength, True)
                     
                'Set Output(Projection)
                lngIndex = lngIndex + 1
                m_outputColl.AddOutput "Feature" & lngIndex, oProjection
                
                Set oAxis = Nothing
                Set oCString = Nothing
                Set oArc = Nothing
                Set arcs = Nothing
                Set oLine = Nothing
                End If
            
            Set obox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
        Case Else
            GoTo ErrorLabel:
    
        End Select
    End If
  Exit Sub
      
ErrorLabel:
    LogErrorGer METHOD & ": " & Err.Description
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "FireDamper definition"
    Debug.Assert False
    Exit Sub
    Resume
    
End Sub

Private Function createProjectedLines(ByVal objOutputColl As Object, _
                    strName As String, _
                    lngIndex As Long, _
                    objP1 As AutoMath.DPosition, _
                    objP2 As AutoMath.DPosition, _
                    objP3 As AutoMath.DPosition, _
                    objP4 As AutoMath.DPosition, _
                    objVec As AutoMath.DVector, _
                    Optional dblClosed As Boolean = True) As Long
                    
    Dim oLine As IngrGeom3D.Line3d
    Dim iElements As IJElements
    Dim complex As IngrGeom3D.ComplexString3d
    Dim Projection As IJProjection
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP1.x, objP1.y, objP1.z, objP2.x, objP2.y, objP2.z)
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    iElements.Add oLine
    Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP2.x, objP2.y, objP2.z, objP3.x, objP3.y, objP3.z)
    complex.AddCurve oLine, True
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP3.x, objP3.y, objP3.z, objP4.x, objP4.y, objP4.z)
    complex.AddCurve oLine, True
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP4.x, objP4.y, objP4.z, objP1.x, objP1.y, objP1.z)
    complex.AddCurve oLine, True
    
    Set Projection = m_GeomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                    complex, objVec.x, objVec.y, objVec.z, objVec.Length, dblClosed)
                                                    
    If lngIndex >= 0 Then
        lngIndex = lngIndex + 1
        objOutputColl.AddOutput strName & Trim$(Str$(lngIndex)), Projection
    Else
        objOutputColl.AddOutput strName, Projection
    End If
    
     
     
End Function
        


     
Private Function vecDir2(Pin As AutoMath.DPosition, _
                         Dir1 As AutoMath.DVector, dblSize1 As Double, _
                         Optional Dir2 As AutoMath.DVector = Nothing, _
                            Optional dblSize2 As Double = 0#, _
                         Optional Dir3 As AutoMath.DVector = Nothing, _
                            Optional dblSize3 As Double = 0#) As AutoMath.DPosition

Dim Dir1X As AutoMath.DVector
Dim Dir2X As AutoMath.DVector

Set Dir1X = Dir1.Clone
Dir1X.Length = dblSize1
Set vecDir2 = Pin.Offset(Dir1X)

If Not Dir2 Is Nothing Then
    Set Dir2X = Dir2.Clone
    Dir2X.Length = dblSize2
    Set vecDir2 = vecDir2.Offset(Dir2X)
End If
If Not Dir3 Is Nothing Then
    Set Dir2X = Dir3.Clone
    Dir2X.Length = dblSize3
    Set vecDir2 = vecDir2.Offset(Dir2X)
End If


End Function



Private Function createRect(ByVal objOutputColl As Object, _
                    strName As String, _
                    lngIndex As Long, _
                    objP1 As AutoMath.DPosition, _
                    objP2 As AutoMath.DPosition, _
                    objP3 As AutoMath.DPosition, _
                    objP4 As AutoMath.DPosition, _
                    dblToInner As Double, _
                    dirDown As AutoMath.DVector) As Long
                    
' create 4 quader resulting in a rectangle:
'               1             2
'               +-------------+
'               |             |
'               |  +-------+  |
'               |  |       |  |
'               |  |       |  |
'               |  |       |  |
'               |  +-------+  |
'               |             |
'               +-------------+
'               4             3
'
'  P1 - P4 are the 4 outer corner points
'  dblToInner is the size from the outer points to the inner points
'  The size is measured for each side.
'  DirDown is the size and direction in down position
'


' Compute the 4 inner Points
Dim objVec1 As AutoMath.DVector
Dim objVec2 As AutoMath.DVector
 

Dim objP1I As AutoMath.DPosition
Dim objP2I As AutoMath.DPosition
Dim objP3I As AutoMath.DPosition
Dim objP4I As AutoMath.DPosition

 

Set objVec1 = objP2.Subtract(objP1)
Set objVec2 = objP4.Subtract(objP1)
Set objP1I = vecDir2(objP1, objVec1, dblToInner, objVec2, dblToInner)
 
Set objVec1 = objP1.Subtract(objP2)
Set objVec2 = objP3.Subtract(objP2)
Set objP2I = vecDir2(objP2, objVec1, dblToInner, objVec2, dblToInner)

Set objVec1 = objP2.Subtract(objP3)
Set objVec2 = objP4.Subtract(objP3)
Set objP3I = vecDir2(objP3, objVec1, dblToInner, objVec2, dblToInner)

Set objVec1 = objP3.Subtract(objP4)
Set objVec2 = objP1.Subtract(objP4)
Set objP4I = vecDir2(objP4, objVec1, dblToInner, objVec2, dblToInner)

' Compute the 4 curves (which will later be projected)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP1, objP2, objP2I, objP1I, dirDown)
        
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP2, objP3, objP3I, objP2I, dirDown)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP3, objP4, objP4I, objP3I, dirDown)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP4, objP1, objP1I, objP4I, dirDown)

End Function

