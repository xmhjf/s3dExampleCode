VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HROUNDElbow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   HROUNDElbow.cls
'   ProgID:         SP3DHRoundElbow.HROUNDElbow
'   Author:         svsmylav
'   Creation Date:  Monday, Apr 24 2006
'   Description:
'   HVAC Round Elbow
'
'   The symbol implements the following Part Data Bases,
'   a)Elbow, specified by center to port dimensions, width and angle-139
'   b)Elbow, symmetrical, specified by center to port-140
'   c)Elbow, asymmetrical, specified by center to port 1 and center to port 2-141
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   24.Apr.2006     SymbolTeam(India)       DI-94663  New HVAC symbol is prepared from existing
'                                           GSCAD symbol.
'  08.SEP.2006     KKC                      DI-95670  Replace names with initials in all revision history sheets and symbols'  23.Jan.2008     VRK  TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'  26.Jan.2008     VRK                      TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "ElbowServices"    ' These 2 variables are for error.
Private Const UNANTICIPATED = 102
Const LINEAR_TOLERANCE = 0.0000001

Private m_outputColl As IMSSymbolEntities.DOutputCollection

Implements IJDUserSymbolServices

'********************************************************************
' Routine: Class_Terminate
'
' Abstract:The method which is called when this class finishes.
'
' Description:
'
'********************************************************************
Private Sub Class_Terminate()
    Set m_outputColl = Nothing
End Sub

'********************************************************************
' Routine: IJDUserSymbolServices_EditOccurence
'
' Abstract:The method for user to edit parameters of an Elbow symbol
'
' Description:
'
'********************************************************************
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

'********************************************************************
' Routine: IJDUserSymbolServices_InitializeSymbolDefinition
'
' Abstract:The method to initialize symbol definition
'
' Description:This removes the previous definition and sets the new one
' (Input, Output, Representaion, RepresentaionEvaluation)
' This gives default value of its parameters elbow symbol
' using InputSub. This defines graphics to output using OutPutSub.
' It defines Evaluate representation using EvaluateSub.
'
'********************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"
    On Error GoTo ErrorHandler
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    InputSub pSymbolDefinition
    OutPutSub pSymbolDefinition
    EvaluateSub pSymbolDefinition

    '===========================================================================
    'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
    'WHICH ARE GRAPHIC ENTITIES.
    '===========================================================================
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = "SP3DHRoundElbow.HROUNDElbow"
End Function
'********************************************************************
' Routine: IJDUserSymbolServices_InstanciateDefinition
'
' Abstract:The method to instanciate an elbow class object
'
' Description:This instanciates a persistent symbol definition object
' and initialize it for the first time.
'********************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParameters As Variant, ByVal ActiveConnection As Object) As Object

    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition:"
    On Error GoTo ErrorHandler
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition

    Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = "SP3DHRoundElbow.HROUNDElbow"
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

'********************************************************************
' Routine: IJDUserSymbolServices_InvokeRepresentation
'
' Abstract:The method to Invoke Representation
'
' Description: This computes the symbol graphics outputs given the specified
' representation and set of parameters.
'********************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler:
    Set m_outputColl = outputcoll
    If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
    ElseIf StrComp(repName, "InsulationRep") = 0 Then
        InsulationRep arrayOfInputs
    End If

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: Physical
'
' Abstract:The method to create a Physical graphics object
'
' Description:This creates a Physical graphic object by using
' creatnozzle and createlbow method with given values and fixed
' values.
'********************************************************************
Private Sub Physical(ByRef arrayOfInputs())
    Dim elCenter(3) As Double
    Dim circleCenter(3) As Double
    Dim nozzleCenter(3) As Double
    Dim nozzleOrient(6) As Double
    Dim NozzleLength As Double
    Dim Angle As Double

    Dim elbowRadius As Double
    Dim Width As Double
    Dim Throat As Double
    Dim parElbowRadiusplusThroat As Double
    Dim parCentertoPort As Double
    Dim parCentertoPort1 As Double
    Dim parCentertoPort2 As Double
    Dim dTangentLength1 As Double
    Dim dTangentLength2 As Double

    Dim nozzleIndex As Integer
    Dim OutNames As String
    Dim parFacetoCenterofTrimmedE As Double
    Dim dPort2toCenter As Double
    Const METHOD = "Physical:"
    On Error GoTo ErrorHandler

    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    '    Width = arrayOfInputs(2)
    '    elbowRadius = arrayOfInputs(2) / 2
    '    Throat = arrayOfInputs(3)
    '    parElbowRadiusplusThroat = Throat + elbowRadius
    '    Angle = arrayOfInputs(4)
    '    InsulationThickness = arrayOfInputs(5)
    '    parCentertoPort = arrayOfInputs(6)
    '    parCentertoPort1 = arrayOfInputs(7)
    '    parCentertoPort2 = arrayOfInputs(8)


    'Checking for the Part Data Basis Property
    Dim lPartdatabasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPart
    lPartdatabasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
    ' a)Elbow, specified by Center to port dimensions, width and angle-139
    ' b)Elbow, symmetrical, specified by center to port-140
    ' c)Elbow, asymmetrical, specified by center to port 1 and center to port 2-141

    Select Case lPartdatabasis
    Case Is <= 1, 139 'Elbow, specified by center to port dimensions, width and angle
        Width = arrayOfInputs(2)
        elbowRadius = arrayOfInputs(2) / 2
        Throat = arrayOfInputs(3)
        parElbowRadiusplusThroat = Throat + elbowRadius
        Angle = arrayOfInputs(4)
        parFacetoCenterofTrimmedE = parElbowRadiusplusThroat * Tan(Angle / 2)
        dPort2toCenter = parFacetoCenterofTrimmedE

    Case 140 'Elbow, symmetrical, specified by center to port
        Width = arrayOfInputs(2)
        Angle = arrayOfInputs(4)
        parCentertoPort = arrayOfInputs(6)
        elbowRadius = parCentertoPort / Tan(Angle / 2)
        parElbowRadiusplusThroat = elbowRadius
        parFacetoCenterofTrimmedE = parElbowRadiusplusThroat * Tan(Angle / 2)
        dPort2toCenter = parFacetoCenterofTrimmedE

    Case 141 'Elbow, asymmetrical, specified by center to port 1 and center to port 2

        Width = arrayOfInputs(2)
        Angle = arrayOfInputs(4)
        parCentertoPort1 = arrayOfInputs(7)
        parCentertoPort2 = arrayOfInputs(8)

        Dim oStpoint As New AutoMath.DPosition
        Dim oEnpoint As New AutoMath.DPosition
        Dim objTanCyl As Object

        'Elbow radius is computed based on smaller of the two center to port dimensions
        'Tangent length is computed such that tangent is created at the port for which face to center dimension is greater
        If CmpDblLessThanOrEqualTo(parCentertoPort1, parCentertoPort2) Then
            elbowRadius = parCentertoPort1 / Tan(Angle / 2)
            dTangentLength2 = parCentertoPort2 - elbowRadius * Tan(Angle / 2)
            parElbowRadiusplusThroat = elbowRadius
            parFacetoCenterofTrimmedE = parElbowRadiusplusThroat * Tan(Angle / 2)
            oStpoint.Set 0, parCentertoPort2 * Sin(Angle), 0
            oEnpoint.Set 0, (parCentertoPort2 - dTangentLength2) * Sin(Angle), 0
            Set objTanCyl = PlaceCylinder(m_outputColl, oStpoint, oEnpoint, Width, True)
            m_outputColl.AddOutput "ObjTangent", objTanCyl
            Set objTanCyl = Nothing
        Else
            elbowRadius = parCentertoPort2 / Tan(Angle / 2)
            dTangentLength1 = parCentertoPort1 - elbowRadius * Tan(Angle / 2)
            parElbowRadiusplusThroat = elbowRadius
            parFacetoCenterofTrimmedE = parElbowRadiusplusThroat * Tan(Angle / 2)
            oStpoint.Set -parCentertoPort1, 0, 0
            oEnpoint.Set -(parCentertoPort1 - dTangentLength1), 0, 0
            Set objTanCyl = PlaceCylinder(m_outputColl, oStpoint, oEnpoint, Width, True)
            m_outputColl.AddOutput "ObjTangent", objTanCyl
            Set objTanCyl = Nothing
        End If
        Set oStpoint = Nothing
        Set oEnpoint = Nothing
        dPort2toCenter = parCentertoPort2
    Case Else
        GoTo ErrorHandler:

    End Select

    'Length of Nozzle
    NozzleLength = 0

    'Set pt from 0,0,0 to Elbow Center
    elCenter(0) = 0
    elCenter(1) = 0
    elCenter(2) = 0

    circleCenter(0) = -parFacetoCenterofTrimmedE
    circleCenter(1) = 0
    circleCenter(2) = 0


    'Create Pipe whose shape is cylinder
    'NozzleCenter(0,0,0)
    nozzleCenter(0) = -parFacetoCenterofTrimmedE
    nozzleCenter(1) = 0
    nozzleCenter(2) = 0
    'NozzleOrientation0(-1,0,0)
    'NozzleOrientation1(0,0,1)
    'These must be unit vectors for the next subroutine, CreateNozzle
    nozzleOrient(0) = -1
    nozzleOrient(1) = 0
    nozzleOrient(2) = 0
    nozzleOrient(3) = 0
    nozzleOrient(4) = 0
    nozzleOrient(5) = 1
    
    OutNames = "HvacNozzle1"
    nozzleIndex = 1

    CreateNozzle oPart, nozzleIndex, nozzleCenter, nozzleOrient, OutNames, Width, lPartdatabasis


    'Create Elbow
    OutNames = "OutElbow"

    If lPartdatabasis = 140 Or lPartdatabasis = 141 Then
        Dim CP As New AutoMath.DPosition    'arc center point
        Dim CV As New AutoMath.DVector    'rotation vector for rotation
        Dim ObjElbow As Object

        CP.Set -parFacetoCenterofTrimmedE, elbowRadius, 0
        CV.Set 0, 0, 1

        Dim objCircle As IngrGeom3D.Circle3d
        Dim geomFactory As IngrGeom3D.GeometryFactory
        Set geomFactory = New IngrGeom3D.GeometryFactory

        ' Construct a circle that will be used to create the revolution
        Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                         CP.x, 0, 0, _
                                                                         1, 0, 0, _
                                                                         Width / 2)
        Set ObjElbow = PlaceRevolution(m_outputColl, objCircle, _
                                       CV, CP, Angle, True)

        m_outputColl.AddOutput "OutElbow", ObjElbow
        Set ObjElbow = Nothing
        Set CP = Nothing
        Set CV = Nothing
        Set objCircle = Nothing
        Set geomFactory = Nothing
    Else
        'CreateElbow nozzleOrient, elCenter, OutNames, elbowRadius
        'Create Pipe whose shape is cylinder
        'Don't change this order
        CreateElbow nozzleOrient, circleCenter, OutNames, elbowRadius, Angle

    End If

    nozzleCenter(0) = dPort2toCenter * Cos(Angle)
    nozzleCenter(1) = dPort2toCenter * Sin(Angle)
    nozzleCenter(2) = 0


    nozzleOrient(0) = Cos(Angle)
    nozzleOrient(1) = Sin(Angle)
    nozzleOrient(2) = 0
    nozzleOrient(3) = 0
    nozzleOrient(4) = 0
    nozzleOrient(5) = 1
    
    OutNames = "HvacNozzle2"
    nozzleIndex = 2

    CreateNozzle oPart, nozzleIndex, nozzleCenter, nozzleOrient, OutNames, Width, lPartdatabasis

    Set oPart = Nothing

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: InputSub
'
' Abstract:The method to Inputs the defalut value to displayed parameters.
'
' Description: This defines 5 parameters user can change and set the
' default value.
'********************************************************************
Private Sub InputSub(pSymbolDefinition As IJDSymbolDefinition)
' Create a new input by new operator
    Dim Inputs(7) As IMSSymbolEntities.IJDInput

    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent

    Const METHOD = "InputSub:"

    On Error GoTo ErrorHandler
    Set PC = New IMSSymbolEntities.DParameterContent    'not persistent PC

    PC.Type = igValue

    'Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    Set Inputs(0) = New IMSSymbolEntities.DInput
    Inputs(0).Name = "Width"
    Inputs(0).Description = "Diameter of Elbow"
    Inputs(0).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Inputs(0).DefaultParameterValue = PC

    Set Inputs(1) = New IMSSymbolEntities.DInput
    Inputs(1).Name = "ElbowCurveRadius"
    Inputs(1).Description = ""
    Inputs(1).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(1).DefaultParameterValue = PC

    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).Name = "Angle"
    Inputs(2).Description = "Angle of Elbow"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 90#
    Inputs(2).DefaultParameterValue = PC

    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).Name = "InsulationThickness"
    Inputs(3).Description = "Insulation Thickness"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(3).DefaultParameterValue = PC

    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).Name = "CentertoPort"
    Inputs(4).Description = "Center to Port"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(4).DefaultParameterValue = PC

    Set Inputs(5) = New IMSSymbolEntities.DInput
    Inputs(5).Name = "CentertoPort1"
    Inputs(5).Description = "Center to Port 1"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(5).DefaultParameterValue = PC

    Set Inputs(6) = New IMSSymbolEntities.DInput
    Inputs(6).Name = "CentertoPort2"
    Inputs(6).Description = "Center to Port 2"
    Inputs(6).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(6).DefaultParameterValue = PC

    Dim index As Integer
    For index = 0 To 6
        InputsIf.SetInput Inputs(index), index + 2
    Next

    For index = 0 To 6 Step 1
        Set Inputs(index) = Nothing
    Next index
    Set PC = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: OutPutSub
'
' Abstract:The method to define output graphics
'
' Description: This defines the graphics to output.Then call RepSub
' method.
'********************************************************************
Private Sub OutPutSub(pSymbolDefinition As IJDSymbolDefinition)
    Dim i As Integer

    'Create the output
    Dim Outs(4) As IMSSymbolEntities.DOutput

    Const METHOD = "OutPutSub:"
    On Error GoTo ErrorHandler
    For i = 0 To 3 Step 1
        Set Outs(i) = New IMSSymbolEntities.DOutput
        Outs(i).Properties = 0
    Next i

    Outs(0).Name = "HvacNozzle1"
    Outs(0).Description = "PipingPort at End2"

    Outs(1).Name = "HvacNozzle2"
    Outs(1).Description = "PipingPort of End1"

    Outs(2).Name = "OutElbow"
    Outs(2).Description = "Outside Elbow"

    Outs(3).Name = "OutElbowIns"
    Outs(3).Description = "Outside Elbow Insulation"

    'This defines Physical, insulation representations.
    RepSub pSymbolDefinition, Outs

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: RepSub
'
' Abstract:The method to define Representation
'
' Description:  This defines 1 representation which is Physical
'
'********************************************************************
Private Sub RepSub(pSymbolDefinition As IJDSymbolDefinition, Outs() As IMSSymbolEntities.DOutput)
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Dim i As Integer

    Const METHOD = "RepSub:"
    On Error GoTo ErrorHandler

    Set rep1 = New IMSSymbolEntities.DRepresentation

    'Define representation "Physical" and set it on the definition
    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of the Elbow"
    '   rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    rep1.Properties = igCOLLECTION_VARIABLE

    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim oRepPhysicalOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepPhysicalOutputs = rep1

    'Set the outputs
    For i = 0 To 2 Step 1
        oRepPhysicalOutputs.SetOutput Outs(i)
        Set Outs(i) = Nothing
    Next i

    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1

    ' Redefine outputs
    oRepPhysicalOutputs.RemoveAllOutput
    Set oRepPhysicalOutputs = Nothing

    '  Insulation representation
    rep1.Name = "InsulationRep"
    rep1.Description = "Insulation of the Elbow"
    'Set the repID to Insulation.
    rep1.RepresentationId = Insulation
    'rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    rep1.Properties = igCOLLECTION_VARIABLE

    Dim oRepInsulationOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepInsulationOutputs = rep1

    'Set the output
    oRepInsulationOutputs.SetOutput Outs(3)
    Set Outs(3) = Nothing
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1

    Set rep1 = Nothing
    Set oRepPhysicalOutputs = Nothing
    Set oRepInsulationOutputs = Nothing
    Set RepsIf = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: EvaluateSub
'
' Abstract:The method to define representaion evaluation.
'
' Description: This defines representaion evaluation.
'********************************************************************
Private Sub EvaluateSub(pSymbolDefinition As IJDSymbolDefinition)
'Define the evaluation for the Physical representation
    Dim PhysicalRepEval As DRepresentationEvaluation
    Dim RepsEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations

    Const METHOD = "EvaluateSub:"
    On Error GoTo ErrorHandler

    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "Physical representation of Elbow"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "SP3DHRoundElbow.HROUNDElbow"

    ' Define the evaluation associated to the Insulation representation
    Dim InsulationRepEval As DRepresentationEvaluation
    Set InsulationRepEval = New DRepresentationEvaluation
    InsulationRepEval.Name = "InsulationRep"
    InsulationRepEval.Description = "Insulation representation"
    InsulationRepEval.Properties = igREPRESENTATION_HIDDEN
    InsulationRepEval.Type = igREPRESENTATION_VBFUNCTION
    InsulationRepEval.ProgId = "SP3DHRoundElbow.HROUNDElbow"

    '  Set the evaluations on symbol definition
    Set RepsEvalsIf = pSymbolDefinition

    RepsEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    RepsEvalsIf.AddRepresentationEvaluation InsulationRepEval

    Set RepsEvalsIf = Nothing
    Set PhysicalRepEval = Nothing
    Set InsulationRepEval = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: CreateElbow
'
' Abstract:The method to create an elbow.
'
' Description: This creates two elbow surface. Each elbow is created
' by 2 steps. The first one is to create circle. The second step is
' to revolve circle 90 degree around one axis.Each circle has the same
' axis to revolve it. It has the same center point. The difference is
' only diameter.
'********************************************************************
Private Sub CreateElbow(ByRef OrientVec, ByRef elCenter, _
                        ByRef OutNames, ByVal Radius As Double, ByVal Angle As Double)
    Dim oGeomFac As New IngrGeom3D.GeometryFactory
    Dim circleNormal(3) As Double
    Dim RadiusVec(3) As Double
    '''   Dim rAng, dAng As Double
    Dim circleCenter(3) As Double
    Dim oElbow As IngrGeom3D.IJRevolution
    Dim AxesRatio As Double
    Dim oEllipse As IngrGeom3D.Ellipse3d
    Dim AxCenter(3) As Double
    Dim AxNormalVec(3) As Double
    'Dim oTrader As New Trader
    'Dim oUOM As IJUnitsOfMeasure
    Dim oGelem As IJDObject
    Dim i As Integer

    Const METHOD = "CreateElbow:"
    On Error GoTo ErrorHandler

    'CONSTRUCTION OF Outside Elbow

    For i = 0 To 2 Step 1
        circleCenter(i) = elCenter(i)
        circleNormal(i) = OrientVec(i)
        RadiusVec(i) = OrientVec(i + 3) * Radius
        '      RadiusVec(i) = RadiusVec(i)
    Next i
    RadiusVec(2) = -(RadiusVec(2))
    AxesRatio = 1#

    Set oEllipse = oGeomFac.Ellipses3d.CreateByCenterNormMajAxisRatio(m_outputColl.ResourceManager, _
                                                                      circleCenter(0), circleCenter(1), circleCenter(2), _
                                                                      circleNormal(0), circleNormal(1), circleNormal(2), _
                                                                      RadiusVec(0), RadiusVec(1), RadiusVec(2), _
                                                                      AxesRatio)

    '''   dAng = Angle
    'Set oUOM = oTrader.Service("UnitsOfMeasure", "")
    'Set oTrader = Nothing

    'rAng = oUOM.ConvertUnitToUnit(UNIT_ANGLE, dAng, ANGLE_DEGREE, ANGLE_RADIAN)
    '''   rAng = ((22#) / (7# * 180#)) * dAng
    'Set oUOM = Nothing


    AxCenter(0) = circleCenter(0)
    AxCenter(1) = -(circleCenter(0) / Tan(Angle / 2))
    AxCenter(2) = 0#


    AxNormalVec(0) = 0#
    AxNormalVec(1) = 0#
    AxNormalVec(2) = 1#


    If oElbow Is Nothing Then

        Set oElbow = oGeomFac.Revolutions3d.CreateByCurve _
                     (m_outputColl.ResourceManager, oEllipse, _
                      AxNormalVec(0), AxNormalVec(1), AxNormalVec(2), _
                      AxCenter(0), AxCenter(1), AxCenter(2), Angle, False)
    Else
    End If

    Set oGelem = oEllipse
    Set oEllipse = Nothing
    oGelem.Remove
    Set oGelem = Nothing
    m_outputColl.AddOutput OutNames, oElbow
    Set oElbow = Nothing
    Set oGeomFac = Nothing

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: CreateNozzle
'
' Abstract:The method to create a nozzle.
'
' Description: This creates a nozzle.  It has a flange.
'********************************************************************
Private Sub CreateNozzle(ByRef oPart As IJDPart, ByVal nozzleIndex As Integer, _
                         ByRef nozzleCenter, ByRef nozzleOrient, ByRef OutNames, _
                         ByVal Width As Double, ByVal lPartdatabasis As Long)

''Construction of nozzle
''===========================
    Const METHOD = "CreateNozzle:"
    On Error GoTo ErrorHandler

    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle

    Dim iLogicalDistPort As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim Depth As Double
    Dim NozzleLength As Double
    Dim CornerRadius As Double
    Dim inozzleindex As Integer
    'assign to meaningful variables from the input array
    Depth = Width
    inozzleindex = nozzleIndex
    ' HVAC parameters for Nozzle
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 2) As Long
    Dim dThickness(1 To 2) As Double
    Dim dFlangeWidth(1 To 2) As Double
    Dim lFlowDir(1 To 2) As Long
    Dim dPortDepth(1 To 2) As Double
    Dim dCptOffSet(1 To 2) As Double
    
    CornerRadius = 0#
    If lPartdatabasis = 140 Or lPartdatabasis = 141 Then
        'Set HVAC nozzle parameters
        Set oHvacColl = oPart.GetNozzles()
        For nozzleIndex = 1 To oHvacColl.Size
            Set oHvacPort = oHvacColl.Item(nozzleIndex)
            lEndPrep(nozzleIndex) = oHvacPort.EndPrep
            dThickness(nozzleIndex) = oHvacPort.Thickness
            dFlangeWidth(nozzleIndex) = oHvacPort.FlangeWidth
            lFlowDir(nozzleIndex) = oHvacPort.FlowDirection
            dPortDepth(nozzleIndex) = oHvacPort.PortDepth
            dCptOffSet(nozzleIndex) = oHvacPort.CptOffset
        Next nozzleIndex
        
        Set oHvacPort = Nothing
        oHvacColl.Clear
        Set oHvacColl = Nothing
        PortStatus = DistribPortStatus_BASE
        NozzleLength = 0
        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(inozzleindex, "SymbDefn", _
                                GSCADNozzleEntities.Round, lEndPrep(inozzleindex), _
                                dThickness(inozzleindex), dFlangeWidth(inozzleindex), lFlowDir(inozzleindex), Width, _
                                Depth, CornerRadius, DimBaseOuter, PortStatus, _
                               OutNames, dPortDepth(inozzleindex), dCptOffSet(inozzleindex), False, _
                                m_outputColl.ResourceManager)
        Set NozzleFactory = Nothing
        
    ElseIf lPartdatabasis <= 1 Or lPartdatabasis = 139 Then
        Dim Thickness As Double
        Dim FlangeWidth As Double
        Dim PortDepth As Double
        Dim CptOffset As Double
        'assign to meaningful variables from the input array
        Depth = Width
        Dim EndPrep As Long
        Dim FlowDir As DistribFlow
        Dim strPortID As String
        EndPrep = 11
        CornerRadius = 0#
        PortDepth = 0#
        CptOffset = 0#
        'NozzleLength Has to be GREATER than NozzleFlangeThickness
        ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
        ' some non-zero value for flange thickness
        ' Flange thickness assumed to be negigible thickness
        FlangeWidth = 0#
        Thickness = 0
        NozzleLength = 0
        FlowDir = DistribFlow_BOTH
        PortStatus = DistribPortStatus_BASE
        strPortID = "Hvac" & CStr(nozzleIndex)
    
        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(nozzleIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                                         Thickness, FlangeWidth, FlowDir, Width, _
                                                         Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                                         strPortID, PortDepth, CptOffset, False, m_outputColl.ResourceManager)
    
        'Original Piping Elbow nozzle
        'Set oNozzle = NozzleFactory.CreatePipeNozzleFromPart(oPart, nozzleIndex, _
         m_outputColl.ResourceManager)
        Set NozzleFactory = Nothing
    End If
    
    pos.Set nozzleCenter(0), nozzleCenter(1), nozzleCenter(2)
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos
    Set pos = Nothing
     'Direction specified here of the nozzle should be the direction in which pipe will be routed.
     'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Set dir = New AutoMath.DVector
    dir.Set nozzleOrient(0), nozzleOrient(1), nozzleOrient(2)
    iDistribPort.SetDirectionVector dir
    Set dir = Nothing
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength
    m_outputColl.AddOutput OutNames, oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing

    Set iDistribPort = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
'Routine:  ReportUnanticipatedError
'Abstract: Helper functions
'Description:
'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub
'********************************************************************
'Private Sub ReportUnanticipatedError(InMethod As String)
'
'    Dim ern As IMSErrorService.JWellKnownErrorNumbers
'    Dim errorService As IMSErrorLog.IJErrorService
'    Dim oTrader As New Trader
'
'    Set errorService = oTrader.Service("ErrorHandler", "")
'
'    ern = errorService.ReportError(Err.Number, MODULE & ":" & InMethod, "UNANTICIPATED", App)
'
'    Select Case ern
'      Case imsAbortApplication:
'            errorService.TerminateApp
'    End Select
'
'    Set errorService = Nothing
'    Set oTrader = Nothing
'
'End Sub

'********************************************************************
' Routine: Insulation
'
' Abstract:The method to create a Insulation graphics object
'
' Description:This creates a Insulation graphic object
'********************************************************************
Private Sub InsulationRep(ByRef arrayOfInputs())
    Dim elCenter(3) As Double
    Dim circleCenter(3) As Double
    Dim nozzleOrient(6) As Double
    Dim Angle As Double

    Dim elbowRadius As Double
    Dim Width As Double
    Dim Throat As Double
    Dim parElbowRadiusplusThroat As Double
    Dim InsulationThickness As Double
    Dim parCentertoPort As Double
    Dim parCentertoPort1 As Double
    Dim parCentertoPort2 As Double
    Dim dTangentLength1 As Double
    Dim dTangentLength2 As Double

    Dim parFacetoCenterofTrimmedE As Double
    Dim OutNames As String

    Const METHOD = "InsulationRep:"
    On Error GoTo ErrorHandler

    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    '    Width = arrayOfInputs(2)
    '    elbowRadius = arrayOfInputs(2) / 2
    '    Throat = arrayOfInputs(3)
    '    parElbowRadiusplusThroat = Throat + elbowRadius
    '    Angle = arrayOfInputs(4)
    InsulationThickness = arrayOfInputs(5)
    '    parCentertoPort = arrayOfInputs(6)
    '    parCentertoPort1 = arrayOfInputs(7)
    '    parCentertoPort2 = arrayOfInputs(8)


    'Checking for the Part Data Basis Property
    Dim lPartdatabasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPart
    lPartdatabasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
    ' a)Elbow, specified by Center to port dimensions, width and angle-139
    ' b)Elbow, symmetrical, specified by center to port-140
    ' c)Elbow, asymmetrical, specified by center to port 1 and center to port 2-141

    Select Case lPartdatabasis
    Case Is <= 1, 139 'Elbow, specified by Center to port dimensions, width and angle
        Width = arrayOfInputs(2)
        elbowRadius = arrayOfInputs(2) / 2
        Throat = arrayOfInputs(3)
        parElbowRadiusplusThroat = Throat + elbowRadius
        Angle = arrayOfInputs(4)
        parFacetoCenterofTrimmedE = parElbowRadiusplusThroat * Tan(Angle / 2)

    Case 140 'Elbow, symmetrical, specified by center to port
        Width = arrayOfInputs(2)
        Angle = arrayOfInputs(4)
        parCentertoPort = arrayOfInputs(6)
        elbowRadius = parCentertoPort / Tan(Angle / 2)
        parElbowRadiusplusThroat = elbowRadius
        parFacetoCenterofTrimmedE = parElbowRadiusplusThroat * Tan(Angle / 2)

    Case 141 'Elbow, asymmetrical, specified by center to port 1 and center to port 2

        Width = arrayOfInputs(2)
        Angle = arrayOfInputs(4)
        parCentertoPort1 = arrayOfInputs(7)
        parCentertoPort2 = arrayOfInputs(8)

        Dim oStpoint As New AutoMath.DPosition
        Dim oEnpoint As New AutoMath.DPosition
        Dim objTanCyl As Object

        'Elbow radius is computed based on smaller of the two center to port dimensions
        'Tangent length is computed such that tangent is created at the port for which face to center dimension is greater
        If CmpDblLessThanOrEqualTo(parCentertoPort1, parCentertoPort2) Then
            elbowRadius = parCentertoPort1 / Tan(Angle / 2)
            dTangentLength2 = parCentertoPort2 - elbowRadius * Tan(Angle / 2)
            parElbowRadiusplusThroat = elbowRadius
            parFacetoCenterofTrimmedE = parElbowRadiusplusThroat * Tan(Angle / 2)
            oStpoint.Set 0, parCentertoPort2 * Sin(Angle), 0
            oEnpoint.Set 0, (parCentertoPort2 - dTangentLength2) * Sin(Angle), 0
            Set objTanCyl = PlaceCylinder(m_outputColl, oStpoint, oEnpoint, Width + 2 * InsulationThickness, True)
            m_outputColl.AddOutput "ObjInsTangent", objTanCyl
            Set objTanCyl = Nothing
        Else
            elbowRadius = parCentertoPort2 / Tan(Angle / 2)
            dTangentLength1 = parCentertoPort1 - elbowRadius * Tan(Angle / 2)
            parElbowRadiusplusThroat = elbowRadius
            parFacetoCenterofTrimmedE = parElbowRadiusplusThroat * Tan(Angle / 2)
            oStpoint.Set -parCentertoPort1, 0, 0
            oEnpoint.Set -(parCentertoPort1 - dTangentLength1), 0, 0
            Set objTanCyl = PlaceCylinder(m_outputColl, oStpoint, oEnpoint, Width + 2 * InsulationThickness, True)
            m_outputColl.AddOutput "ObjInsTangent", objTanCyl
            Set objTanCyl = Nothing
        End If
        Set oStpoint = Nothing
        Set oEnpoint = Nothing

    Case Else
        GoTo ErrorHandler:

    End Select


    'Set pt from 0,0,0 to Elbow Center
    elCenter(0) = 0
    elCenter(1) = 0
    elCenter(2) = 0

    circleCenter(0) = -parFacetoCenterofTrimmedE
    circleCenter(1) = 0
    circleCenter(2) = 0

    nozzleOrient(0) = -1
    nozzleOrient(1) = 0
    nozzleOrient(2) = 0
    nozzleOrient(3) = 0
    nozzleOrient(4) = 0
    nozzleOrient(5) = 1

    'Create Elbow
    OutNames = "OutElbowIns"

    '   Prepare insulation envelop's elbow radius
    If lPartdatabasis = 140 Or lPartdatabasis = 141 Then
        Dim CP As New AutoMath.DPosition  'arc center point
        Dim CV As New AutoMath.DVector    'rotation vector for rotation
        Dim ObjElbow As Object

        CP.Set -parFacetoCenterofTrimmedE, elbowRadius, 0
        CV.Set 0, 0, 1

        Dim objCircle As IngrGeom3D.Circle3d
        Dim geomFactory As IngrGeom3D.GeometryFactory
        Set geomFactory = New IngrGeom3D.GeometryFactory
        Dim dInsulationRadius As Double
        dInsulationRadius = (Width / 2) + InsulationThickness
            
        'Limit Insulation until parInsulationRadius coincides with the Radius of the Elbow
        If CmpDblLessThanOrEqualTo(dInsulationRadius, elbowRadius) Then
            ' Construct a circle that will be used to create the revolution
            Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                             CP.x, 0, 0, _
                                                                             1, 0, 0, _
                                                                             dInsulationRadius)
            Set ObjElbow = PlaceRevolution(m_outputColl, objCircle, _
                                            CV, CP, Angle, True)
    
            m_outputColl.AddOutput "OutElbowIns", ObjElbow
            Set ObjElbow = Nothing
            Set CV = Nothing
            Set CP = Nothing
            Set objCircle = Nothing
            Set geomFactory = Nothing
        Else
            Dim HalfCord As Double
            HalfCord = Sqr(dInsulationRadius * dInsulationRadius - elbowRadius * elbowRadius)
            Dim points(1 To 13)   As New AutoMath.DPosition
            Dim objBsplineI  As IngrGeom3D.BSplineCurve3d
            points(1).Set -parFacetoCenterofTrimmedE, elbowRadius, 0
            points(2).Set -parFacetoCenterofTrimmedE, elbowRadius, HalfCord
            points(3).Set -parFacetoCenterofTrimmedE, elbowRadius / 2, dInsulationRadius
            points(4).Set -parFacetoCenterofTrimmedE, 0, dInsulationRadius
            points(5).Set -parFacetoCenterofTrimmedE, -dInsulationRadius / 2, dInsulationRadius
            points(6).Set -parFacetoCenterofTrimmedE, -dInsulationRadius, dInsulationRadius / 2
            points(7).Set -parFacetoCenterofTrimmedE, -dInsulationRadius, 0
            points(8).Set -parFacetoCenterofTrimmedE, -dInsulationRadius, -dInsulationRadius / 2
            points(9).Set -parFacetoCenterofTrimmedE, -dInsulationRadius / 2, -dInsulationRadius
            points(10).Set -parFacetoCenterofTrimmedE, 0, -dInsulationRadius
            points(11).Set -parFacetoCenterofTrimmedE, elbowRadius / 2, -dInsulationRadius
            points(12).Set -parFacetoCenterofTrimmedE, elbowRadius, -HalfCord
            points(13).Set -parFacetoCenterofTrimmedE, elbowRadius, 0
            Set objBsplineI = PlaceTrBspline(5, points)
            ' Revolve the Bspline for 90 degree about the Z axis passing by the origin
            Set ObjElbow = PlaceRevolution(m_outputColl, objBsplineI, _
                        CV, CP, Angle, True)
                        Set objBsplineI = Nothing
            m_outputColl.AddOutput "OutElbowIns", ObjElbow
            Set objBsplineI = Nothing
            Set ObjElbow = Nothing
        End If

    Else
        Dim dInsElbowRadius As Double
        If CmpDblLessThan(Throat, InsulationThickness) Then
            InsulationThickness = Throat * 0.99
        End If
        dInsElbowRadius = elbowRadius + InsulationThickness
        CreateElbow nozzleOrient, circleCenter, OutNames, dInsElbowRadius, Angle

    End If

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub

Public Function CmpDblLessThan(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
                               Optional ByVal Tolerance As Double = LINEAR_TOLERANCE) As Double

    If (LeftVariable < (RightVariable + Tolerance)) Then
        CmpDblLessThan = True

    Else
        CmpDblLessThan = False
    End If

End Function
