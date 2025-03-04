VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HElbow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   HElbow.cls
'   ProgID:         SP3DHElbow.HElbow
'   Author:         svsmylav
'   Creation Date:  Monday, Apr 24 2006
'   Description:
'   Hvac Elbow
'
'   The symbol implements the following Part Data Bases,
'   Symmetrical Elbow-
'   a)Elbow, specified by throat radius, width, depth and angle-137
'   b)Elbow, specified by leg lengths, width, depth and angle-138
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   24.Apr.2006     SymbolTeam(India)       DI-94663  New HVAC symbol is prepared from existing
'                                           GSCAD symbol.
'  08.SEP.2006     KKC                      DI-95670  Replace names with initials in all revision history sheets and symbols
'  07.Nov.2007     RUK                      DI-CP·56531 Ability for custom weight and cg calculation for parametric standard components.
'                                           Added code for claculating weight and CG and added one output
'                                                in Physical and one output in Insulation.
'  30.Jan.2008     VRK                      TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'  07.May.2008     VRK                      TR-CP-141832  Two duct ATPs CDuctCopyBPsAndTurns, CDuctBranchTurnCases are failing.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "HElbow"    ' These 2 variables are for error.
Private Const UNANTICIPATED = 102
Private PI As Double
Private RAD As Double
Private m_ProgId As String
Private m_ProjName As String
Private m_ClassName As String

Private m_outputColl As IMSSymbolEntities.DOutputCollection
Const Tolerance = 0.0000001
Const LINEAR_TOLERANCE = 0.0000001
Const NEGLIGIBLE_THICKNESS = 0.0001

Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Dim iCount As Integer

    m_ProjName = "SP3DHElbow"
    m_ClassName = "HElbow"
    m_ProgId = m_ProjName & "." & m_ClassName

    PI = 4 * Atn(1)
    RAD = 180 / PI

    Exit Sub

ErrorHandler:
    ReportUnanticipatedError METHOD

End Sub

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
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, _
                                                    ByVal TransactionMgr As Object) As Boolean
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
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition( _
        pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"

    'msgbox "enter IJDUserSymbolServices_InitializeSymbolDefinition"

    On Error GoTo ErrorHandler

    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    InputSub pSymbolDefinition
    OutPutSub pSymbolDefinition
    EvaluateSub pSymbolDefinition
    'msgbox "past evaluatesub"
    '===========================================================================
    'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
    'WHICH ARE GRAPHIC ENTITIES.
    '===========================================================================
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
    'msgbox "exit IJDUserSymbolServices_InitializeSymbolDefinition"
    Exit Sub

ErrorHandler:
    'msgbox "failed IJDUserSymbolServices_InitializeSymbolDefinition"
    ReportUnanticipatedError METHOD
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Sub
Public Function IJDUserSymbolServices_GetDefinitionName( _
       ByVal definitionParameters As Variant) As String
'msgbox "IJDUserSymbolServices_GetDefinitionName[" & m_ProgId & "]"
'  IJDUserSymbolServices_GetDefinitionName = "SP3DHElbow.HElbow"
    IJDUserSymbolServices_GetDefinitionName = m_ProgId
End Function
'********************************************************************
' Routine: IJDUserSymbolServices_InstanciateDefinition
'
' Abstract:The method to instanciate an elbow class object
'
' Description:This instanciates a persistent symbol definition object
' and initialize it for the first time.
'********************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition( _
       ByVal CodeBase As String, ByVal defParameters As Variant, _
       ByVal ActiveConnection As Object) As Object

    On Error GoTo ErrorHandler
    'msgbox "enter IJDUserSymbolServices_InstanciateDefinition"
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition

    Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

    ' Set definition progId and codebase
    '  oSymbolDefinition.ProgId = "SP3DHElbow.HElbow"
    oSymbolDefinition.ProgId = m_ProgId
    oSymbolDefinition.CodeBase = CodeBase

    ' Give a unique name to the symbol definition
    oSymbolDefinition.Name = oSymbolDefinition.ProgId

    'return symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
    Set oSymbolDefinition = Nothing
    Set oSymbolFactory = Nothing
    'msgbox "exit IJDUserSymbolServices_InstanciateDefinition"

    Exit Function

ErrorHandler:
    'msgbox "failed IJDUserSymbolServices_InstanciateDefinition"
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Function

'********************************************************************
' Routine: IJDUserSymbolServices_InvokeRepresentation
'
' Abstract:The method to Invoke Representation
'
' Description: This computes the symbol graphics outputs given the specified
' representation and set of parameters.
'********************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                                                      ByVal repName As String, _
                                                      ByVal outputcoll As Object, _
                                                      ByRef arrayOfInputs())

    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler
    'msgbox "enter IJDUserSymbolServices_InvokeRepresentation"

    Set m_outputColl = outputcoll
    'msgbox "in IJDUserSymbolServices_InvokeRepresentation repname=" & repName
    If StrComp(repName, "Physical") = 0 Then
        'msgbox "call physical"
        Physical arrayOfInputs
        'msgbox "back from physical"
    ElseIf StrComp(repName, "Insulation") = 0 Then
        InsulationRep arrayOfInputs
    End If
    'msgbox "exit IJDUserSymbolServices_InvokeRepresentation"
    Exit Sub

ErrorHandler:
    'msgbox "failed IJDUserSymbolServices_InvokeRepresentation"
    ReportUnanticipatedError METHOD
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
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
    Dim NozzleLength As Double
    Dim HVACShape As Integer
    Dim Angle As Double
    Dim ElbowThroatRadius As Double
    Dim elbowRadius As Double
    Dim Width As Double
    Dim Depth As Double
    Dim CornerRadius As Double
    Dim OutNames As String
    Dim PlaneofTurn As Integer
    Dim CptOffset As Double
    Dim FlWidth As Double
    Dim PDepth As Double
    Dim parLeglength1 As Double
    Dim parLeglength2 As Double
    Dim parBWidth As Double

    On Error GoTo ErrorHandler

    Dim oPart As PartFacelets.IJDPart
    Dim nozzleIndex As Integer


    Set oPart = arrayOfInputs(1)
    '    HVACShape = arrayOfInputs(2)
    '    Width = arrayOfInputs(3)
    '    Depth = arrayOfInputs(4)
    '    CornerRadius = arrayOfInputs(5)
    '    ElbowThroatRadius = arrayOfInputs(6)
    '    Angle = arrayOfInputs(7)
    '    PlaneofTurn = arrayOfInputs(8)
    '    InsulationThickness = arrayOfInputs(9)
    '    parBWidth=arrayOfInputs(10)
    '    parLeglength1 =arrayOfInputs(11)
    '    parLeglength2 = arrayOfInputs(12)


    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPart
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing

    Select Case lPartDataBasis
'   a)Elbow, specified by throat radius, width, depth and angle-137
'   b)Elbow, specified by leg lengths, width, depth and angle-138
    Case Is <= 1, 137 'Elbow, specified by throat radius, width, depth and angle

        HVACShape = arrayOfInputs(2)
        Width = arrayOfInputs(3)
        Depth = arrayOfInputs(4)
        CornerRadius = arrayOfInputs(5)
        ElbowThroatRadius = arrayOfInputs(6)
        Angle = arrayOfInputs(7)
        PlaneofTurn = arrayOfInputs(8)

        ' Insert your code for output 1(body of the elbow)
        If CmpDblEqual(Angle, 0) Then
            Angle = 0.1 / RAD
        End If

        Dim LL As Double    ' leg length from the origin
        Dim CP As New AutoMath.DPosition    'arc center point
        Dim CV As New AutoMath.DVector    'rotation vector for rotation
        Dim AV As New AutoMath.DVector    'orientation of the circle to sweep
        Dim NO1 As New AutoMath.DVector    'orientation of the profile
        Dim NO2 As New AutoMath.DVector    'orientation of the profile

        'if round make the depth = the width
        If HVACShape = 4 Then Depth = Width

        'check to see that the elbow radius is larger than the profile _
         so the profile doesn't cross the revolution axis

        If PlaneofTurn = 1 Then
            elbowRadius = ElbowThroatRadius + Width / 2
        Else
            elbowRadius = ElbowThroatRadius + Depth / 2
        End If

        '    If Angle <= (90 / RAD) Then
        LL = elbowRadius * Tan(Angle / 2)  'Commented else condition as the formula is valid upto 180 degrees.
        '    Else
        '        LL = elbowRadius
        '    End If

        'Length of Nozzle
        NozzleLength = 0.0002

        CP.Set -LL, 0, 0
        CV.Set -1, 0, 0

        Dim objProfile As Object

        'create the profile for the sweep
        If HVACShape = FlatOval Then
            Set objProfile = CreFlatOval(CP, Width, Depth, PlaneofTurn)
        ElseIf HVACShape = Oval Then
            Set objProfile = CreOval(CP, Width, Depth, PlaneofTurn)
        ElseIf HVACShape = Rectangular Or HVACShape = 0 Then
            Set objProfile = CreRectangle(CP, Width, Depth, CornerRadius, PlaneofTurn)
        ElseIf HVACShape = UShape Then
            Set objProfile = CreUShape(CP, Width, Depth, PlaneofTurn)
        ElseIf HVACShape = 4 Then    'Round=4
            Set objProfile = PlaceTrCircleByCenter(CP, CV, Width / 2)
        End If

        If objProfile Is Nothing Then GoTo ErrorHandler

        Dim axis As New AutoMath.DVector
        Dim objRevolution As Object
        CP.y = elbowRadius
        axis.Set 0, 0, 1

        Set objRevolution = PlaceRevolution(m_outputColl, _
                                            objProfile, axis, CP, Angle, False)

        Set objProfile = Nothing
        m_outputColl.AddOutput "OutElbow", objRevolution
        Set objRevolution = Nothing

        'Create the nozzles

        '    If Angle <= (90 / RAD) Then
        '        LL = -elbowRadius * Tan(Angle / 2) - CptOffset + PDepth
        '    Else
        '        LL = -elbowRadius - CptOffset + PDepth
        '    End If

        '    CP.Set LL, 0, 0
        CP.Set -LL - CptOffset + PDepth, 0, 0
        NO1.Set -1, 0, 0

        If PlaneofTurn = 0 Then
            NO2.Set 0, 1, 0
        Else
            NO2.Set 0, 0, 1
        End If

        OutNames = "HvacNozzle1"
        nozzleIndex = 1

        CreateHVACNozzle oPart, nozzleIndex, CP, NO1, NO2, _
                         OutNames, Width, Depth, CornerRadius, HVACShape

        'create nozzle 2

        ''    If Angle <= (90 / RAD) Then
        ''        LL = elbowRadius * Tan(Angle / 2) + CptOffset - PDepth
        ''        CP.Set LL * Cos(Angle), LL * Sin(Angle), 0
        ''    Else
        ''        LL = elbowRadius * Tan((Angle - (90 / RAD)) / 2)
        ''        CP.Set (LL + CptOffset - PDepth) * Cos(Angle), _
         ''                        elbowRadius + LL + LL * Sin(Angle), 0
        ''    End If

        CP.Set (LL + CptOffset - PDepth) * Cos(Angle), (LL + CptOffset - PDepth) * Sin(Angle), 0

        NO1.Set Cos(Angle), Sin(Angle), 0

        'set the orientation of the nozzle
        If PlaneofTurn = 0 Then
            NO2.Set Cos(Angle + 90 / RAD), Sin(Angle + 90 / RAD), 0
        Else
            NO2.Set 0, 0, 1
        End If

        OutNames = "HvacNozzle2"
        nozzleIndex = 2

        CreateHVACNozzle oPart, nozzleIndex, CP, NO1, NO2, _
                         OutNames, Width, Depth, CornerRadius, HVACShape

        Set oPart = Nothing

    Case 138 'Elbow, specified by leg lengths, width, depth and angle

        HVACShape = arrayOfInputs(2)
        Width = arrayOfInputs(3)
        Depth = arrayOfInputs(4)
        Angle = arrayOfInputs(7)
        PlaneofTurn = arrayOfInputs(8)
        parBWidth = arrayOfInputs(10)
        parLeglength1 = arrayOfInputs(11)
        parLeglength2 = arrayOfInputs(12)

        Dim geomFactory As IngrGeom3D.GeometryFactory
        Set geomFactory = New IngrGeom3D.GeometryFactory
        Dim dlinepts(0 To 17) As Double
        Dim oLine As IngrGeom3D.Line3d
        Dim oArc As IngrGeom3D.EllipticalArc3d
        Dim objCurve As Object
        Dim oComplexStr As IngrGeom3D.ComplexString3d
        Dim objCollection As Collection
        Set objCollection = New Collection

        Dim x, y As Double
        x = parLeglength2 / (Tan(Angle / 2))
        y = parLeglength1 / (Tan(Angle / 2))

        dlinepts(0) = -parLeglength2 * (x + (Width / 2)) / x
        dlinepts(1) = -Depth / 2
        dlinepts(2) = Width / 2

        dlinepts(3) = dlinepts(0)
        dlinepts(4) = -Depth / 2
        dlinepts(5) = -Width / 2

        dlinepts(6) = (parLeglength1 * (y + (parBWidth / 2)) / y) * Cos(Angle) - ((parBWidth / 2) * Sin(Angle))
        dlinepts(7) = -Depth / 2
        dlinepts(8) = -(parLeglength1 * (y + (parBWidth / 2)) / y) * Sin(Angle) - ((parBWidth / 2) * Cos(Angle))
        
        dlinepts(9) = dlinepts(6) + parBWidth * Sin(Angle)
        dlinepts(10) = -Depth / 2
        dlinepts(11) = dlinepts(8) + parBWidth * Cos(Angle)

        'at L1 and L2
        dlinepts(12) = -(((parLeglength2 * (x + Width / 2)) / x) - parLeglength2)
        dlinepts(13) = -Depth / 2
        dlinepts(14) = -Width / 2

        'For Arc
        Dim mMRatio As Double
        Dim a, b As Double
        x = parLeglength2 / (Tan(Angle / 2))
        b = Width + x
        Dim dNum, dDen As Double
        dNum = (dlinepts(9) - dlinepts(0)) ^ 2
        dDen = 1 - (((dlinepts(11) + x + dlinepts(2)) ^ 2) / (b ^ 2))
        a = Sqr(dNum / dDen)
        mMRatio = a / b
        Set oArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                            dlinepts(0), -Depth / 2, -(x + (Width / 2)), _
                            0, 1, 0, _
                            0, 0, x + Width, mMRatio, 0, Angle)
                            
        objCollection.Add oArc
        Set oArc = Nothing

        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                        dlinepts(9), dlinepts(10), dlinepts(11), _
                                                        dlinepts(6), dlinepts(7), dlinepts(8))
        objCollection.Add oLine
        Set oLine = Nothing

        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                        dlinepts(6), dlinepts(7), dlinepts(8), _
                                                        dlinepts(12), dlinepts(13), dlinepts(14))
        objCollection.Add oLine
        Set oLine = Nothing

        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                        dlinepts(12), dlinepts(13), dlinepts(14), _
                                                        dlinepts(3), dlinepts(4), dlinepts(5))
        objCollection.Add oLine
        Set oLine = Nothing

        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                        dlinepts(3), dlinepts(4), dlinepts(5), _
                                                        dlinepts(0), dlinepts(1), dlinepts(2))
        objCollection.Add oLine
        Set oLine = Nothing
        
        Dim oStPoint As AutoMath.DPosition
        Set oStPoint = New AutoMath.DPosition
        oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
        Dim oAxisVec As AutoMath.DVector
        Set oAxisVec = New DVector
        oAxisVec.Set 0, 1, 0
        Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
        Set objCurve = PlaceProjection(m_outputColl, oComplexStr, oAxisVec, Depth, True)
        m_outputColl.AddOutput "objCurve", objCurve
        Set objCurve = Nothing
        
        
        'Create the nozzles
        CP.Set -parLeglength2 * (x + (Width / 2)) / x - CptOffset + PDepth, 0, 0
        NO1.Set -1, 0, 0

        If PlaneofTurn = 0 Then
            NO2.Set 0, 0, 1
        Else
            NO2.Set 0, 1, 0
        End If

        OutNames = "HvacNozzle1"
        nozzleIndex = 1

        CreateHVACNozzle oPart, nozzleIndex, CP, NO1, NO2, _
                         OutNames, Width, Depth, CornerRadius, HVACShape

        'create nozzle 2
        CP.Set ((parLeglength1 * (y + (parBWidth / 2)) / y) + CptOffset - PDepth) * Cos(Angle), 0, -((parLeglength1 * (y + (parBWidth / 2)) / y) + CptOffset - PDepth) * Sin(Angle)
        NO1.Set Cos(Angle), 0, -Sin(Angle)

        'set the orientation of the nozzle
        If PlaneofTurn = 0 Then
            NO2.Set Cos(Angle + 90 / RAD), 0, -Sin(Angle + 90 / RAD)
        Else
            NO2.Set 0, 1, 0
        End If

        OutNames = "HvacNozzle2"
        nozzleIndex = 2

        CreateHVACNozzle oPart, nozzleIndex, CP, NO1, NO2, _
                         OutNames, parBWidth, Depth, CornerRadius, HVACShape

        Set oPart = Nothing

    Case Else
        GoTo ErrorHandler:    'Invalid specification

    End Select

    'New code for WCGOutputs DI-CP·56531
    Dim oGenBOFactory As IJGeneralBusinessObjectsFactory
    Dim oAttributes As IJDAttributes
    Dim oAttribCol As IJDAttributesCol

    Set oGenBOFactory = New GeneralBusinessObjectsFactory
    Set oAttributes = oGenBOFactory.CreateWCGPropHolder(m_outputColl.ResourceManager)

    'In V8 we should use the IJWCGPropHolder interface to set these values
    Set oAttribCol = oAttributes.CollectionOfAttributes(CVar("IJWCGPropHolder"))

    Dim dCGX As Double, dCGY As Double, dCGZ As Double
    Dim dVolume As Double
    Dim dThick As Double

    dThick = 1
    If HVACShape = 4 Then    'Round
        dVolume = PI * Width * (elbowRadius * Angle)
    ElseIf HVACShape = Rectangular Then
        dVolume = 2 * (Width + Depth) * (elbowRadius * Angle)
    ElseIf HVACShape = FlatOval Then
        dVolume = 2 * (Width + (Depth / 2) * PI) * (elbowRadius * Angle)
    End If

    dCGX = ((elbowRadius * Sin(Angle)) / Angle) - LL
    dCGY = elbowRadius - ((elbowRadius * (1 - Cos(Angle))) / Angle)
    dCGZ = 0

    oAttribCol.Item("Volume").Value = dVolume
    oAttribCol.Item("CGX").Value = dCGX
    oAttribCol.Item("CGY").Value = dCGY
    oAttribCol.Item("CGZ").Value = dCGZ

    m_outputColl.AddOutput "StrippedWCG", oAttributes
    Set oAttribCol = Nothing
    Set oAttributes = Nothing

    Exit Sub
ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Sub
'********************************************************************
' Routine: Insulation
'
' Abstract:The method to create a Insulation graphics object
'
' Description:This creates a Insulation graphic object by using
' createlbow method with given values and fixed values.
'********************************************************************
Private Sub InsulationRep(ByRef arrayOfInputs())
    Dim NozzleLength As Double
    Dim HVACShape As Integer
    Dim Angle As Double
    Dim ElbowThroatRadius As Double
    Dim elbowRadius As Double
    Dim Width As Double
    Dim Depth As Double
    Dim CornerRadius As Double
    Dim OutNames As String
    Dim PlaneofTurn As Integer
    Dim CptOffset As Double
    Dim FlWidth As Double
    Dim PDepth As Double
    Dim parLeglength1 As Double
    Dim parLeglength2 As Double
    Dim parBWidth As Double
    Dim InsulationThickness As Double

    On Error GoTo ErrorHandler

    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    '    HVACShape = arrayOfInputs(2)
    '    Width = arrayOfInputs(3)
    '    Depth = arrayOfInputs(4)
    '    CornerRadius = arrayOfInputs(5)
    '    ElbowThroatRadius = arrayOfInputs(6)
    '    Angle = arrayOfInputs(7)
    '    PlaneofTurn = arrayOfInputs(8)
        InsulationThickness = arrayOfInputs(9)
    '    parBWidth=arrayOfInputs(10)
    '    parLeglength1 =arrayOfInputs(11)
    '    parLeglength2 = arrayOfInputs(12)


    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPart
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing

    Select Case lPartDataBasis
        '137,138
    Case Is <= 1, 137 'Elbow, specified by throat radius, width, depth and angle

        HVACShape = arrayOfInputs(2)
        Width = arrayOfInputs(3)
        Depth = arrayOfInputs(4)
        CornerRadius = arrayOfInputs(5)
        ElbowThroatRadius = arrayOfInputs(6)
        Angle = arrayOfInputs(7)
        PlaneofTurn = arrayOfInputs(8)

        'If round make the depth = the width
        If HVACShape = 4 Then Depth = Width

        ' Insert your code for output 1(body of the elbow)
        If CmpDblEqual(Angle, 0) Then Angle = 0.1 / RAD

        '   Check InsulationThickness and adjust it so that elbow radius is larger than
        '   the profile so the profile doesn't cross the revolution axis
        If CmpDblLessThan(ElbowThroatRadius, InsulationThickness) Then
            InsulationThickness = ElbowThroatRadius * 0.99
        End If

        If PlaneofTurn = 1 Then
            elbowRadius = ElbowThroatRadius + Width / 2
        Else
            elbowRadius = ElbowThroatRadius + Depth / 2
        End If

        '   Add insulation thickness to Elbow dimensions
        Width = Width + 2 * InsulationThickness
        Depth = Depth + 2 * InsulationThickness
        CornerRadius = CornerRadius + InsulationThickness

        Dim LL As Double    ' leg length from the origin
        Dim CP As New AutoMath.DPosition    'arc center point
        Dim CV As New AutoMath.DVector    'rotation vector for rotation

        '    If Angle <= (90 / RAD) Then
        LL = elbowRadius * Tan(Angle / 2)
        '    Else
        '        LL = elbowRadius
        '    End If

        CP.Set -LL, 0, 0
        CV.Set -1, 0, 0

        Dim objProfile As Object

        'create the profile for the sweep
        If HVACShape = FlatOval Then
            Set objProfile = CreFlatOval(CP, Width, Depth, PlaneofTurn)
        ElseIf HVACShape = Oval Then
            Set objProfile = CreOval(CP, Width, Depth, PlaneofTurn)
        ElseIf HVACShape = Rectangular Or HVACShape = 0 Then
            Set objProfile = CreRectangle(CP, Width, Depth, CornerRadius, PlaneofTurn)
        ElseIf HVACShape = UShape Then
            Set objProfile = CreUShape(CP, Width, Depth, PlaneofTurn)
        ElseIf HVACShape = 4 Then    'Round=4
            Set objProfile = PlaceTrCircleByCenter(CP, CV, Width / 2)
        End If

        If objProfile Is Nothing Then GoTo ErrorHandler

        Dim axis As New AutoMath.DVector
        Dim objRevolution As Object
        CP.y = elbowRadius
        axis.Set 0, 0, 1

        Set objRevolution = PlaceRevolution(m_outputColl, _
                                            objProfile, axis, CP, Angle, False)

        Set objProfile = Nothing
        m_outputColl.AddOutput "InsElbow", objRevolution
        Set objRevolution = Nothing

        Set CP = Nothing
        Set CV = Nothing
        Set axis = Nothing
        Set oPart = Nothing

    Case 138 'Elbow, specified by leg lengths, width, depth and angle

        HVACShape = arrayOfInputs(2)
        Width = arrayOfInputs(3)
        Depth = arrayOfInputs(4)
        Angle = arrayOfInputs(7)
        PlaneofTurn = arrayOfInputs(8)
        parBWidth = arrayOfInputs(10)
        parLeglength1 = arrayOfInputs(11)
        parLeglength2 = arrayOfInputs(12)

        'Add insulation thickness to Elbow dimensions
        Width = Width + 2 * InsulationThickness
        Depth = Depth + 2 * InsulationThickness
        parBWidth = parBWidth + 2 * InsulationThickness

        Dim geomFactory As IngrGeom3D.GeometryFactory
        Set geomFactory = New IngrGeom3D.GeometryFactory
        Dim dlinepts(0 To 17) As Double
        Dim oLine As IngrGeom3D.Line3d
        Dim oArc As IngrGeom3D.EllipticalArc3d
        Dim objCurve As Object
        Dim oComplexStr As IngrGeom3D.ComplexString3d
        Dim objCollection As Collection
        Set objCollection = New Collection

        Dim x, y As Double
        x = parLeglength2 / (Tan(Angle / 2))
        y = parLeglength1 / (Tan(Angle / 2))

        dlinepts(0) = -parLeglength2 * (x + (Width / 2)) / x
        dlinepts(1) = -Depth / 2
        dlinepts(2) = Width / 2

        dlinepts(3) = dlinepts(0)
        dlinepts(4) = -Depth / 2
        dlinepts(5) = -Width / 2

        dlinepts(6) = (parLeglength1 * (y + (parBWidth / 2)) / y) * Cos(Angle) - ((parBWidth / 2) * Sin(Angle))
        dlinepts(7) = -Depth / 2
        dlinepts(8) = -(parLeglength1 * (y + (parBWidth / 2)) / y) * Sin(Angle) - ((parBWidth / 2) * Cos(Angle))
        
        dlinepts(9) = dlinepts(6) + parBWidth * Sin(Angle)
        dlinepts(10) = -Depth / 2
        dlinepts(11) = dlinepts(8) + parBWidth * Cos(Angle)

        'at L1 and L2
        dlinepts(12) = -(((parLeglength2 * (x + Width / 2)) / x) - parLeglength2)
        dlinepts(13) = -Depth / 2
        dlinepts(14) = -Width / 2

        'For Arc
        Dim mMRatio As Double
        Dim a, b As Double
        x = parLeglength2 / (Tan(Angle / 2))
        b = Width + x
        Dim dNum, dDen As Double
        dNum = (dlinepts(9) - dlinepts(0)) ^ 2
        dDen = 1 - (((dlinepts(11) + x + dlinepts(2)) ^ 2) / (b ^ 2))
        a = Sqr(dNum / dDen)
        mMRatio = a / b
        Set oArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                          dlinepts(0), -Depth / 2, -(x + (Width / 2)), _
                            0, 1, 0, _
                            0, 0, x + Width, mMRatio, 0, Angle)
                            
        objCollection.Add oArc
        Set oArc = Nothing

        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                        dlinepts(9), dlinepts(10), dlinepts(11), _
                                                        dlinepts(6), dlinepts(7), dlinepts(8))
        objCollection.Add oLine
        Set oLine = Nothing

        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                        dlinepts(6), dlinepts(7), dlinepts(8), _
                                                        dlinepts(12), dlinepts(13), dlinepts(14))
        objCollection.Add oLine
        Set oLine = Nothing

        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                        dlinepts(12), dlinepts(13), dlinepts(14), _
                                                        dlinepts(3), dlinepts(4), dlinepts(5))
        objCollection.Add oLine
        Set oLine = Nothing

        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                        dlinepts(3), dlinepts(4), dlinepts(5), _
                                                        dlinepts(0), dlinepts(1), dlinepts(2))
        objCollection.Add oLine
        Set oLine = Nothing

        Dim oStPoint As AutoMath.DPosition
        Set oStPoint = New AutoMath.DPosition
        oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
        Dim oAxisVec As AutoMath.DVector
        Set oAxisVec = New DVector
        oAxisVec.Set 0, 1, 0
        Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
        Set objCurve = PlaceProjection(m_outputColl, oComplexStr, oAxisVec, Depth, True)
        m_outputColl.AddOutput "objInsCurve", objCurve
        Set objCurve = Nothing

    Case Else
        GoTo ErrorHandler:    'Invalid specification

    End Select

    'New code for WCGOutputs DI-CP·56531
    Dim oGenBOFactory As IJGeneralBusinessObjectsFactory
    Dim oAttributes As IJDAttributes
    Dim oAttribCol As IJDAttributesCol

    Set oGenBOFactory = New GeneralBusinessObjectsFactory
    Set oAttributes = oGenBOFactory.CreateWCGPropHolder(m_outputColl.ResourceManager)

    'In V8 we should use the IJWCGPropHolder interface to set these values
    Set oAttribCol = oAttributes.CollectionOfAttributes(CVar("IJWCGPropHolder"))

    Dim dCGX As Double, dCGY As Double, dCGZ As Double
    Dim dPhyVolume As Double    'Volume of the part which persist only in Physical resentation
    Dim dInsVolume As Double    'Volume of the part which persist only in Insulation resentation
    Dim dTotVolume As Double    'Total Volume (Physical + Insulation)
    Dim dThick As Double

    'Calculate Volume by assuming negligible duct thickness
    'Total Volume
    If HVACShape = 4 Then    'Round
        dTotVolume = PI * ((Width / 2) ^ 2) * (elbowRadius * Angle)
    ElseIf HVACShape = Rectangular Then
        dTotVolume = ((Width * Depth) - (PI * (CornerRadius ^ 2))) * elbowRadius * Angle
    ElseIf HVACShape = FlatOval Then
        dTotVolume = ((Width - Depth) * Depth + (PI * (Depth / 2) ^ 2)) * elbowRadius * Angle
    End If

    'Physical Volume
    Dim dActWidth As Double     'Actual Width
    Dim dActDepth As Double     'Actual Depth
    dActWidth = Width - 2 * InsulationThickness
    dActDepth = Depth - 2 * InsulationThickness

    If HVACShape = 4 Then    'Round
        dPhyVolume = PI * ((dActWidth / 2) ^ 2) * elbowRadius * Angle
    ElseIf HVACShape = Rectangular Then
        dPhyVolume = ((dActWidth * dActDepth) - PI * (CornerRadius ^ 2)) _
                   * (elbowRadius * Angle)
    ElseIf HVACShape = FlatOval Then
        dTotVolume = ((dActWidth - dActDepth) * Depth + (PI * (dActDepth / 2) ^ 2)) * _
                     elbowRadius * Angle
    End If

    dInsVolume = dTotVolume - dPhyVolume

    dCGX = ((elbowRadius * Sin(Angle)) / Angle) - LL
    dCGY = elbowRadius - ((elbowRadius * (1 - Cos(Angle))) / Angle)
    dCGZ = 0

    oAttribCol.Item("Volume").Value = dInsVolume
    oAttribCol.Item("CGX").Value = dCGX
    oAttribCol.Item("CGY").Value = dCGY
    oAttribCol.Item("CGZ").Value = dCGZ

    m_outputColl.AddOutput "InsulationWCG", oAttributes
    Set oAttribCol = Nothing
    Set oAttributes = Nothing

    Exit Sub

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
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
    Dim Inputs(11) As IMSSymbolEntities.IJDInput

    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent

    On Error GoTo ErrorHandler
    Set PC = New IMSSymbolEntities.DParameterContent    'not persistent PC
    PC.Type = igValue

    'Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    Dim I As Integer
    I = 0

    Set Inputs(I) = New IMSSymbolEntities.DInput
    Inputs(I).Name = "HVACShape"
    Inputs(I).Description = "HVAC shape of Elbow"
    Inputs(I).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 1    'rectangular default
    Inputs(I).DefaultParameterValue = PC
    I = I + 1

    Set Inputs(I) = New IMSSymbolEntities.DInput
    Inputs(I).Name = "Width"
    Inputs(I).Description = "Width of shape"
    Inputs(I).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Inputs(I).DefaultParameterValue = PC
    I = I + 1

    Set Inputs(I) = New IMSSymbolEntities.DInput
    Inputs(I).Name = "Depth"
    Inputs(I).Description = "Depth of shape"
    Inputs(I).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Inputs(I).DefaultParameterValue = PC
    I = I + 1

    Set Inputs(I) = New IMSSymbolEntities.DInput
    Inputs(I).Name = "CornerRadius"
    Inputs(I).Description = "Corner radius of shape"
    Inputs(I).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(I).DefaultParameterValue = PC
    I = I + 1

    Set Inputs(I) = New IMSSymbolEntities.DInput
    Inputs(I).Name = "ElbowThroatRadius"
    Inputs(I).Description = ""
    Inputs(I).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(I).DefaultParameterValue = PC
    I = I + 1

    Set Inputs(I) = New IMSSymbolEntities.DInput
    Inputs(I).Name = "Angle"
    Inputs(I).Description = "Angle of Elbow"
    Inputs(I).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 2 * Atn(1)    '90 deg
    Inputs(I).DefaultParameterValue = PC
    I = I + 1

    Set Inputs(I) = New IMSSymbolEntities.DInput
    Inputs(I).Name = "PlaneOfTurn"
    Inputs(I).Description = "0 turn around depth, 1 turn around width side"
    Inputs(I).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 1
    Inputs(I).DefaultParameterValue = PC
    I = I + 1

    Set Inputs(I) = New IMSSymbolEntities.DInput
    Inputs(I).Name = "InsulationThickness"
    Inputs(I).Description = "Insulation thickness of body"
    Inputs(I).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(I).DefaultParameterValue = PC
    I = I + 1

    Set Inputs(I) = New IMSSymbolEntities.DInput
    Inputs(I).Name = "BWidth"
    Inputs(I).Description = "Width 2"
    Inputs(I).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(I).DefaultParameterValue = PC
    I = I + 1

    Set Inputs(I) = New IMSSymbolEntities.DInput
    Inputs(I).Name = "LegLength1"
    Inputs(I).Description = "Leg Length 1"
    Inputs(I).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(I).DefaultParameterValue = PC
    I = I + 1

    Set Inputs(I) = New IMSSymbolEntities.DInput
    Inputs(I).Name = "LegLength2"
    Inputs(I).Description = "Leg Length 2"
    Inputs(I).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    PC.UomValue = 0
    Inputs(I).DefaultParameterValue = PC

    Dim index As Integer
    For index = 0 To I
        InputsIf.SetInput Inputs(index), index + 2
    Next

    For index = 0 To 2 Step 1
        Set Inputs(index) = Nothing
    Next index
    Set PC = Nothing
    'msgbox "at the end of inputsub initialize"
    Exit Sub

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
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
    Dim I As Integer

    'Create the output
    Dim Outs(4) As IMSSymbolEntities.DOutput
    On Error GoTo ErrorHandler
    For I = 0 To 3 Step 1
        Set Outs(I) = New IMSSymbolEntities.DOutput
        Outs(I).Properties = 0
    Next I

    Outs(0).Name = "OutElbow"
    Outs(0).Description = "Outside Elbow"

    Outs(1).Name = "HvacNozzle1"
    Outs(1).Description = "HVACPort at End2"

    Outs(2).Name = "HvacNozzle2"
    Outs(2).Description = "HVACPort of End1"

    Outs(3).Name = "InsElbow"
    Outs(3).Description = "Insulation for Elbow"

    '.Aspect = SimplePhysical
    'This defines Physical representation.
    RepSub pSymbolDefinition, Outs
    'msgbox "at the end of outputsub initialize"

    Exit Sub

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
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
    Dim I As Integer

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
    For I = 0 To 2 Step 1
        oRepPhysicalOutputs.SetOutput Outs(I)
        Set Outs(I) = Nothing
    Next I

    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1

    ' Redefine outputs
    oRepPhysicalOutputs.RemoveAllOutput
    Set oRepPhysicalOutputs = Nothing

    '  Insulation representation
    rep1.Name = "Insulation"
    rep1.Description = "Insulation of the Elbow"
    'Set the repID to Insulation. See GSCADSymbolServices library to see
    'different repIDs available.
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
    Set oRepInsulationOutputs = Nothing
    Set RepsIf = Nothing

    Exit Sub

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
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
    Dim InsulationRepEval As DRepresentationEvaluation
    Dim RepsEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    On Error GoTo ErrorHandler

    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "Physical representation of Elbow"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    '   PhysicalRepEval.ProgId = "SP3DHElbow.HElbow"
    PhysicalRepEval.ProgId = m_ProgId

    ' Define the evaluation associated to the Insulation representation
    Set InsulationRepEval = New DRepresentationEvaluation
    InsulationRepEval.Name = "Insulation"
    InsulationRepEval.Description = "Insulation representation"
    InsulationRepEval.Properties = igREPRESENTATION_HIDDEN
    InsulationRepEval.Type = igREPRESENTATION_VBFUNCTION
    InsulationRepEval.ProgId = m_ProgId

    '   Set the evaluations on symbol definition
    Set RepsEvalsIf = pSymbolDefinition

    RepsEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    RepsEvalsIf.AddRepresentationEvaluation InsulationRepEval

    Set RepsEvalsIf = Nothing
    Set PhysicalRepEval = Nothing
    Set InsulationRepEval = Nothing

    Exit Sub

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Sub



'********************************************************************
'Routine:  ReportUnanticipatedError
'Abstract: Helper functions
'Description:
'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub
'********************************************************************
Private Sub ReportUnanticipatedError(InMethod As String)

    Const METHOD = "ReportUnanticipatedError:"
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

    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext

End Sub

Public Function RetrieveHVACParameters(index As Long, _
                                       ByRef partInput As PartFacelets.IJDPart, _
                                       ByVal objOutputColl As Object, _
                                       ByVal Shape As Integer, _
                                       ByRef Width As Double, _
                                       ByRef Depth As Double, _
                                       ByRef CRad As Double, _
                                       ByRef CptOffset As Double, _
                                       ByRef PortDepth As Double, _
                                       ByRef FlangeWidth As Double)

    Dim oHVACPort As GSCADNozzleEntities.IJDHvacPort
    Dim NozzleFactory As GSCADNozzleEntities.NozzleFactory
    Dim oNozzle As GSCADNozzleEntities.IJDNozzle
    Dim gscadElem As IJDObject
    Dim Xsec As IJDOutfittingCrossSection


    Const METHOD = "RetrieveHVACParameters:"
    On Error GoTo ErrorHandler

    Set NozzleFactory = New GSCADNozzleEntities.NozzleFactory

    Set oNozzle = NozzleFactory.CreateHvacNozzleFromPart(partInput, index, False, _
                                                         objOutputColl.ResourceManager)

    Set NozzleFactory = Nothing
    Set oHVACPort = oNozzle

    Set Xsec = oHVACPort.GetCrossSection
    Shape = Xsec.GetShape
    Width = Xsec.Width
    Depth = Xsec.Depth
    CRad = Xsec.CornerRadius
    CptOffset = oHVACPort.CptOffset
    PortDepth = oHVACPort.PortDepth
    FlangeWidth = oHVACPort.FlangeWidth
    '''    Thickness = oHVACPort.Thickness

    ' Release objects
    Set oHVACPort = Nothing
    Set gscadElem = oNozzle
    Set oNozzle = Nothing
    gscadElem.Remove

    Exit Function
ErrorHandler:
    ReportUnanticipatedError METHOD

End Function


'********************************************************************
' Routine: CreateHVACNozzle
'
' Abstract:The method to create a nozzle.
'
' Description: This creates a HVAC nozzle.  It has a flange.
'********************************************************************
Private Sub CreateHVACNozzle(ByRef oPart As IJDPart, _
                             ByVal nozzleIndex As Integer, _
                             ByRef nozzleCenter As AutoMath.DPosition, _
                             ByRef PortDirection As AutoMath.DVector, _
                             ByRef RadialDirection As AutoMath.DVector, _
                             ByRef OutNames, _
                             ByVal Width As Double, _
                             ByVal Depth As Double, _
                             ByVal CornerRadius As Double, _
                             ByVal HVACShape As Integer)

''Construction of nozzle
''===========================
    Const METHOD = "CreateHVACNozzle:"
    On Error GoTo ErrorHandler

    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle

    Dim iLogicalDistPort As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim NozzleLength As Double
    Dim iNozzleIndex As Integer
    iNozzleIndex = nozzleIndex
    
    ' HVAC parameters for Nozzle
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim oHVACPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 2) As Long
    Dim dThickness(1 To 2) As Double
    Dim dFlangeWidth(1 To 2) As Double
    Dim lFlowDir(1 To 2) As Long
    Dim dPortDepth(1 To 2) As Double
    Dim dCptOffSet(1 To 2) As Double
    
    CornerRadius = 0#
    
    'Set HVAC nozzle parameters
    Set oHvacColl = oPart.GetNozzles()
    For nozzleIndex = 1 To oHvacColl.Size
        Set oHVACPort = oHvacColl.Item(nozzleIndex)
        lEndPrep(nozzleIndex) = oHVACPort.EndPrep
        dThickness(nozzleIndex) = oHVACPort.Thickness
        dFlangeWidth(nozzleIndex) = oHVACPort.FlangeWidth
        lFlowDir(nozzleIndex) = oHVACPort.FlowDirection
        dPortDepth(nozzleIndex) = oHVACPort.PortDepth
        dCptOffSet(nozzleIndex) = oHVACPort.CptOffset
    Next nozzleIndex
        
    Set oHVACPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    If CmpDblGreaterthan(dThickness(iNozzleIndex), 0) Then
        NozzleLength = dThickness(iNozzleIndex) + NEGLIGIBLE_THICKNESS
    Else
        NozzleLength = NEGLIGIBLE_THICKNESS
    End If
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHVACNozzle(iNozzleIndex, "SymbDefn", _
                            HVACShape, lEndPrep(iNozzleIndex), _
                            dThickness(iNozzleIndex), dFlangeWidth(iNozzleIndex), lFlowDir(iNozzleIndex), Width, _
                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                            OutNames, dPortDepth(iNozzleIndex), dCptOffSet(iNozzleIndex), False, _
                            m_outputColl.ResourceManager)

    Set NozzleFactory = Nothing

    pos.Set nozzleCenter.x, nozzleCenter.y, nozzleCenter.z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos
    Set pos = Nothing
    '     Direction specified here of the nozzle should be the direction in which
    '     pipe will be routed. Graphics of the nozzle will appear in opposite
    '     direction to the direction specified on the nozzle.
    Set dir = New AutoMath.DVector
    dir.Set PortDirection.x, PortDirection.y, PortDirection.z
    iDistribPort.SetDirectionVector dir

    dir.Set RadialDirection.x, RadialDirection.y, RadialDirection.z
    iDistribPort.SetRadialOrient dir
    Set dir = Nothing
    '
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength
    m_outputColl.AddOutput OutNames, oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing

    Exit Sub

ErrorHandler:
    ReportUnanticipatedError METHOD
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Sub
Public Function CreFlatOval(ByVal centerPoint As AutoMath.DPosition, _
                            ByVal Width As Double, _
                            ByVal Depth As Double, _
                            ByVal orient As Integer) _
                            As IngrGeom3D.ComplexString3d

    Dim Lines As Collection
    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim objCStr As IngrGeom3D.ComplexString3d
    Dim CP As New AutoMath.DPosition
    Dim Pt(6) As New AutoMath.DPosition

    Const METHOD = "CreFlatOval:"
    On Error GoTo ErrorHandler

    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    If orient = 0 Then
        Pt(1).Set CP.x, CP.y + Depth / 2, CP.z - (Width - Depth) / 2
        Pt(2).Set CP.x, Pt(1).y, CP.z + (Width - Depth) / 2
        Pt(3).Set CP.x, CP.y, CP.z + Width / 2
        Pt(4).Set CP.x, CP.y - Depth / 2, Pt(2).z
        Pt(5).Set CP.x, Pt(4).y, Pt(1).z
        Pt(6).Set CP.x, CP.y, CP.z - Width / 2
    Else
        Pt(1).Set CP.x, CP.y - (Width - Depth) / 2, CP.z + Depth / 2
        Pt(2).Set CP.x, CP.y + (Width - Depth) / 2, Pt(1).z
        Pt(3).Set CP.x, CP.y + Width / 2, CP.z
        Pt(4).Set CP.x, Pt(2).y, CP.z - Depth / 2
        Pt(5).Set CP.x, Pt(1).y, Pt(4).z
        Pt(6).Set CP.x, CP.y - Width / 2, CP.z
    End If

    Set oLine = PlaceTrLine(Pt(1), Pt(2))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(2), Pt(4), Pt(3))
    Lines.Add oArc
    Set oLine = PlaceTrLine(Pt(4), Pt(5))
    Lines.Add oLine
    Set oArc = PlaceTrArcBy3Pts(Pt(5), Pt(1), Pt(6))
    Lines.Add oArc

    Set objCStr = PlaceTrCString(Pt(1), Lines)
    Set CreFlatOval = objCStr
    Set oLine = Nothing
    Set oArc = Nothing
    Dim iCount As Integer
    For iCount = 1 To Lines.count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    Exit Function

ErrorHandler:
    ReportUnanticipatedError METHOD

End Function



Public Function PlaceTrCString(ByVal startPosition As AutoMath.DPosition, _
                               ByVal curves As Collection) _
                               As IngrGeom3D.ComplexString3d

''' This function creates transient (non-persistent) complex string curve
''' based on collection of curves (IJElements with StartPoint/EndPoints methods)
''' exapmle of use
'''    Dim lines           As Collection
'''    Dim oLine           As IngrGeom3D.Line3d
'''    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
'''    Dim objCStr         As IngrGeom3D.ComplexString3d
'''
'''    Set lines = New Collection
'''    Set oGeomFactory = New IngrGeom3D.GeometryFactory
'''    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, 0.5, 0, 0)
'''    lines.Add oLine
'''    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, 0, 0.5, 0)
'''    lines.Add oLine
'''    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, -0.5, 0, 0)
'''    lines.Add oLine
'''    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, 0, -0.5, 0)
'''    lines.Add oLine
'''    stPoint.Set 0, 0, -1
'''    Set objCStr = PlaceTrCString(stPoint, lines)
'''    Set oLine = Nothing
'''    Dim iCount As Integer
'''    For iCount = 1 To lines.Count
'''        lines.Remove 1
'''    Next iCount
'''    Set lines = Nothing
'''    ......     use our complex string (e.g. for projection or revolution)
'''    Set objCStr = Nothing

    Const METHOD = "PlaceTrCString:"
    On Error GoTo ErrorHandler

    Dim objCString As IngrGeom3D.ComplexString3d
    Dim curPoint As New AutoMath.DPosition
    Dim curve As IngrGeom3D.IJCurve
    Dim trCurve As IngrGeom3D.IJTransform
    Dim objCurve As Object
    Dim Elems As IJElements
    Set Elems = New JObjectCollection
    Dim vMove As New AutoMath.DVector
    Dim vCompare As New AutoMath.DVector

    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim x1 As Double
    Dim y1 As Double
    Dim z1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim z2 As Double

    Set curPoint = startPosition
    Dim count As Integer
    count = 1
    For Each objCurve In curves
        Set curve = objCurve
        curve.EndPoints x1, y1, z1, x2, y2, z2
        vCompare.Set x2 - x1, y2 - y1, z2 - z1
        If vCompare.Length < Tolerance Then
            '            If vCompare.Length = 0 Then
            '                MsgBox "Start and end points of curve " & count & " are the same"
            '                MsgBox "pts[" & x1 & ", " & y1 & ", " & z1 & "][" & x2 & ", " & y2 & ", " & z2 & "]"
            '            Else
            '                MsgBox "Start and end points of curve " & count & " are too close"
            '            End If
            '            Exit For
        Else    'my addition to skip the short stuff
            '        End If
            count = count + 1
            vMove.Set curPoint.x - x1, curPoint.y - y1, curPoint.z - z1
            Dim tForm As New AutoMath.DT4x4
            tForm.Translate vMove
            Set trCurve = objCurve
            trCurve.Transform tForm
            Set tForm = Nothing
            Elems.Add trCurve
            curve.EndPoints x1, y1, z1, x2, y2, z2
            curPoint.Set x2, y2, z2
        End If    'myaddition to skip short stuff
    Next objCurve

    Set objCString = geomFactory.ComplexStrings3d.CreateByCurves( _
                     Nothing, Elems)

    'MsgBox "out of geomFactory.ComplexStrings3d"

    Set PlaceTrCString = objCString

    Set objCString = Nothing
    Set geomFactory = Nothing
    Elems.Clear
    Set Elems = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError METHOD

End Function
'''<{(Complex string curve end)}>

'''<{(Arcby3pts begin)}>
Public Function PlaceTrArcBy3Pts(ByRef startPoint As AutoMath.DPosition, _
                                 ByRef endPoint As AutoMath.DPosition, _
                                 ByRef onarcPoint As AutoMath.DPosition) _
                                 As IngrGeom3D.Arc3d

''' This function creates transient (non-persistent) arc by start, end, and point on
''' Example of call:
''' Dim points(3)   As new AutoMath.DPosition
''' Dim objArc  As IngrGeom3D.Arc3d
''' points(0).set 1, 0, 0
''' points(1).set 0, 1, 0
''' points(2).set 0, 0, 0
''' set objArc = PlaceTrArcBy3Pts(points(0), points(1), points(2))
''' ......... use this object (e.g. to create projection)
''' set objArc = Nothing


    Const METHOD = "PlaceTrArcBy3Pts:"
    On Error GoTo ErrorHandler

    Dim oArc As IngrGeom3D.Arc3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    'MsgBox "about to create the arc"
    ' Create arc object

    Set oArc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                                  startPoint.x, startPoint.y, startPoint.z, _
                                                  onarcPoint.x, onarcPoint.y, onarcPoint.z, _
                                                  endPoint.x, endPoint.y, endPoint.z)
    Set PlaceTrArcBy3Pts = oArc
    Set oArc = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError METHOD

End Function
'''<{(Arcby3pts end)}>


'''<{(PlaceTrArcByCen begin)}>
Public Function PlaceTrArcByCen( _
       ByRef centerPoint As AutoMath.DPosition, _
       ByRef startPoint As AutoMath.DPosition, _
       ByRef endPoint As AutoMath.DPosition, _
       ByRef normVector As AutoMath.DVector) _
       As IngrGeom3D.Arc3d

''' This function creates transient (non-persistent) arc by start, end, and point on
''' Example of call:
''' Dim points(3)   As new AutoMath.DPosition
''' Dim objArc  As IngrGeom3D.Arc3d
''' points(0).set 1, 0, 0
''' points(1).set 0, 1, 0
''' points(2).set 0, 0, 0
''' set objArc = PlaceTrArcBy3Pts(points(0), points(1), points(2))
''' ......... use this object (e.g. to create projection)
''' set objArc = Nothing


    Const METHOD = "PlaceTrArcByCen:"
    On Error GoTo ErrorHandler

    Dim oArc As IngrGeom3D.Arc3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set oArc = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                          centerPoint.x, centerPoint.y, centerPoint.z, _
                                                          normVector.x, normVector.y, normVector.z, _
                                                          startPoint.x, startPoint.y, startPoint.z, _
                                                          endPoint.x, endPoint.y, endPoint.z)

    Set PlaceTrArcByCen = oArc
    Set oArc = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError METHOD

End Function
'''<{(PlaceTrArcByCen end)}>



'''<{(Line begin)}>
Public Function PlaceTrLine(ByRef startPoint As AutoMath.DPosition, _
                            ByRef endPoint As AutoMath.DPosition) _
                            As IngrGeom3D.Line3d

''' This function creates transient (non-persistent) Line from 2 points
''' Example of call:
''' Dim points(2)   As new AutoMath.DPosition
''' Dim objLine  As IngrGeom3D.Line3d
''' points(0).set 0, 0, 0
''' points(1).set 1, 0, 0
''' set objLine = PlaceTrLine(points(0), points(1))
''' ......... use this object (e.g. to create projection)
''' set objLine = Nothing


    Const METHOD = "PlaceTrLine:"
    On Error GoTo ErrorHandler

    Dim oLine As IngrGeom3D.Line3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    'MsgBox "about to create the line"
    ' Create Line object
    Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                    startPoint.x, startPoint.y, startPoint.z, _
                                                    endPoint.x, endPoint.y, endPoint.z)
    Set PlaceTrLine = oLine
    Set oLine = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError METHOD

End Function
'''<{(Line end)}>

'''<{(Revolution begin)}>
Public Function PlaceRevolution(ByVal objOutputColl As Object, _
                                ByVal objCurve As Object, _
                                ByVal axisVector As AutoMath.DVector, _
                                ByVal centerPoint As AutoMath.DPosition, _
                                revAngle As Double, _
                                isCapped As Boolean)


''' This function creates persistent revolution based on curve
''' axis of revolution and angle
''' Example of call:
''' Dim centPoint   As new AutoMath.DPosition
''' Dim axis        As new AutoMath.DVector
''' Dim objRevolution  As object
''' centPoint.set 0, 0, 0
''' axis.set 0, 1, 0
''' Dim points(1 to 4)   As new AutoMath.DPosition
''' Dim objBspline  As IngrGeom3D.BSplineCurve3d
''' points(0).set 0, 0, 0
''' points(1).set 1, 0, 0
''' points(2).set 1, 1, 0
''' points(3).set 0, 1, 0
''' set objBspline = PlaceTrBspline(points)
''' set objRevolution = PlaceRevolution(m_OutputColl, objBSpline, axis, centPoint, 3.141586/2, True)
''' set objBspline = Nothing
''' m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRevolution
''' Set objRevolution = Nothing

' Construction of revolution
    Const METHOD = "PlaceRevolution:"
    On Error GoTo ErrorHandler

    Dim objRevolution As IngrGeom3D.Revolution3d
    Dim geomFactory As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set objRevolution = geomFactory.Revolutions3d.CreateByCurve( _
                        objOutputColl.ResourceManager, _
                        objCurve, _
                        axisVector.x, axisVector.y, axisVector.z, _
                        centerPoint.x, centerPoint.y, centerPoint.z, _
                        revAngle, isCapped)


    Set PlaceRevolution = objRevolution
    Set objRevolution = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError METHOD

End Function
'''<{(Revolution end)}>

'''<{(Circle begin)}>
Public Function PlaceTrCircleByCenter(ByRef centerPoint As AutoMath.DPosition, _
                                      ByRef normalVector As AutoMath.DVector, _
                                      ByRef Radius As Double) _
                                      As IngrGeom3D.Circle3d

''' This function creates transient (non-persistent) circle
''' Example of call:
''' Dim point   As new AutoMath.DPosition
''' Dim normal  As new AutoMath.DVector
''' Dim objCircle  As IngrGeom3D.circle3d
''' point.set 0, 0, 0
''' normal.set 0, 0, 1
''' set objCircle = PlaceTrCircleByCenter(point, normal, 2 )
''' ......... use this object (e.g. to create projection)
''' set objCircle = Nothing


    Const METHOD = "PlaceTrCircleByCenter:"
    On Error GoTo ErrorHandler

    Dim oCircle As IngrGeom3D.Circle3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    'MsgBox "about to create the Circle"
    ' Create Circle object
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                   centerPoint.x, centerPoint.y, centerPoint.z, _
                                                                   normalVector.x, normalVector.y, normalVector.z, _
                                                                   Radius)
    Set PlaceTrCircleByCenter = oCircle
    Set oCircle = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError METHOD

End Function
'''<{(Circle end)}>
Public Function CreRectangle(ByVal centerPoint As AutoMath.DPosition, _
                             ByVal Width As Double, _
                             ByVal Depth As Double, _
                             ByVal CornerRadius As Double, _
                             ByVal orient As Integer) _
                             As IngrGeom3D.ComplexString3d

    Dim Lines As Collection
    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim objCStr As IngrGeom3D.ComplexString3d
    Dim CP As New AutoMath.DPosition
    Dim Pt(12) As New AutoMath.DPosition

    Const METHOD = "CreRectangle:"
    On Error GoTo ErrorHandler

    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim HD As Double
    Dim HW As Double
    Dim CR As Double
    HD = Depth / 2
    HW = Width / 2
    CR = CornerRadius
    If CR <= Tolerance Then

        If orient = 0 Then
            Pt(1).Set CP.x, CP.y + HD, CP.z - HW
            Pt(2).Set CP.x, CP.y + HD, CP.z + HW
            Pt(3).Set CP.x, CP.y - HD, CP.z + HW
            Pt(4).Set CP.x, CP.y - HD, CP.z - HW
        Else
            Pt(1).Set CP.x, CP.y - HW, CP.z - HD
            Pt(2).Set CP.x, CP.y + HW, CP.z - HD
            Pt(3).Set CP.x, CP.y + HW, CP.z + HD
            Pt(4).Set CP.x, CP.y - HW, CP.z + HD
        End If

        Set oLine = PlaceTrLine(Pt(1), Pt(2))
        Lines.Add oLine
        Set oLine = PlaceTrLine(Pt(2), Pt(3))
        Lines.Add oLine
        Set oLine = PlaceTrLine(Pt(3), Pt(4))
        Lines.Add oLine
        Set oLine = PlaceTrLine(Pt(4), Pt(1))
        Lines.Add oLine
    Else
        Dim NV As New AutoMath.DVector
        NV.Set -1, 0, 0
        If orient = 0 Then
            Pt(1).Set CP.x, CP.y + HD, CP.z - (HW - CR)
            Pt(2).Set CP.x, CP.y + (HD - CR), CP.z - (HW - CR)
            Pt(3).Set CP.x, CP.y + (HD - CR), CP.z - HW
            Pt(4).Set CP.x, CP.y - (HD - CR), CP.z - HW
            Pt(5).Set CP.x, CP.y - (HD - CR), CP.z - (HW - CR)
            Pt(6).Set CP.x, CP.y - HD, CP.z - (HW - CR)
            Pt(7).Set CP.x, CP.y - HD, CP.z + (HW - CR)
            Pt(8).Set CP.x, CP.y - (HD - CR), CP.z + (HW - CR)
            Pt(9).Set CP.x, CP.y - (HD - CR), CP.z + HW
            Pt(10).Set CP.x, CP.y + (HD - CR), CP.z + HW
            Pt(11).Set CP.x, CP.y + (HD - CR), CP.z + (HW - CR)
            Pt(12).Set CP.x, CP.y + HD, CP.z + (HW - CR)

        Else
            Pt(1).Set CP.x, CP.y + HW, CP.z - (HD - CR)
            Pt(2).Set CP.x, CP.y + (HW - CR), CP.z - (HD - CR)
            Pt(3).Set CP.x, CP.y + (HW - CR), CP.z - HD
            Pt(4).Set CP.x, CP.y - (HW - CR), CP.z - HD
            Pt(5).Set CP.x, CP.y - (HW - CR), CP.z - (HD - CR)
            Pt(6).Set CP.x, CP.y - HW, CP.z - (HD - CR)
            Pt(7).Set CP.x, CP.y - HW, CP.z + (HD - CR)
            Pt(8).Set CP.x, CP.y - (HW - CR), CP.z + (HD - CR)
            Pt(9).Set CP.x, CP.y - (HW - CR), CP.z + HD
            Pt(10).Set CP.x, CP.y + (HW - CR), CP.z + HD
            Pt(11).Set CP.x, CP.y + (HW - CR), CP.z + (HD - CR)
            Pt(12).Set CP.x, CP.y + HW, CP.z + (HD - CR)
        End If

        Set oArc = PlaceTrArcByCen(Pt(2), Pt(1), Pt(3), NV)
        Lines.Add oArc

        Set oLine = PlaceTrLine(Pt(3), Pt(4))
        Lines.Add oLine

        Set oArc = PlaceTrArcByCen(Pt(5), Pt(4), Pt(6), NV)
        Lines.Add oArc

        Set oLine = PlaceTrLine(Pt(6), Pt(7))
        Lines.Add oLine

        Set oArc = PlaceTrArcByCen(Pt(8), Pt(7), Pt(9), NV)
        Lines.Add oArc

        Set oLine = PlaceTrLine(Pt(9), Pt(10))
        Lines.Add oLine

        Set oArc = PlaceTrArcByCen(Pt(11), Pt(10), Pt(12), NV)
        Lines.Add oArc

        Set oLine = PlaceTrLine(Pt(12), Pt(1))
        Lines.Add oLine

    End If



    Set objCStr = PlaceTrCString(Pt(1), Lines)
    Set CreRectangle = objCStr
    Set oLine = Nothing

    Dim iCount As Integer
    For iCount = 1 To Lines.count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError METHOD

End Function

Public Function CreUShape(ByVal centerPoint As AutoMath.DPosition, _
                          ByVal Width As Double, _
                          ByVal Depth As Double, _
                          ByVal orient As Integer) _
                          As IngrGeom3D.ComplexString3d

    Dim Lines As Collection
    Dim oLine As IngrGeom3D.Line3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim objCStr As IngrGeom3D.ComplexString3d
    Dim CP As New AutoMath.DPosition
    Dim Pt(6) As New AutoMath.DPosition
    Const METHOD = "CreUShape:"
    On Error GoTo ErrorHandler


    Set CP = centerPoint
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim HD As Double
    Dim HW As Double
    HD = Depth / 2
    HW = Width / 2

    If orient = 0 Then
        Pt(1).Set CP.x, CP.y + HD, CP.z - HW
        Pt(2).Set CP.x, CP.y + HD, CP.z + HW
        Pt(3).Set CP.x, CP.y - HD, CP.z + HW
        Pt(4).Set CP.x, CP.y - HD, CP.z - HW
    Else
        Pt(1).Set CP.x, CP.y - HW, CP.z - HD
        Pt(2).Set CP.x, CP.y + HW, CP.z - HD
        Pt(3).Set CP.x, CP.y + HW, CP.z + HD
        Pt(4).Set CP.x, CP.y - HW, CP.z + HD
    End If

    Set oLine = PlaceTrLine(Pt(2), Pt(3))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(3), Pt(4))
    Lines.Add oLine
    Set oLine = PlaceTrLine(Pt(4), Pt(1))
    Lines.Add oLine

    Set objCStr = PlaceTrCString(Pt(2), Lines)
    Set CreUShape = objCStr

    Set oLine = Nothing
    Dim iCount As Integer
    For iCount = 1 To Lines.count
        Lines.Remove 1
    Next iCount
    Set Lines = Nothing
    Exit Function

ErrorHandler:
    ReportUnanticipatedError METHOD

End Function

Public Function CreOval(ByVal centerPoint As AutoMath.DPosition, _
                        ByVal Width As Double, _
                        ByVal Depth As Double, _
                        ByVal orient As Integer) _
                        As IngrGeom3D.Ellipse3d

    Dim Lines As Collection
    Dim oOval As IngrGeom3D.Ellipse3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim objCStr As IngrGeom3D.ComplexString3d
    Dim CP As New AutoMath.DPosition
    Dim Pt(6) As New AutoMath.DPosition
    Dim MX As New AutoMath.DPosition
    Dim normalVec As New AutoMath.DVector

    Const METHOD = "CreOval:"
    On Error GoTo ErrorHandler

    Set CP = centerPoint
    normalVec.Set -1, 0, 0
    Set Lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    If orient = 0 Then
        MX.Set 0, 0, Width / 2
    Else
        MX.Set 0, Width / 2, 0
    End If

    'MsgBox "about to create the ellipse"
    'Create ellipse object
    Set oOval = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio( _
                Nothing, _
                CP.x, CP.y, CP.z, _
                normalVec.x, normalVec.y, normalVec.z, _
                MX.x, MX.y, MX.z, Depth / Width)
    Set CreOval = oOval
    Set oOval = Nothing
    Set oGeomFactory = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError METHOD

End Function

Public Function CmpDblEqual(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
                            Optional ByVal Tolerance As Double = LINEAR_TOLERANCE) As Double

    If (LeftVariable >= (RightVariable - Tolerance)) And _
       (LeftVariable <= (RightVariable + Tolerance)) Then
        CmpDblEqual = True

    Else
        CmpDblEqual = False
    End If

End Function

Public Function CmpDblLessThan(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
                               Optional ByVal Tolerance As Double = LINEAR_TOLERANCE) As Double

    If (LeftVariable < (RightVariable + Tolerance)) Then
        CmpDblLessThan = True

    Else
        CmpDblLessThan = False
    End If

End Function

