VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ElbowNozzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    '******************************************************************
' Copyright (C) 2016, Intergraph Corporation. All rights reserved.
'
'File
'    ElbowNozzleServices.cls
'
'Author
'    MU
'
'Description
'    Definition of ElbowNozzleServices Symbol
'History
'   29  Nov  Mule Creation
'   03-Jul-02 VS     Changed the X and Z directions to be consistent with the straight nozzles
'   17-Nov-09 GUK    Added insulation aspect to elbow nozzle as a part of 3D-Accessors-Sprint-2009-11
'
'Notes
'
'    <notes>
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  20.SEP.2016    Navateja  DM-CP-297135       Error at port location on female Piping Elbow Nozzle
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const E_FAIL = -2147467259
Private Const TOLERANCE = 0.000001
private Const ANGULARTOLERANCE = 0.0000000001 'Relative tolerance
Dim m_outputColl As IJDOutputCollection
Dim m_bApplyDepthAndOffset As Boolean
Implements IJDUserSymbolServices
Private PI As Double

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    On Error GoTo ErrorHandler
    IJDUserSymbolServices_EditOccurence = False
     Exit Function
ErrorHandler:
    Err.Raise E_FAIL
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    On Error GoTo ErrorHandler
    IJDUserSymbolServices_GetDefinitionName = "GSCADNozzleVBsymb.ElbowNozzle"
    Exit Function
ErrorHandler:
    Err.Raise E_FAIL
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

    On Error GoTo ErrorHandler
    
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    'Feed ElbowNozzle Defnition
    ' Inputs:
    '          1. "FlangeDiameter"
    '          2. "FlangeThickness"
    '          3. "PipingDiameter"
    '          4. "PipeLength"
    '          5. "CptOffset"
    '          6. "Depth"
    '          7. "BendLength"
    '          8. "BendRadius"
    '          9. "RaisedFaceOrSocketDiameter"
    '         10. "InsulationCriteria"
    '         11. "InsulationThickness"
    '         12. "BendAngle"
    ' Representations :
    '           Physical
    
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    ' Create a new input by new operator
    Dim Inputs(1 To 12) As IMSSymbolEntities.DInput
    
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
    
    'get value of PI
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    PI = oGeomFactory.GeometryServices.GeomPI
    Set oGeomFactory = Nothing
    
    'Define inputs for Elbow Nozzle
    Dim index As Integer
    For index = 1 To 12
        Set Inputs(index) = New IMSSymbolEntities.DInput
        Inputs(index).Properties = igINPUT_IS_A_PARAMETER
        Select Case index
             Case 1:
                 PC.UomValue = 0.2
                 Inputs(index).Name = "FlangeDiameter"
                 Inputs(index).Description = "Diameter of the Flange"
            Case 2:
                 PC.UomValue = 0.02
                 Inputs(index).Name = "FlangeThickness"
                 Inputs(index).Description = "Thickness of the Flange"
            Case 3:
                 PC.UomValue = 0.08
                 Inputs(index).Name = "PipingDiameter"
                 Inputs(index).Description = "Diameter of the connector pipe"
            Case 4:
                 PC.UomValue = 0
                 Inputs(index).Name = "PipeLength"
                 Inputs(index).Description = "Length of the connector pipe"
            Case 5:
                 PC.UomValue = 0
                 Inputs(index).Name = "CptOffset"
                 Inputs(index).Description = "Connect Point Offset"
            Case 6:
                 PC.UomValue = 0.02
                 Inputs(index).Name = "Depth"
                 Inputs(index).Description = "Connect Depth"
            Case 7:
                 PC.UomValue = 0.2
                 Inputs(index).Name = "BendLength"
            Case 8:
                 PC.UomValue = 0.08
                 Inputs(index).Name = "BendRadius"
            Case 9:
                 PC.UomValue = 0.14
                 Inputs(index).Name = "RaisedFaceOrSocketDiameter"
            Case 10:
                 PC.UomValue = 0#
                 Inputs(index).Name = "InsulationCriteria"
            Case 11:
                 PC.UomValue = -1#
                 Inputs(index).Name = "InsulationThickness"
            Case 12:
                 PC.UomValue = PI / 2# 
                 Inputs(index).Name = "BendAngle"
                 
        End Select
        
        Inputs(index).DefaultParameterValue = PC
    Next
    
    For index = 1 To 12
        InputsIf.SetInput Inputs(index), index
        Set Inputs(index) = Nothing
    Next
    
    'Define the outputs
    Dim O1 As IMSSymbolEntities.DOutput
    Set O1 = New IMSSymbolEntities.DOutput

    O1.Name = "Cylinder"
    O1.Description = "RightCircularCylinder"
    O1.Properties = 0

    Dim O2 As IMSSymbolEntities.DOutput
    Set O2 = New IMSSymbolEntities.DOutput
    
    O2.Name = "Tori"
    O2.Description = "RightCircularTorus"
    O2.Properties = 0
    
    Dim O3 As IMSSymbolEntities.DOutput
    Set O3 = New IMSSymbolEntities.DOutput
    
    O3.Name = "Cylinder2"
    O3.Description = "RightCircularCylinder"
    O3.Properties = 0
    
    Dim O4 As IMSSymbolEntities.DOutput
    Set O4 = New IMSSymbolEntities.DOutput
    
    O4.Name = "Cylinder3"
    O4.Description = "RightCircularCylinder"
    O4.Properties = 0
    
    Dim O5 As IMSSymbolEntities.DOutput
    Set O5 = New IMSSymbolEntities.DOutput
    
    O5.Name = "Cylinder4"
    O5.Description = "RightCircularCylinder"
    O5.Properties = 0
    'Define the representation "Symbolic"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation

    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of ElbowNozzle"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim IJDOutputs As IMSSymbolEntities.IJDOutputs
    Set IJDOutputs = rep1

    ' Set the output
    IJDOutputs.SetOutput O1
    Set O1 = Nothing
    
     ' Set the output
    IJDOutputs.SetOutput O2
    Set O2 = Nothing
    
     ' Set the output
    IJDOutputs.SetOutput O3
    Set O3 = Nothing
    
     ' Set the output
    IJDOutputs.SetOutput O4
    Set O4 = Nothing
    
    If CheckIfDepthandoffsetFileExists = False Then
         ' Set the output
        IJDOutputs.SetOutput O5
        Set O5 = Nothing
    End If
    
    'Set the representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1

    Set rep1 = Nothing
    Set RepsIf = Nothing
    Set IJDOutputs = Nothing
    
    'Define evaluation for Physical representation
    Dim PhysicalRepEval As IJDRepresentationEvaluation
    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "Physical representation"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "GSCADNozzleVBsymb.ElbowNozzle"
    
    'Set the evaluations for the Symbolic and Physical representations on the definition
    Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    Set PhysicalRepEval = Nothing
    Set RepEvalsIf = Nothing
    
'=============================================================================================
'  Insulation Representation of Nozzle
'
'==============================================================================================
    ' Insulation output of base
    Dim OBaseInsul As IMSSymbolEntities.DOutput
    Set OBaseInsul = New IMSSymbolEntities.DOutput
    OBaseInsul.Name = "BaseInsulation"
    OBaseInsul.Description = "Surface of Revolution"
    OBaseInsul.Properties = 0
    
    ' Insulation output of Torus part of Elbow Nozzle
    Dim OToriInsul As IMSSymbolEntities.DOutput
    Set OToriInsul = New IMSSymbolEntities.DOutput
    OToriInsul.Name = "ToriInsulation"
    OToriInsul.Description = "RightCircularTorus"
    OToriInsul.Properties = 0
    
    'Insulation output of Neck part of Elbow Nozzle
    Dim ONeckInsul As IMSSymbolEntities.DOutput
    Set ONeckInsul = New IMSSymbolEntities.DOutput
    ONeckInsul.Name = "NeckInsulation"
    ONeckInsul.Description = "Surface of Revolution"
    ONeckInsul.Properties = 0
    
    'Insulation out put of Flange
    Dim OFlngInsul As IMSSymbolEntities.DOutput
    Set OFlngInsul = New IMSSymbolEntities.DOutput
    OFlngInsul.Name = "FlangeInsulation"
    OFlngInsul.Description = "Surface of Revolution"
    OFlngInsul.Properties = 0
    
    'Define the representation "Symbolic"
    Dim repInsul As IMSSymbolEntities.DRepresentation
    Set repInsul = New IMSSymbolEntities.DRepresentation
    repInsul.Name = "Insulation"
    repInsul.Description = "Insulation Representation of ElbowNozzle"
    repInsul.Properties = igREPRESENTATION_ISVBFUNCTION
    repInsul.RepresentationId = Insulation
    
    'Set the outputs
    Dim IJDOutputsInsul As IMSSymbolEntities.IJDOutputs
    Set IJDOutputsInsul = repInsul
    
    ' Set the output
    IJDOutputsInsul.SetOutput OBaseInsul
    Set OBaseInsul = Nothing
    
    ' Set the output
    IJDOutputsInsul.SetOutput OToriInsul
    Set OToriInsul = Nothing
    
    ' Set the output
    IJDOutputsInsul.SetOutput ONeckInsul
    Set ONeckInsul = Nothing
    
    ' Set the output
    IJDOutputsInsul.SetOutput OFlngInsul
    Set OFlngInsul = Nothing
    
    'Set the representation of Insulation to definition
    Dim RepsInsulIf As IMSSymbolEntities.IJDRepresentations
    Set RepsInsulIf = pSymbolDefinition
    RepsInsulIf.SetRepresentation repInsul

    Set repInsul = Nothing
    Set RepsInsulIf = Nothing
    Set IJDOutputsInsul = Nothing
    'Set the evaluation progids of the insulation representation, that we are evaluating the
    'Geometry onlyy this progid itself
    'Define evaluation for Insualtion representation
    Dim InsulationRepEval As IJDRepresentationEvaluation
    Set InsulationRepEval = New DRepresentationEvaluation
    InsulationRepEval.Name = "Insulation"
    InsulationRepEval.Description = "Insulation representation"
    InsulationRepEval.Properties = igREPRESENTATION_HIDDEN
    InsulationRepEval.Type = igREPRESENTATION_VBFUNCTION
    InsulationRepEval.ProgId = "GSCADNozzleVBsymb.ElbowNozzle"
    
    'Set the evaluations for the Symbolic and Physical representations on the definition
    Dim InsulationRepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set InsulationRepEvalsIf = pSymbolDefinition
    InsulationRepEvalsIf.AddRepresentationEvaluation InsulationRepEval
    Set InsulationRepEval = Nothing
    Set InsulationRepEvalsIf = Nothing
    
'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

Exit Sub
ErrorHandler:
  Err.Raise E_FAIL
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal ActiveConnection As Object) As Object
    On Error GoTo ErrorHandler
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
    Set oSymbolFactory = Nothing
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = "GSCADNozzleVBsymb.ElbowNozzle"
    oSymbolDefinition.CodeBase = CodeBase
    
    ' Give a unique name to the symbol definition
    oSymbolDefinition.Name = oSymbolDefinition.ProgId
    
    'return symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
    Set oSymbolDefinition = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise E_FAIL
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal RepName As String, ByVal OutputColl As Object, arrayOfInputs() As Variant)
    On Error GoTo ErrorHandler
    Set m_outputColl = OutputColl
    
    If CheckIfDepthandoffsetFileExists = True Then
        m_bApplyDepthAndOffset = True
    Else
        m_bApplyDepthAndOffset = False
    End If
    
    'get value of PI
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    PI = oGeomFactory.GeometryServices.GeomPI
    Set oGeomFactory = Nothing
    
    Dim dBAngle As Double
    dBAngle = arrayOfInputs(12)
    If StrComp(RepName, "Physical") = 0 Then
        Physical arrayOfInputs
    ElseIf StrComp(RepName, "Insulation") = 0 Then
        If (Abs(dBAngle - PI / 2#)) < ANGULARTOLERANCE Then
            InsulateElbowNozzle arrayOfInputs
        Else
            InsulateElbowNozzleForBendAngeleOtherThan90 arrayOfInputs 'If bend angle is otherthan 90
        End If
    Else
        'Do nothing this case cannot come at all
    End If
Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF Elbow Nozzle
'=========================================================================
Private Sub Physical(ByRef arrayOfInputs())
    On Error GoTo ErrorLabel
    Dim a As Double, b As Double, c As Double, dBendAngle As Double
    Dim dFlangeDiameter As Double, dFlangeThickness As Double, dDepth As Double, dRaisedFaceOrSocketDiameter As Double, dPipingDiameter As Double, dPipeLength As Double, dCptOffset As Double
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    
    dFlangeDiameter = arrayOfInputs(1)
    dFlangeThickness = arrayOfInputs(2)
    dPipingDiameter = arrayOfInputs(3)
    a = arrayOfInputs(4) 'a = Pipe Length
    dCptOffset = arrayOfInputs(5)
    dDepth = arrayOfInputs(6)
    b = arrayOfInputs(7) 'b = Bend Length 'These might be interchanged later.
    c = arrayOfInputs(8) 'c = Bend Radius
    dRaisedFaceOrSocketDiameter = arrayOfInputs(9)
    dBendAngle = arrayOfInputs(12)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    oErrors.Clear
    If (dRaisedFaceOrSocketDiameter = 0) Then: dRaisedFaceOrSocketDiameter = dPipingDiameter
    If b <= 0 Or c <= 0 Or dPipingDiameter <= 0 Or dBendAngle <= 0 Then
        Dim strx As String
        strx = b & ":" & c & ":" & dPipingDiameter & ":" & dBendAngle
        oErrors.Add E_FAIL, "GSCADNozzleVBsymb.ElbowNozzle", strx, "ZeroOrNegative"
        GoTo ErrorLabel
    ElseIf c < dPipingDiameter / 2 Then
        oErrors.Add E_FAIL, "GSCADNozzleVBsymb.ElbowNozzle", "The Combination of Npd,NpdUnitType,EndPreparation,ScheduleThickness,EndStandard,PressureRating makes the input invalid", "MajorGreaterThanMinor"
        GoTo ErrorLabel
    ElseIf b < c Then
        oErrors.Add E_FAIL, "GSCADNozzleVBsymb.ElbowNozzle", "B must be Greater than or equal to C", "B<C"
        GoTo ErrorLabel
    End If
    
    Dim lIndex As Integer
    Dim oCylStartPos As New DPosition
    Dim dCyl1Length As Double, dCyl2Length As Double
    Dim oGeomFactory As GeometryFactory
    Dim ObjElbow As Object
    
    If m_bApplyDepthAndOffset = False Then
    
        If Abs(dBendAngle - (PI / 2)) < ANGULARTOLERANCE Then '  if bend angle is 90 deg
            lIndex = 1
            If (a - c) > 0 Then: CreateCylinder dPipingDiameter, a - c, b, a, 0, 0, -1, 0, 0, -1, 0, "Cylinder"
            
            lIndex = 2
            If ((b - c - dDepth - dFlangeThickness) > 0) Then: CreateCylinder dPipingDiameter, Abs(b - c - dDepth - dFlangeThickness), (b - c), 0, 0, 1, 0, 0, -1, 0, 0, "Cylinder2"
                    
            lIndex = 5
            CreateTorus dPipingDiameter, b, c, 0, 0
               
        Else 'Calculate the details if bend angle is other than 90
            
            dCyl1Length = a - c * Tan(dBendAngle / 2#)
            dCyl2Length = b - c * Tan(dBendAngle / 2#) - dDepth - dFlangeThickness
                        
            oCylStartPos.X = (dCyl2Length + dFlangeThickness + dDepth) + (c) * Sin(dBendAngle)
            oCylStartPos.y = c - (c) * Cos(dBendAngle)
                    
            lIndex = 1
            If dCyl1Length > 0 Then: CreateCylinder dPipingDiameter, dCyl1Length, oCylStartPos.X, oCylStartPos.y, 0, Cos(dBendAngle), Sin(dBendAngle), 0, Cos(dBendAngle), Sin(dBendAngle), 0, "Cylinder"
               
            lIndex = 2
            If (dCyl2Length > 0) Then: CreateCylinder dPipingDiameter, dCyl2Length, (dDepth + dFlangeThickness), 0, 0, 1, 0, 0, 1, 0, 0, "Cylinder2"
            
                    
            lIndex = 5
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set ObjElbow = oGeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(m_outputColl.ResourceManager, _
                                 (dCyl2Length + dFlangeThickness + dDepth), c, 0, 0, 0, 1, 0, -1, 0, c, dPipingDiameter / 2, dBendAngle, True)
            m_outputColl.AddOutput "Tori", ObjElbow
        End If
        
      
        lIndex = 3
        If (dFlangeThickness > 0 And (a >= (dFlangeDiameter / 2))) Then: CreateCylinder dFlangeDiameter, Abs(dFlangeThickness), (dDepth + dFlangeThickness), 0, 0, 1, 0, 0, -1, 0, 0, "Cylinder3"
    
        lIndex = 4
        If (dDepth > 0) Then: CreateCylinder dRaisedFaceOrSocketDiameter, Abs(dDepth), (dDepth), 0, 0, 1, 0, 0, -1, 0, 0, "Cylinder4"
    
    Else
        'Constructing the graphics of elbow nozzle symbol by considering Depth and Offset values of nozzle. The following code will be triggered only if customer creates UseLatestElbowNozzleSymbol.txt
        'file in SharedContent\Data\Equipment\.. folder.
        If Abs(dBendAngle - (PI / 2)) < ANGULARTOLERANCE Then '  if bend angle is 90 deg
            lIndex = 1
            If (a - c) > 0 Then: CreateCylinder dPipingDiameter, Abs(a - c), b - (dDepth - dCptOffset), a, 0, 0, -1, 0, 0, -1, 0, "Cylinder"
            
            lIndex = 2
            If ((b - c - dFlangeThickness) > 0) Then: CreateCylinder dPipingDiameter, Abs(b - c - dFlangeThickness), b - c - (dDepth - dCptOffset), 0, 0, 1, 0, 0, -1, 0, 0, "Cylinder2"
                    
            lIndex = 5
            CreateTorus dPipingDiameter, b, c, dDepth, dCptOffset
                   
        Else 'Calculate the details if bend angle is other than 90
            dCyl1Length = a - c * Tan(dBendAngle / 2#)
            dCyl2Length = b - c * Tan(dBendAngle / 2#) - dFlangeThickness
            
           
            oCylStartPos.X = (dCyl2Length + dFlangeThickness - (dDepth - dCptOffset)) + (c) * Sin(dBendAngle)
            oCylStartPos.y = c - (c) * Cos(dBendAngle)
                    
            lIndex = 1
            If dCyl1Length > 0 Then: CreateCylinder dPipingDiameter, Abs(dCyl1Length), oCylStartPos.X, oCylStartPos.y, 0, Cos(dBendAngle), Sin(dBendAngle), 0, Cos(dBendAngle), Sin(dBendAngle), 0, "Cylinder"
               
            lIndex = 2
            If (dCyl2Length > 0) Then: CreateCylinder dPipingDiameter, Abs(dCyl2Length), (dFlangeThickness - (dDepth - dCptOffset)), 0, 0, 1, 0, 0, 1, 0, 0, "Cylinder2"
            
                    
            lIndex = 5
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set ObjElbow = oGeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(m_outputColl.ResourceManager, _
                                 b - c * Tan(dBendAngle / 2#) - (dDepth - dCptOffset), c, 0, 0, 0, 1, 0, -1, 0, c, dPipingDiameter / 2, dBendAngle, True)
            m_outputColl.AddOutput "Tori", ObjElbow
        End If
        
      
        lIndex = 3
        If (dFlangeThickness > 0 And (a >= (dFlangeDiameter / 2))) Then: CreateCylinder dFlangeDiameter, Abs(dFlangeThickness), (dFlangeThickness - dDepth + dCptOffset), 0, 0, 1, 0, 0, -1, 0, 0, "Cylinder3"
    
    End If

    Exit Sub
ErrorLabel:
    Select Case lIndex:
        Case 1: oErrors.Add E_FAIL, "GSCADNozzleVBsymb.ElbowNozzle", "PipingDiameter must be greater than zero", "PipingDiameter>0"
        Case 2: oErrors.Add E_FAIL, "GSCADNozzleVBsymb.ElbowNozzle", "PipingDiameter must be greater than zero", "PipingDiameter>0"
        Case 3: oErrors.Add E_FAIL, "GSCADNozzleVBsymb.ElbowNozzle", "FlangeDiameter must be greater than zero", "FlangeDiameter>0"
        Case 4: oErrors.Add E_FAIL, "GSCADNozzleVBsymb.ElbowNozzle", "RaisedFaceOrSocketDiameter must be greater than zero", "RaisedFaceOrSocketDiameter>0"
    End Select
    Err.Raise E_FAIL
End Sub

Private Sub CreateCylinder(CylinderDia As Double, CylinderWidth As Double, circleCenterX As Double, _
                          circleCenterY As Double, circleCenterZ As Double, circleNormalX As Double, _
                          circleNormalY As Double, circleNormalZ As Double, projVecX As Double, _
                          projVecY As Double, projVecZ As Double, strOutput As String)
   
   
    Dim circlePointVecX As Double, circlePointVecY As Double, circlePointVecZ As Double
    Dim axesRatio As Double
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oCircle As IngrGeom3D.Circle3d
    Dim oProjection As IngrGeom3D.Projection3d
    
    On Error GoTo ErrorHandler
    
    circlePointVecX = 0#
    circlePointVecY = CylinderDia * 0.5
    circlePointVecZ = 0#
    axesRatio = 1#
    
    '=====================================
    ' CONSTRUCTION OF Cylinder
    '=====================================
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                circleCenterX, circleCenterY, circleCenterZ, _
                                circleNormalX, circleNormalY, circleNormalZ, _
                                CylinderDia / 2)
    Set oProjection = oGeomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, oCircle, _
                                    projVecX, projVecY, projVecZ, CylinderWidth, True)
    
    m_outputColl.AddOutput strOutput, oProjection
    Set oProjection = Nothing
    
    Dim gelem As IJDObject
    'Delete ellipse which was created for creating projection.
    Set gelem = oCircle
    Debug.Assert Not gelem Is Nothing
    Set oCircle = Nothing
    gelem.Remove
    
    Set oGeomFactory = Nothing
     Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

Private Sub CreateTorus(diameter As Double, b As Double, c As Double, dDepth As Double, dOffset As Double)
    
    ' Insert your code for output 1(Body of Elbow)
    Dim objCircle    As IngrGeom3D.Circle3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim ObjElbow     As Object
    
    On Error GoTo ErrorHandler
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    oErrors.Clear

    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set ObjElbow = oGeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(m_outputColl.ResourceManager, _
                     b - c - dDepth + dOffset, c, 0, 0, 0, 1, 0, -1, 0, c, diameter / 2, (PI / 2#), True)
    ' Set the output
    m_outputColl.AddOutput "Tori", ObjElbow
    Set ObjElbow = Nothing
    Set oGeomFactory = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF Elbow Nozzle
'=========================================================================
Private Sub InsulateElbowNozzle(ByRef arrayOfInputs())
On Error GoTo ErrorHandler
    '--------------------------------------------------
    '   Variable Declarations and their explanations
    '--------------------------------------------------
    Dim dFd As Double 'Flange Diameter
    Dim da As Double 'Nozzle length
    Dim db As Double 'Bend length
    Dim dc As Double 'bend Radius
    Dim dFt As Double 'Flange Thickness
    Dim dDt As Double 'Flange Depth Thickness
    Dim dDia As Double 'Pipe diameter
    Dim dT As Double 'Insulation Thickness
    Dim dUntracableVal As Double
    Dim dRaisedFaceOrSocketDiameter As Double
    Dim dInsulationCriteria As Long
    Dim dBendAngle As Double
    'Initialize Variables from Array of Inputs
    dFd = arrayOfInputs(1)
    dFt = arrayOfInputs(2)
    dDia = arrayOfInputs(3)
    da = arrayOfInputs(4)
    dUntracableVal = arrayOfInputs(5) 'Not sure what is this value
    dDt = arrayOfInputs(6)
    db = arrayOfInputs(7) 'b = Bend Length 'These might be interchanged later.
    dc = arrayOfInputs(8) 'c = Bend Radius
    dRaisedFaceOrSocketDiameter = arrayOfInputs(9) 'we donot need this in the insulation aspect
    dInsulationCriteria = arrayOfInputs(10) 'Insulation Criteria
    dT = arrayOfInputs(11) 'Insulation Thickness
    dBendAngle = arrayOfInputs(12)

    If (dRaisedFaceOrSocketDiameter < TOLERANCE) Then: dRaisedFaceOrSocketDiameter = dDia
    
    'if there is no insulation thick ness then better we exit from here itself so that we
    'will remove all the existing outputs from the symbol machinery minimum of 0.1 mm insulation thickness
    'should exist
    If dT <= TOLERANCE Then
        CreateDummyOutput (dFt)
        Exit Sub
    End If
    'we donot insulate the nozzle if the insulation criteria is set to none
    If (dInsulationCriteria < 1) Then
        CreateDummyOutput (dFt)
        Exit Sub
    End If
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oLinestr As IngrGeom3D.LineString3d
    Dim dCornerpts(0 To 14) As Double
    Dim oBaseuInsulGeom As IngrGeom3D.Revolution3d
    Dim oNeckInsulGeom As IngrGeom3D.Revolution3d
    Dim oFlngInsulGeom As IngrGeom3D.Revolution3d
    Dim OInsulElbow As Object
    Dim dNeckLength As Double
    Dim dFlngCornerpts(0 To 20) As Double
    Dim dMinorRad As Double
    
    If m_bApplyDepthAndOffset = False Then
        'Construct the base insulation --> This is a surface of revolution and will exist only on certain
        'condition. i.e. Nozzle length > Bend radius
        If (da > dc) Then
            'Create the cross section of recatangle and revolve the same along
            
            dCornerpts(0) = db - (dDia / 2#)
            dCornerpts(1) = da
            dCornerpts(2) = 0#
            
            dCornerpts(3) = db - (dDia / 2#) - dT
            dCornerpts(4) = da
            dCornerpts(5) = 0#
            
            dCornerpts(6) = db - (dDia / 2#) - dT
            dCornerpts(7) = dc
            dCornerpts(8) = 0#
            
            dCornerpts(9) = db - (dDia / 2#)
            dCornerpts(10) = dc
            dCornerpts(11) = 0#
            
            dCornerpts(12) = dCornerpts(0)
            dCornerpts(13) = dCornerpts(1)
            dCornerpts(14) = dCornerpts(2)
            
            Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dCornerpts)
            'Create the surface of Revolution using the same line string around Y axis
            Set oBaseuInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                           0#, -1#, 0#, _
                                                                           db, dc, 0#, (PI * 2#), True)
            'Add the surface to out put of insulation aspect
            m_outputColl.AddOutput "BaseInsulation", oBaseuInsulGeom
        End If
        'Now Create the Neck Insulation
        'Create the Insulation on the neck only if neck exists. --> Bend length - Bend radius - Flange Depth Thkness -
        'Flange thickness if that value is greater than Tol then only create
        dNeckLength = db - dc - dDt - dFt
        If (dNeckLength > TOLERANCE) Then
            
            dCornerpts(0) = dDt + dFt
            dCornerpts(1) = dDia / 2#
            dCornerpts(2) = 0#
            
            dCornerpts(3) = dDt + dFt
            dCornerpts(4) = (dDia / 2#) + dT
            dCornerpts(5) = 0#
            
            dCornerpts(6) = db - dc
            dCornerpts(7) = (dDia / 2#) + dT
            dCornerpts(8) = 0#
            
            dCornerpts(9) = db - dc
            dCornerpts(10) = dDia / 2#
            dCornerpts(11) = 0#
            
            dCornerpts(12) = dCornerpts(0)
            dCornerpts(13) = dCornerpts(1)
            dCornerpts(14) = dCornerpts(2)
            
            'Construct the Line string with above points
            Set oLinestr = Nothing
            Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dCornerpts)
            'Create the surface of Revolution using the same line string around X axis
            Set oNeckInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                           1#, 0#, 0#, _
                                                                           0#, 0#, 0#, (PI * 2#), True)
            'Add the surface to out put of insulation aspect
            m_outputColl.AddOutput "NeckInsulation", oNeckInsulGeom
        End If
        
        'Insulate the Flange also if user chooses to insulate the flange also
        If (dInsulationCriteria > 1) Then
            'Here we need to check whether he had any flang for this.
            If (dFt > TOLERANCE) Then
                'here the curve is a complicated curve of 6 sides there by making our points to 7
                
                dFlngCornerpts(0) = dDt
                dFlngCornerpts(1) = dFd / 2#
                dFlngCornerpts(2) = 0#
                
                dFlngCornerpts(3) = dDt
                dFlngCornerpts(4) = (dFd / 2#) + dT
                dFlngCornerpts(5) = 0#
                
                dFlngCornerpts(6) = dDt + dFt + dT
                dFlngCornerpts(7) = (dFd / 2#) + dT
                dFlngCornerpts(8) = 0#
                
                dFlngCornerpts(9) = dDt + dFt + dT
                dFlngCornerpts(10) = (dDia / 2#) + dT
                dFlngCornerpts(11) = 0#
                
                dFlngCornerpts(12) = dDt + dFt
                dFlngCornerpts(13) = (dDia / 2#) + dT
                dFlngCornerpts(14) = 0#
                
                dFlngCornerpts(15) = dDt + dFt
                dFlngCornerpts(16) = (dFd / 2#)
                dFlngCornerpts(17) = 0#
                
                dFlngCornerpts(18) = dFlngCornerpts(0)
                dFlngCornerpts(19) = dFlngCornerpts(1)
                dFlngCornerpts(20) = dFlngCornerpts(2)
                
                'Construct the Line string with above points
                Set oLinestr = Nothing
                Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dFlngCornerpts)
                'Create the surface of Revolution using the same line string around X axis
                Set oFlngInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                           1#, 0#, 0#, _
                                                                           0#, 0#, 0#, (PI * 2#), True)
                'Add the surface to out put of insulation aspect
                m_outputColl.AddOutput "FlangeInsulation", oFlngInsulGeom
            End If
        End If
        'Finally Construct the Torus that belongs Insulation aspect
        'For Torus to exist we need to have bend length > 0.0 and bend Radius > 0.0
        If (dc > 0# And db > 0#) Then
            'Create the Torus here that is of Insulation aspect
            ' Axis Center = db - dc, dc, 0
            ' Axis direction = 0, 0, 1
            ' Origin = 0, -1, 0
            ' Major Radius = dc
            ' Minor Radius = dDia / 2# + InsulationThickness
            ' Sweep Angle = PI /2#
            dMinorRad = (dDia / 2#) + dT
            'make some checks to make a Valid Torus
            If (dc >= dMinorRad) Then
                'Now create the Torus
                Set OInsulElbow = oGeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(m_outputColl.ResourceManager, _
                     db - dc, dc, 0, 0, 0, 1, 0, -1, 0, dc, dMinorRad, PI / 2#, True)
                ' Set the output
                m_outputColl.AddOutput "ToriInsulation", OInsulElbow
            End If
        End If
    Else
        'Construct the base insulation --> This is a surface of revolution and will exist only on certain
        'condition. i.e. Nozzle length > Bend radius
        If (da > dc) Then
            'Create the cross section of recatangle and revolve the same along
            
            dCornerpts(0) = db - (dDt - dUntracableVal) - (dDia / 2#)
            dCornerpts(1) = da
            dCornerpts(2) = 0#
            
            dCornerpts(3) = db - (dDt - dUntracableVal) - (dDia / 2#) - dT
            dCornerpts(4) = da
            dCornerpts(5) = 0#
            
            dCornerpts(6) = db - (dDt - dUntracableVal) - (dDia / 2#) - dT
            dCornerpts(7) = dc
            dCornerpts(8) = 0#
    
            dCornerpts(9) = db - (dDt - dUntracableVal) - (dDia / 2#)
            dCornerpts(10) = dc
            dCornerpts(11) = 0#
            
            dCornerpts(12) = dCornerpts(0)
            dCornerpts(13) = dCornerpts(1)
            dCornerpts(14) = dCornerpts(2)
            
            Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dCornerpts)
            'Create the surface of Revolution using the same line string around Y axis
            Set oBaseuInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                           0#, -1#, 0#, _
                                                                           db - (dDt - dUntracableVal), dc, 0#, (PI * 2#), True)
            'Add the surface to out put of insulation aspect
            m_outputColl.AddOutput "BaseInsulation", oBaseuInsulGeom
        End If
        'Now Create the Neck Insulation
        'Create the Insulation on the neck only if neck exists. --> Bend length - Bend radius - Flange Depth Thkness -
        'Flange thickness if that value is greater than Tol then only create
        dNeckLength = db - dc - dFt
        If (dNeckLength > TOLERANCE) Then
            
            dCornerpts(0) = dFt - (dDt - dUntracableVal)
            dCornerpts(1) = dDia / 2#
            dCornerpts(2) = 0#
            
            dCornerpts(3) = dFt - (dDt - dUntracableVal)
            dCornerpts(4) = (dDia / 2#) + dT
            dCornerpts(5) = 0#
            
            dCornerpts(6) = db - dc - (dDt - dUntracableVal)
            dCornerpts(7) = (dDia / 2#) + dT
            dCornerpts(8) = 0#
            
            dCornerpts(9) = db - dc - (dDt - dUntracableVal)
            dCornerpts(10) = dDia / 2#
            dCornerpts(11) = 0#
            
            dCornerpts(12) = dCornerpts(0)
            dCornerpts(13) = dCornerpts(1)
            dCornerpts(14) = dCornerpts(2)
            
            'Construct the Line string with above points
            Set oLinestr = Nothing
            Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dCornerpts)
            'Create the surface of Revolution using the same line string around X axis
            Set oNeckInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                           1#, 0#, 0#, _
                                                                           0#, 0#, 0#, (PI * 2#), True)
            'Add the surface to out put of insulation aspect
            m_outputColl.AddOutput "NeckInsulation", oNeckInsulGeom
        End If
        
        'Insulate the Flange also if user chooses to insulate the flange also
        If (dInsulationCriteria > 1) Then
            'Here we need to check whether he had any flang for this.
            If (dFt > TOLERANCE) Then
                
                dFlngCornerpts(0) = -dDt + dUntracableVal
                dFlngCornerpts(1) = dFd / 2#
                dFlngCornerpts(2) = 0#
                
                dFlngCornerpts(3) = -dDt + dUntracableVal
                dFlngCornerpts(4) = (dFd / 2#) + dT
                dFlngCornerpts(5) = 0#
                
                dFlngCornerpts(6) = dFt - dDt + dUntracableVal + dT
                dFlngCornerpts(7) = (dFd / 2#) + dT
                dFlngCornerpts(8) = 0#
                
                dFlngCornerpts(9) = dFt - dDt + dUntracableVal + dT
                dFlngCornerpts(10) = (dDia / 2#) + dT
                dFlngCornerpts(11) = 0#
    
                dFlngCornerpts(12) = dFt - dDt + dUntracableVal
                dFlngCornerpts(13) = (dDia / 2#) + dT
                dFlngCornerpts(14) = 0#
                
                dFlngCornerpts(15) = dFt - dDt + dUntracableVal
                dFlngCornerpts(16) = (dFd / 2#)
                dFlngCornerpts(17) = 0#
                
                dFlngCornerpts(18) = dFlngCornerpts(0)
                dFlngCornerpts(19) = dFlngCornerpts(1)
                dFlngCornerpts(20) = dFlngCornerpts(2)
                
                'Construct the Line string with above points
                Set oLinestr = Nothing
                Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dFlngCornerpts)
                'Create the surface of Revolution using the same line string around X axis
                Set oFlngInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                           1#, 0#, 0#, _
                                                                           0#, 0#, 0#, (PI * 2#), True)
                'Add the surface to out put of insulation aspect
                m_outputColl.AddOutput "FlangeInsulation", oFlngInsulGeom
            End If
        End If
        'Finally Construct the Torus that belongs Insulation aspect
        'For Torus to exist we need to have bend length > 0.0 and bend Radius > 0.0
        If (dc > 0# And db > 0#) Then
            'Create the Torus here that is of Insulation aspect
            ' Axis Center = db - dc, dc, 0
            ' Axis direction = 0, 0, 1
            ' Origin = 0, -1, 0
            ' Major Radius = dc
            ' Minor Radius = dDia / 2# + InsulationThickness
            ' Sweep Angle = PI /2#
            dMinorRad = (dDia / 2#) + dT
            'make some checks to make a Valid Torus
            If (dc >= dMinorRad) Then
                'Now create the Torus
                Set OInsulElbow = oGeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(m_outputColl.ResourceManager, _
                     db - dc - (dDt - dUntracableVal), dc, 0, 0, 0, 1, 0, -1, 0, dc, dMinorRad, PI / 2#, True)
                ' Set the output
                m_outputColl.AddOutput "ToriInsulation", OInsulElbow
            End If
        End If
    End If
Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

Private Sub CreateDummyOutput(ByVal dLoc As Double)
On Error GoTo ErrorHandler
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim oPoint As IngrGeom3D.Point3d
    Set oPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, dLoc, 0#, 0#)
    'Add the output to output Collection
    m_outputColl.AddOutput "BaseInsulation", oPoint
Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub
Private Sub InsulateElbowNozzleForBendAngeleOtherThan90(ByRef arrayOfInputs) 'If bend angle is otherthan 90
On Error GoTo ErrorHandler
    '--------------------------------------------------
    '   Variable Declarations and their explanations
    '--------------------------------------------------
    Dim dFd As Double 'Flange Diameter
    Dim da As Double 'Nozzle length
    Dim db As Double 'Bend length
    Dim dc As Double 'bend Radius
    Dim dFt As Double 'Flange Thickness
    Dim dDt As Double 'Flange Depth Thickness
    Dim dDia As Double 'Pipe diameter
    Dim dT As Double 'Insulation Thickness
    Dim dUntracableVal As Double
    Dim dRaisedFaceOrSocketDiameter As Double
    Dim dInsulationCriteria As Long
    Dim dBendAngle As Double
    'Initialize Variables from Array of Inputs
    dFd = arrayOfInputs(1)
    dFt = arrayOfInputs(2)
    dDia = arrayOfInputs(3)
    da = arrayOfInputs(4)
    dUntracableVal = arrayOfInputs(5) 'Not sure what is this value
    dDt = arrayOfInputs(6)
    db = arrayOfInputs(7) 'b = Bend Length 'These might be interchanged later.
    dc = arrayOfInputs(8) 'c = Bend Radius
    dRaisedFaceOrSocketDiameter = arrayOfInputs(9) 'we donot need this in the insulation aspect
    dInsulationCriteria = arrayOfInputs(10) 'Insulation Criteria
    dT = arrayOfInputs(11) 'Insulation Thickness
    dBendAngle = arrayOfInputs(12)
    
    If (dRaisedFaceOrSocketDiameter < TOLERANCE) Then: dRaisedFaceOrSocketDiameter = dDia
    
    
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oLinestr As IngrGeom3D.LineString3d
    Dim dCornerpts(0 To 14) As Double
    Dim oCylStartPos As New DPosition
    Dim dCyl1Length As Double, dCyl2Length As Double
    Dim oBaseuInsulGeom As IngrGeom3D.Revolution3d
    Dim oNeckInsulGeom As IngrGeom3D.Revolution3d
    Dim dFlngCornerpts(0 To 20) As Double
    Dim oFlngInsulGeom As IngrGeom3D.Revolution3d
    Dim OInsulElbow As Object
    Dim dMinorRad As Double
    
    If m_bApplyDepthAndOffset = False Then
        dCyl1Length = (da - dc * Tan((dBendAngle) / 2)) ' / Sin(dBendAngle)
        dCyl2Length = db - dc * Tan((dBendAngle) / 2) - dDt - dFt
        
        'if there is no insulation thick ness then better we exit from here itself so that we
        'will remove all the existing outputs from the symbol machinery minimum of 0.1 mm insulation thickness
        'should exist
        If dT <= TOLERANCE Then
            CreateDummyOutput (dFt)
            Exit Sub
        End If
        'we donot insulate the nozzle if the insulation criteria is set to none
        If (dInsulationCriteria < 1) Then
            CreateDummyOutput (dFt)
            Exit Sub
        End If
        
        
          oCylStartPos.X = (dCyl2Length + dFt + dDt) + (dc) * Sin(dBendAngle)
          oCylStartPos.y = dc - (dc) * Cos(dBendAngle)
          'Construct the base insulation --> This is a surface of revolution and will exist only on certain
          'condition. i.e. Nozzle length > Bend radius
          If (dCyl1Length > 0) Then
              'Create the cross section of recatangle and revolve the same along
              
              dCornerpts(0) = oCylStartPos.X + dCyl1Length * Cos(dBendAngle) - (dDia / 2) * Sin(dBendAngle) 'db - (dDia / 2#)
              dCornerpts(1) = oCylStartPos.y + dCyl1Length * Sin(dBendAngle) + (dDia / 2) * Cos(dBendAngle) 'da
              dCornerpts(2) = 0#
              
              dCornerpts(3) = oCylStartPos.X + dCyl1Length * Cos(dBendAngle) - (dDia / 2) * Sin(dBendAngle) - dT * Sin(dBendAngle) 'db - (dDia / 2#) - dT
              dCornerpts(4) = oCylStartPos.y + dCyl1Length * Sin(dBendAngle) + (dDia / 2) * Cos(dBendAngle) + dT * Cos(dBendAngle) 'da
              dCornerpts(5) = 0#
              
              dCornerpts(6) = oCylStartPos.X - (dDia / 2) * Sin(dBendAngle) - dT * Sin(dBendAngle) 'db - (dDia / 2#) - dT
              dCornerpts(7) = oCylStartPos.y + (dDia / 2) * Cos(dBendAngle) + dT * Cos(dBendAngle) 'dc
              dCornerpts(8) = 0#
              
              dCornerpts(9) = oCylStartPos.X - (dDia / 2) * Sin(dBendAngle) 'db - (dDia / 2#)
              dCornerpts(10) = oCylStartPos.y + (dDia / 2) * Cos(dBendAngle) 'dc
              dCornerpts(11) = 0#
              
              dCornerpts(12) = dCornerpts(0)
              dCornerpts(13) = dCornerpts(1)
              dCornerpts(14) = dCornerpts(2)
              
              Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dCornerpts)
              'Create the surface of Revolution using the same line string around Y axis
              Set oBaseuInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                             Cos(dBendAngle), Sin(dBendAngle), 0#, _
                                                                             oCylStartPos.X, oCylStartPos.y, 0#, (PI * 2#), True)
              'Add the surface to out put of insulation aspect
              m_outputColl.AddOutput "BaseInsulation", oBaseuInsulGeom
          End If
          'Now Create the Neck Insulation
          'Create the Insulation on the neck only if neck exists. --> Bend length - Bend radius - Flange Depth Thkness -
          'Flange thickness if that value is greater than Tol then only create
          
          If (dCyl2Length > 0) Then
              
              dCornerpts(0) = dDt + dFt
              dCornerpts(1) = dDia / 2#
              dCornerpts(2) = 0#
              
              dCornerpts(3) = dDt + dFt
              dCornerpts(4) = (dDia / 2#) + dT
              dCornerpts(5) = 0#
              
              dCornerpts(6) = dCyl2Length + dDt + dFt 'db - dc
              dCornerpts(7) = (dDia / 2#) + dT
              dCornerpts(8) = 0#
              
              dCornerpts(9) = dCyl2Length + dDt + dFt 'db - dc
              dCornerpts(10) = dDia / 2#
              dCornerpts(11) = 0#
              
              dCornerpts(12) = dCornerpts(0)
              dCornerpts(13) = dCornerpts(1)
              dCornerpts(14) = dCornerpts(2)
              
              'Construct the Line string with above points
              Set oLinestr = Nothing
              Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dCornerpts)
              'Create the surface of Revolution using the same line string around X axis
              Set oNeckInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                             1#, 0#, 0#, _
                                                                             0#, 0#, 0#, (PI * 2#), True)
              'Add the surface to out put of insulation aspect
              m_outputColl.AddOutput "NeckInsulation", oNeckInsulGeom
          End If
          
            'Insulate the Flange also if user chooses to insulate the flange also
          If (dInsulationCriteria > 1) Then
                'Here we need to check whether he had any flang for this.
                If (dFt > TOLERANCE) Then
                    'here the curve is a complicated curve of 6 sides there by making our points to 7
                    
                    dFlngCornerpts(0) = dDt
                    dFlngCornerpts(1) = dFd / 2#
                    dFlngCornerpts(2) = 0#
                    
                    dFlngCornerpts(3) = dDt
                    dFlngCornerpts(4) = (dFd / 2#) + dT
                    dFlngCornerpts(5) = 0#
                    
                    dFlngCornerpts(6) = dDt + dFt + dT
                    dFlngCornerpts(7) = (dFd / 2#) + dT
                    dFlngCornerpts(8) = 0#
                    
                    dFlngCornerpts(9) = dDt + dFt + dT
                    dFlngCornerpts(10) = (dDia / 2#) + dT
                    dFlngCornerpts(11) = 0#
                    
                    dFlngCornerpts(12) = dDt + dFt
                    dFlngCornerpts(13) = (dDia / 2#) + dT
                    dFlngCornerpts(14) = 0#
                    
                    dFlngCornerpts(15) = dDt + dFt
                    dFlngCornerpts(16) = (dFd / 2#)
                    dFlngCornerpts(17) = 0#
                    
                    dFlngCornerpts(18) = dFlngCornerpts(0)
                    dFlngCornerpts(19) = dFlngCornerpts(1)
                    dFlngCornerpts(20) = dFlngCornerpts(2)
                    
                    'Construct the Line string with above points
                    Set oLinestr = Nothing
                    Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dFlngCornerpts)
                    'Create the surface of Revolution using the same line string around X axis
                    Set oFlngInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                               1#, 0#, 0#, _
                                                                               0#, 0#, 0#, (PI * 2#), True)
                    'Add the surface to out put of insulation aspect
                    m_outputColl.AddOutput "FlangeInsulation", oFlngInsulGeom
                End If
          End If
          'Finally Construct the Torus that belongs Insulation aspect
          'For Torus to exist we need to have bend length > 0.0 and bend Radius > 0.0
          If (dc > 0# And db > 0#) Then

              dMinorRad = (dDia / 2#) + dT
              'make some checks to make a Valid Torus
              If (dc >= dMinorRad) Then
                  'Now create the Torus
                  Set OInsulElbow = oGeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(m_outputColl.ResourceManager, _
                       (dCyl2Length + dDt + dFt), dc, 0, 0, 0, 1, 0, -1, 0, dc, dMinorRad, dBendAngle, True)
                  ' Set the output
                  m_outputColl.AddOutput "ToriInsulation", OInsulElbow
              End If
          End If
    Else
        dCyl1Length = (da - dc * Tan((dBendAngle) / 2)) ' / Sin(dBendAngle)
        dCyl2Length = db - dc * Tan((dBendAngle) / 2) - dFt
        
        'if there is no insulation thick ness then better we exit from here itself so that we
        'will remove all the existing outputs from the symbol machinery minimum of 0.1 mm insulation thickness
        'should exist
        If dT <= TOLERANCE Then
            CreateDummyOutput (dFt)
            Exit Sub
        End If
        'we donot insulate the nozzle if the insulation criteria is set to none
        If (dInsulationCriteria < 1) Then
            CreateDummyOutput (dFt)
            Exit Sub
        End If
        
        
          oCylStartPos.X = (dCyl2Length + dFt - (dDt - dUntracableVal)) + (dc) * Sin(dBendAngle)
          oCylStartPos.y = dc - (dc) * Cos(dBendAngle)
          'Construct the base insulation --> This is a surface of revolution and will exist only on certain
          'condition. i.e. Nozzle length > Bend radius
          If (dCyl1Length > 0) Then
              'Create the cross section of recatangle and revolve the same along
              
              dCornerpts(0) = oCylStartPos.X + dCyl1Length * Cos(dBendAngle) - (dDia / 2) * Sin(dBendAngle) 'db - (dDia / 2#)
              dCornerpts(1) = oCylStartPos.y + dCyl1Length * Sin(dBendAngle) + (dDia / 2) * Cos(dBendAngle) 'da
              dCornerpts(2) = 0#
              
              dCornerpts(3) = oCylStartPos.X + dCyl1Length * Cos(dBendAngle) - (dDia / 2) * Sin(dBendAngle) - dT * Sin(dBendAngle) 'db - (dDia / 2#) - dT
              dCornerpts(4) = oCylStartPos.y + dCyl1Length * Sin(dBendAngle) + (dDia / 2) * Cos(dBendAngle) + dT * Cos(dBendAngle) 'da
              dCornerpts(5) = 0#
              
              dCornerpts(6) = oCylStartPos.X - (dDia / 2) * Sin(dBendAngle) - dT * Sin(dBendAngle) 'db - (dDia / 2#) - dT
              dCornerpts(7) = oCylStartPos.y + (dDia / 2) * Cos(dBendAngle) + dT * Cos(dBendAngle) 'dc
              dCornerpts(8) = 0#
              
              dCornerpts(9) = oCylStartPos.X - (dDia / 2) * Sin(dBendAngle) 'db - (dDia / 2#)
              dCornerpts(10) = oCylStartPos.y + (dDia / 2) * Cos(dBendAngle) 'dc
              dCornerpts(11) = 0#
              
              dCornerpts(12) = dCornerpts(0)
              dCornerpts(13) = dCornerpts(1)
              dCornerpts(14) = dCornerpts(2)
              
              Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dCornerpts)
              'Create the surface of Revolution using the same line string around Y axis
              Set oBaseuInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                             Cos(dBendAngle), Sin(dBendAngle), 0#, _
                                                                             oCylStartPos.X, oCylStartPos.y, 0#, (PI * 2#), True)
              'Add the surface to out put of insulation aspect
              m_outputColl.AddOutput "BaseInsulation", oBaseuInsulGeom
          End If
          'Now Create the Neck Insulation
          'Create the Insulation on the neck only if neck exists. --> Bend length - Bend radius - Flange Depth Thkness -
          'Flange thickness if that value is greater than Tol then only create
          
          If (dCyl2Length > 0) Then
              
              dCornerpts(0) = dFt - (dDt - dUntracableVal)
              dCornerpts(1) = dDia / 2#
              dCornerpts(2) = 0#
              
              dCornerpts(3) = dFt - (dDt - dUntracableVal)
              dCornerpts(4) = (dDia / 2#) + dT
              dCornerpts(5) = 0#
              
              dCornerpts(6) = dCyl2Length + dFt - (dDt - dUntracableVal) 'db - dc
              dCornerpts(7) = (dDia / 2#) + dT
              dCornerpts(8) = 0#
              
              dCornerpts(9) = dCyl2Length + dFt - (dDt - dUntracableVal) 'db - dc
              dCornerpts(10) = dDia / 2#
              dCornerpts(11) = 0#
              
              dCornerpts(12) = dCornerpts(0)
              dCornerpts(13) = dCornerpts(1)
              dCornerpts(14) = dCornerpts(2)
              
              'Construct the Line string with above points
              Set oLinestr = Nothing
              Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dCornerpts)
              'Create the surface of Revolution using the same line string around X axis
              Set oNeckInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                             1#, 0#, 0#, _
                                                                             0#, 0#, 0#, (PI * 2#), True)
              'Add the surface to out put of insulation aspect
              m_outputColl.AddOutput "NeckInsulation", oNeckInsulGeom
          End If
          
            'Insulate the Flange also if user chooses to insulate the flange also
          If (dInsulationCriteria > 1) Then
                'Here we need to check whether he had any flang for this.
                If (dFt > TOLERANCE) Then
                    'here the curve is a complicated curve of 6 sides there by making our points to 7
                    
                    dFlngCornerpts(0) = -dDt + dUntracableVal
                    dFlngCornerpts(1) = dFd / 2#
                    dFlngCornerpts(2) = 0#
                    
                    dFlngCornerpts(3) = -dDt + dUntracableVal
                    dFlngCornerpts(4) = (dFd / 2#) + dT
                    dFlngCornerpts(5) = 0#
                    
                    dFlngCornerpts(6) = dFt - dDt + dUntracableVal + dT
                    dFlngCornerpts(7) = (dFd / 2#) + dT
                    dFlngCornerpts(8) = 0#
                    
                    dFlngCornerpts(9) = dFt - dDt + dUntracableVal + dT
                    dFlngCornerpts(10) = (dDia / 2#) + dT
                    dFlngCornerpts(11) = 0#
        
                    dFlngCornerpts(12) = dFt - dDt + dUntracableVal
                    dFlngCornerpts(13) = (dDia / 2#) + dT
                    dFlngCornerpts(14) = 0#
                    
                    dFlngCornerpts(15) = dFt - dDt + dUntracableVal
                    dFlngCornerpts(16) = (dFd / 2#)
                    dFlngCornerpts(17) = 0#
                    
                    dFlngCornerpts(18) = dFlngCornerpts(0)
                    dFlngCornerpts(19) = dFlngCornerpts(1)
                    dFlngCornerpts(20) = dFlngCornerpts(2)
                    
                    'Construct the Line string with above points
                    Set oLinestr = Nothing
                    Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dFlngCornerpts)
                    'Create the surface of Revolution using the same line string around X axis
                    Set oFlngInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                               1#, 0#, 0#, _
                                                                               0#, 0#, 0#, (PI * 2#), True)
                    'Add the surface to out put of insulation aspect
                    m_outputColl.AddOutput "FlangeInsulation", oFlngInsulGeom
                End If
          End If
          'Finally Construct the Torus that belongs Insulation aspect
          'For Torus to exist we need to have bend length > 0.0 and bend Radius > 0.0
          If (dc > 0# And db > 0#) Then
              
              dMinorRad = (dDia / 2#) + dT
              'make some checks to make a Valid Torus
              If (dc >= dMinorRad) Then
                  'Now create the Torus
                  Set OInsulElbow = oGeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(m_outputColl.ResourceManager, _
                       db - dc * Tan(dBendAngle / 2#) - (dDt - dUntracableVal), dc, 0, 0, 0, 1, 0, -1, 0, dc, dMinorRad, dBendAngle, True)
                  ' Set the output
                  m_outputColl.AddOutput "ToriInsulation", OInsulElbow
              End If
          End If
    End If
Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

Private Function CheckIfDepthandoffsetFileExists() As Boolean
    Dim strTxtFile As String
    Dim strResults As String
    Dim oJContext As IJContext
    Set oJContext = GetJContext()
    Dim sSymbolSharePath  As String
    sSymbolSharePath = oJContext.GetVariable("OLE_SERVER")
    strTxtFile = sSymbolSharePath + "\Data\Equipment\ReactToElbowNzlDepthAndOffset.txt"
    Set oJContext = Nothing
    strResults = Dir$(strTxtFile)
    If strResults = "" Then
        CheckIfDepthandoffsetFileExists = False
    Else
        CheckIfDepthandoffsetFileExists = True
    End If
End Function
