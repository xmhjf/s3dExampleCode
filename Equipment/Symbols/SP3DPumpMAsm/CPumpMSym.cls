VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPumpMSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:13 AM"
'******************************************************************
' Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'File
'    CPumpMSym.cls
'
'Author
'    Nabil Nicolas
'
'Description
'
'
'   30 Jan 2003     Nabil Nicolas           Created a mirrored copy of Pump.Pumpservices
'                                                           and mirrored the pump discharge.
'   16 Feb 2003     Nabil Nicolas           Apply transform translate along Y axis by -Ymax for mirror about the symbol XZ plane.
'   29 Jan 2003     BG                    TR# 39362 Modified the control point diameter to zero to avoid interference.
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection


Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DPumpMAsm.CPumpMSym"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  Set oSymbolFactory = Nothing
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DPumpMAsm.CPumpMSym"
  oSymbolDefinition.CodeBase = CodeBase

  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId
  
  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  Set oSymbolDefinition = Nothing
  
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

    ' Feed Pump Definition
    ' Inputs:
    '          1. "Part"  ( Catalog part )
    '          2. "ImpellerWidth"
    '          3. "ImpellerDiameter"
    '          4. "ShaftLength"
    '          5. "ShaftDiameter"
    '          6. "MotorWidth"
    '          7. "MotorDiameter"
    '          8. "SupportLength"( Length of support plate )
    '          9. "SupportHeight" ( Height of support plate )
    '         10. "SupportThickness" ( Thickness of support plate )
    '         11. "BasePlaceThickness" (Thickness of BASE PLATE)
    ' Representations :
    '           Physical
  
    On Error GoTo ErrorHandler

    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
    
    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    ' Create a new input by new operator
    Dim Inputs(1 To 10) As IMSSymbolEntities.DInput
  
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue

    'Define inputs for Impeller, Shaft and Motor
    Dim Index As Integer
    For Index = 1 To 3
    
        Set Inputs(Index * 2 - 1) = New IMSSymbolEntities.DInput
        Inputs(Index * 2 - 1).Description = "Length" + Str(Index) + " of the Pump"
        Inputs(Index * 2 - 1).Properties = igINPUT_IS_A_PARAMETER
        
        Set Inputs(Index * 2) = New IMSSymbolEntities.DInput
        Inputs(Index * 2).Description = "Diameter" + Str(Index) + " of the Pump"
        Inputs(Index * 2).Properties = igINPUT_IS_A_PARAMETER
        
        If (Index = 1) Then
            PC.UomValue = 0.3
            Inputs(Index * 2 - 1).Name = "ImpellerWidth"
        Else
            If (Index = 2) Then
                PC.UomValue = 0.2
                Inputs(Index * 2 - 1).Name = "ShaftLength"
            Else
                PC.UomValue = 0.65
                Inputs(Index * 2 - 1).Name = "MotorWidth"
            End If
        End If
        Inputs(Index * 2 - 1).DefaultParameterValue = PC
        
        If (Index = 1) Then
            PC.UomValue = 0.8
            Inputs(Index * 2).Name = "ImpellerDiameter"
        Else
            If (Index = 2) Then
                PC.UomValue = 0.08
                Inputs(Index * 2).Name = "ShaftDiameter"
            Else
                PC.UomValue = 0.4
                Inputs(Index * 2).Name = "MotorDiameter"
            End If
        End If
        Inputs(Index * 2).DefaultParameterValue = PC
    Next
  
    'Define inputs for impeller and motor supports
    For Index = 1 To 3
        Set Inputs(Index + 6) = New IMSSymbolEntities.DInput
        Inputs(Index + 6).Description = "Dim" + Str(Index) + " of the Block of Pump"
        Inputs(Index + 6).Properties = igINPUT_IS_A_PARAMETER
        
        If (Index = 1) Then
           PC.UomValue = 0.45
           Inputs(Index + 6).Name = "SupportLength"
        Else
           If (Index = 2) Then
               PC.UomValue = 0.25
               Inputs(Index + 6).Name = "SupportHeight"
           Else
               PC.UomValue = 0.15
               Inputs(Index + 6).Name = "SupportThickness"
           End If
        End If
        Inputs(Index + 6).DefaultParameterValue = PC
    Next

    'Define basePlateThickness input
    Set Inputs(10) = New IMSSymbolEntities.DInput
    Inputs(10).Name = "BasePlateThickness"
    Inputs(10).Description = "Thickness of Base Plate"
    Inputs(10).Properties = igINPUT_IS_A_PARAMETER

    PC.UomValue = 0.1
    Inputs(10).DefaultParameterValue = PC
      


    For Index = 1 To 10
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next



    'Define the outputs
    Dim O1 As IMSSymbolEntities.DOutput
    Set O1 = New IMSSymbolEntities.DOutput

    O1.Name = "ImpellerCylinder"
    O1.Description = "Cylinder representing impeller of Pump"
    O1.Properties = 0

    Dim O2 As IMSSymbolEntities.DOutput
    Set O2 = New IMSSymbolEntities.DOutput
  
    O2.Name = "ShaftCylinder"
    O2.Description = "Cylinder representing shaft of Pump"
    O2.Properties = 0
  
    Dim O3 As IMSSymbolEntities.DOutput
    Set O3 = New IMSSymbolEntities.DOutput
  
    O3.Name = "MotorCylinder"
    O3.Description = "Cylinder representing motor of Pump"
    O3.Properties = 0
   
    
    Dim O4 As IMSSymbolEntities.DOutput
    Set O4 = New IMSSymbolEntities.DOutput
    
    O4.Name = "EqpFoundationPort"
    O4.Description = "EqpFoundationPort of Pump"
    O4.Properties = 0
    
    Dim O5 As IMSSymbolEntities.DOutput
    Set O5 = New IMSSymbolEntities.DOutput
    
    O5.Name = "PumpControlPoint"
    O5.Description = "Control Point of Pump Skid"
    O5.Properties = 0
    
    Dim B1P(0 To 5) As IMSSymbolEntities.DOutput
    For Index = 0 To 5
        Set B1P(Index) = New IMSSymbolEntities.DOutput
        B1P(Index).Name = "B1P" + Str(Index + 1)
        B1P(Index).Description = "B1P" + Str(Index + 1) + " of Pump"
        B1P(Index).Properties = 0
    Next

    B1P(0).Name = "B1P1"
    B1P(1).Name = "B1P2"
    B1P(2).Name = "B1P3"
    B1P(3).Name = "B1P4"
    B1P(4).Name = "B1P5"
    B1P(5).Name = "B1P6"

    Dim Support1 As New IMSSymbolEntities.DOutput
    Support1.Name = "Support1"
    Support1.Description = "Support 1 of the Pump"
    Support1.Properties = 0
    
    Dim Support2 As New IMSSymbolEntities.DOutput
    Support2.Name = "Support2"
    Support2.Description = "Support 2 of the Pump"
    Support2.Properties = 0
  
    Dim locatePlane As New IMSSymbolEntities.DOutput
    locatePlane.Name = "DefaultSurface"
    locatePlane.Description = "Default Plane for Relations"
    locatePlane.Properties = 0

    'Define the representation "Symbolic"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation

    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of Pump"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical


    Dim IJDOutputs As IMSSymbolEntities.IJDOutputs
    Set IJDOutputs = rep1

    ' Set the output
    IJDOutputs.SetOutput O1
    Set O1 = Nothing
    IJDOutputs.SetOutput O2
    Set O2 = Nothing
    IJDOutputs.SetOutput O3
    Set O3 = Nothing
    IJDOutputs.SetOutput O4
    Set O4 = Nothing
    IJDOutputs.SetOutput O5
    Set O5 = Nothing

    For Index = 0 To 5
        IJDOutputs.SetOutput B1P(Index)
        Set B1P(Index) = Nothing
    Next

    IJDOutputs.SetOutput Support1
    Set Support1 = Nothing

    IJDOutputs.SetOutput Support2
    Set Support2 = Nothing
    
    IJDOutputs.SetOutput locatePlane
    Set locatePlane = Nothing

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
    PhysicalRepEval.ProgId = "SP3DPumpMAsm.CPumpMSym"

    'Set the evaluations for the Symbolic and Physical representations on the definition
    Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    Set PhysicalRepEval = Nothing
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
    End If

End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF PUMP
'=========================================================================
Private Sub Physical(ByRef arrayOfInputs())

    Dim Offset As Double
    Dim X1 As Double, Y1 As Double, Z1 As Double
    Dim X2 As Double, Y2 As Double, Z2 As Double
    Offset = 0.05  '0.02
    Dim impellerWidth As Double, impellerDia As Double
    Dim shaftWidth As Double, shaftDia As Double
    Dim motorWidth As Double, motorDia As Double
    Dim ii As Integer
    Dim Coords(1 To 6) As Double

    
    Dim supportLength As Double, supportHeight As Double, supportThk As Double, _
        basePlateThk As Double
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
   
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    
    'assign to meaningful variables from the input array
    impellerWidth = arrayOfInputs(2)
    impellerDia = arrayOfInputs(3)
    shaftWidth = arrayOfInputs(4)
    shaftDia = arrayOfInputs(5)
    motorWidth = arrayOfInputs(6)
    motorDia = arrayOfInputs(7)
    supportLength = arrayOfInputs(8)
    supportHeight = arrayOfInputs(9)
    supportThk = arrayOfInputs(10)
    basePlateThk = arrayOfInputs(11)

'=================
'BUILD BASE PLATE
'=================
    Dim Xmin As Double, Ymin As Double, Zmin As Double
    Dim Xmax As Double, Ymax As Double, Zmax As Double
    Dim Points(0 To 11) As Double

    Xmin = 0#
    Ymin = 0#
    Zmin = 0#
    Xmax = Offset + impellerWidth + shaftWidth + motorWidth + Offset
    Ymax = Offset + supportLength + Offset
    Zmax = basePlateThk
    
    'Define transform to translate the pump along + Ymax
    Dim m_transform As IngrGeom3D.IJDT4x4
    Dim iVector As IJDVector
    Set m_transform = New DT4x4
    m_transform.LoadIdentity
    
    'translate the output along y-axis
        'Ymax = Offset + supportLength + Offset
    Set iVector = New DVector
    
    iVector.x = 0
    iVector.y = -Ymax
    iVector.z = 0
    m_transform.Translate iVector
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim plane As IngrGeom3D.Plane3d

    Points(0) = Xmin
    Points(1) = Ymin
    Points(2) = Zmin
    Points(3) = Xmin
    Points(4) = Ymin
    Points(5) = Zmax
    Points(6) = Xmin
    Points(7) = Ymax
    Points(8) = Zmax
    Points(9) = Xmin
    Points(10) = Ymax
    Points(11) = Zmin
    Set plane = geomFactory.Planes3d.CreateByPoints(m_outputColl.ResourceManager, 4, Points)
    plane.Transform m_transform

    m_outputColl.AddOutput "B1P1", plane
    Set plane = Nothing

    Points(0) = Xmax
    Points(1) = Ymin
    Points(2) = Zmin
    Points(3) = Xmax
    Points(4) = Ymax
    Points(5) = Zmin
    Points(6) = Xmax
    Points(7) = Ymax
    Points(8) = Zmax
    Points(9) = Xmax
    Points(10) = Ymin
    Points(11) = Zmax
    Set plane = geomFactory.Planes3d.CreateByPoints(m_outputColl.ResourceManager, 4, Points)
    plane.Transform m_transform

    m_outputColl.AddOutput "B1P2", plane
    Set plane = Nothing

    Points(0) = Xmin
    Points(1) = Ymin
    Points(2) = Zmin
    Points(3) = Xmax
    Points(4) = Ymin
    Points(5) = Zmin
    Points(6) = Xmax
    Points(7) = Ymin
    Points(8) = Zmax
    Points(9) = Xmin
    Points(10) = Ymin
    Points(11) = Zmax
    Set plane = geomFactory.Planes3d.CreateByPoints(m_outputColl.ResourceManager, 4, Points)
    plane.Transform m_transform
 
    m_outputColl.AddOutput "B1P3", plane
    Set plane = Nothing
    
    Points(0) = Xmin
    Points(1) = Ymax
    Points(2) = Zmin
    Points(3) = Xmin
    Points(4) = Ymax
    Points(5) = Zmax
    Points(6) = Xmax
    Points(7) = Ymax
    Points(8) = Zmax
    Points(9) = Xmax
    Points(10) = Ymax
    Points(11) = Zmin
    Set plane = geomFactory.Planes3d.CreateByPoints(m_outputColl.ResourceManager, 4, Points)
    plane.Transform m_transform
    m_outputColl.AddOutput "B1P4", plane
    Set plane = Nothing

    Points(0) = Xmin
    Points(1) = Ymin
    Points(2) = Zmin
    Points(3) = Xmin
    Points(4) = Ymax
    Points(5) = Zmin
    Points(6) = Xmax
    Points(7) = Ymax
    Points(8) = Zmin
    Points(9) = Xmax
    Points(10) = Ymin
    Points(11) = Zmin
    Set plane = geomFactory.Planes3d.CreateByPoints(m_outputColl.ResourceManager, 4, Points)
    plane.Transform m_transform
    m_outputColl.AddOutput "B1P5", plane
    Set plane = Nothing

    Points(0) = Xmin
    Points(1) = Ymin
    Points(2) = Zmax
    Points(3) = Xmax
    Points(4) = Ymin
    Points(5) = Zmax
    Points(6) = Xmax
    Points(7) = Ymax
    Points(8) = Zmax
    Points(9) = Xmin
    Points(10) = Ymax
    Points(11) = Zmax
    Set plane = geomFactory.Planes3d.CreateByPoints(m_outputColl.ResourceManager, 4, Points)
    plane.Transform m_transform
    m_outputColl.AddOutput "B1P6", plane
    Set plane = Nothing
    

'=====================================
' CONSTRUCTION OF  Impeller Cylinder
'=====================================
    Dim circleCenterX As Double, circleCenterY As Double, circleCenterZ As Double
    circleCenterX = Offset
    circleCenterY = Offset + supportLength * 0.5
    circleCenterZ = basePlateThk + supportHeight + impellerDia * 0.5

    Dim circleNormalX As Double, circleNormalY As Double, circleNormalZ As Double
    circleNormalX = 1#
    circleNormalY = 0#
    circleNormalZ = 0#

    Dim circlePointVecX As Double, circlePointVecY As Double, circlePointVecZ As Double
    circlePointVecX = 0#
    circlePointVecY = impellerDia * 0.5
    circlePointVecZ = 0#


    Dim axesRatio As Double
    axesRatio = 1#

    Dim ellipse As IngrGeom3D.Ellipse3d
    Set ellipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                circleCenterX, circleCenterY, circleCenterZ, _
                                circleNormalX, circleNormalY, circleNormalZ, _
                                circlePointVecX, circlePointVecY, circlePointVecZ, _
                                axesRatio)

    Dim projVecX As Double, projVecY As Double, projVecZ As Double
    projVecX = 1#
    projVecY = 0#
    projVecZ = 0#
    Dim projection As IngrGeom3D.Projection3d
    Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, ellipse, _
                                    projVecX, projVecY, projVecZ, impellerWidth, True)
    
    projection.Transform m_transform

    m_outputColl.AddOutput "ImpellerCylinder", projection
    Set projection = Nothing
    
''''    'Delete ellipse which was created for creating projection.
''''    Dim gelem As IJDObject
''''    Set gelem = ellipse
''''    Debug.Assert Not gelem Is Nothing
''''    Set ellipse = Nothing
''''    gelem.Remove
''''

'=====================================
' CONSTRUCTION OF  Shaft Cylinder
'=====================================
    circleCenterX = circleCenterX + impellerWidth
    circlePointVecY = shaftDia * 0.5
    Set ellipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                circleCenterX, circleCenterY, circleCenterZ, _
                                circleNormalX, circleNormalY, circleNormalZ, _
                                circlePointVecX, circlePointVecY, circlePointVecZ, _
                                axesRatio)

    Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, ellipse, _
                                    projVecX, projVecY, projVecZ, shaftWidth, True)
    projection.Transform m_transform
    m_outputColl.AddOutput "ShaftCylinder", projection
    Set projection = Nothing
    
''''    'Delete ellipse which was created for creating projection.
''''    Set gelem = ellipse
''''    Debug.Assert Not gelem Is Nothing
''''    Set ellipse = Nothing
''''    gelem.Remove
''''
'=====================================
' CONSTRUCTION OF  Motor Cylinder
'=====================================
    circleCenterX = circleCenterX + shaftWidth
    circlePointVecY = motorDia * 0.5
    Set ellipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                circleCenterX, circleCenterY, circleCenterZ, _
                                circleNormalX, circleNormalY, circleNormalZ, _
                                circlePointVecX, circlePointVecY, circlePointVecZ, _
                                axesRatio)

    Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, ellipse, _
                                    projVecX, projVecY, projVecZ, motorWidth, True)
    projection.Transform m_transform
    m_outputColl.AddOutput "MotorCylinder", projection
    Set projection = Nothing

''''    'Delete ellipse which was created for creating projection.
''''    Set gelem = ellipse
''''    Debug.Assert Not gelem Is Nothing
''''    Set ellipse = Nothing
''''    gelem.Remove
''''
    Set geomFactory = Nothing

'=======================
'BUILD SUPPORT PLATE1
'=======================
    Coords(1) = Offset + impellerWidth * 0.5 - supportThk * 0.5
    Coords(2) = Offset - Ymax
    Coords(3) = basePlateThk
    Coords(4) = Coords(1) + supportThk
    Coords(5) = Coords(2) + supportLength
    Coords(6) = Coords(3) + supportHeight + impellerDia * 0.25

    ' Get or create the definition
    Dim defColl As IJDDefinitionCollection
    Set defColl = oSymbolFactory.DefinitionCollection(m_outputColl.ResourceManager)
    Dim boxDef As IJDSymbolDefinition
    Dim definitionParams As Variant
    definitionParams = ""
    Set boxDef = defColl.GetDefinitionByProgId(True, "Box.BoxServices", vbNullString, definitionParams)
    
    Dim Support1 As IMSSymbolEntities.DSymbol
    Set Support1 = oSymbolFactory.PlaceSymbol(boxDef, m_outputColl.ResourceManager)
'    Support1.Transform m_transform

    Dim newEnumArg2 As IJDEnumArgument
    Dim IJEditJDArg2 As IJDEditJDArgument
    Set newEnumArg2 = New DEnumArgument
    Set IJEditJDArg2 = newEnumArg2.IJDEditJDArgument

    Dim PC2 As IJDParameterContent
    Dim argument2 As IJDArgument
    For ii = 1 To 6
        Set PC2 = New DParameterContent
        Set argument2 = New DArgument

        PC2.UomValue = Coords(ii)
        PC2.Type = igValue
        PC2.UomType = 1
        ' Feed the Argument
        argument2.Index = ii
        argument2.Entity = PC2
        ' Add the argument to the arg collection
        IJEditJDArg2.SetArg argument2
        Set PC2 = Nothing
        Set argument2 = Nothing
    Next

    Support1.IJDValuesArg.SetValues newEnumArg2
    Dim IJDInputsArg2 As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg2 = Support1
    IJDInputsArg2.Update
    Set IJDInputsArg2 = Nothing
    Set IJEditJDArg2 = Nothing
    Set newEnumArg2 = Nothing

    m_outputColl.AddOutput "Support1", Support1

'=======================
'BUILD SUPPORT PLATE2
'=======================
    Coords(1) = Offset + impellerWidth + shaftWidth + _
                motorWidth * 0.5 - supportThk * 0.5
    Coords(2) = -Ymax + Offset + supportLength * 0.5 - motorDia * 0.6
    Coords(3) = basePlateThk
    Coords(4) = Coords(1) + supportThk
    Coords(5) = Coords(2) + motorDia * 1.2
    Coords(6) = Coords(3) + supportHeight + impellerDia * 0.5 - motorDia * 0.25

    
    Dim Support2 As IMSSymbolEntities.DSymbol
    Set Support2 = oSymbolFactory.PlaceSymbol(boxDef, m_outputColl.ResourceManager)

    Dim newEnumArg3 As IJDEnumArgument
    Dim IJEditJDArg3 As IJDEditJDArgument
    Set newEnumArg3 = New DEnumArgument
    Set IJEditJDArg3 = newEnumArg3.IJDEditJDArgument

    Dim PC3 As IJDParameterContent
    Dim argument3 As IJDArgument
    For ii = 1 To 6
        Set PC3 = New DParameterContent
        Set argument3 = New DArgument

        PC3.UomValue = Coords(ii)
        PC3.Type = igValue
        PC3.UomType = 1
        ' Feed the Argument
        argument3.Index = ii
        argument3.Entity = PC3
        ' Add the argument to the arg collection
        IJEditJDArg3.SetArg argument3
        Set PC3 = Nothing
        Set argument3 = Nothing
    Next

    Support2.IJDValuesArg.SetValues newEnumArg3
    Dim IJDInputsArg3 As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg3 = Support2
    IJDInputsArg3.Update
    Set IJDInputsArg3 = Nothing
    Set IJEditJDArg3 = Nothing
    Set newEnumArg3 = Nothing
'    Support2.Transform m_transform
    m_outputColl.AddOutput "Support2", Support2

    
'==========================================
'Construction of Equipment Foundatiopn Port
'==========================================
    Dim oEqpFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As New NozzlePHFactory

    Set oEqpFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPart, "PumpFoundationPort", _
                                                False, m_outputColl.ResourceManager)
                                                
    Call oEqpFoundationPort.PutCS(0#, -Ymax, 0#, 1#, 0#, 0#, 0#, 0#, -1#)
'    oEqpFoundationPort.Transform m_transform
    m_outputColl.AddOutput "EqpFoundationPort", oEqpFoundationPort
    Set oEqpFoundationPort = Nothing
    
'==========================================
'Construction of Pump Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_outputColl.ResourceManager, 0, 0, 0, 0, , , True)
    m_oControlPoint.SubType = cpProcessEquipment
    m_oControlPoint.Type = cpControlPoint

    m_outputColl.AddOutput "PumpControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    '============================================================================
    'BUILD DEFAULT SURFACE
    '============================================================================
    'This is constructed at the bottom of the base plate
    
    Xmin = 0#
    Ymin = 0#
    Zmin = 0#
    Xmax = Offset + impellerWidth + shaftWidth + motorWidth + Offset
    Ymax = Offset + supportLength + Offset
    Zmax = 0#
    
    Points(0) = Xmin
    Points(1) = Ymin
    Points(2) = Zmin
    Points(3) = Xmin
    Points(4) = Ymax
    Points(5) = Zmin
    Points(6) = Xmax
    Points(7) = Ymax
    Points(8) = Zmin
    Points(9) = Xmax
    Points(10) = Ymin
    Points(11) = Zmin
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set plane = geomFactory.Planes3d.CreateByPoints(m_outputColl.ResourceManager, 4, Points)
    plane.Transform m_transform
    m_outputColl.AddOutput "DefaultSurface", plane
    Set plane = Nothing
    Set geomFactory = Nothing
    
    Set boxDef = Nothing
    Set defColl = Nothing

    
    Set oPart = Nothing
    
    Set oSymbolFactory = Nothing
    
    Set iVector = Nothing
    Set m_transform = Nothing
    
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "Pump definition"
        Debug.Assert False
End Sub

'Following method will be removed very soon. This has been
'included to improve the performance. These lines will be removed as soon as
'symbol subsytem allows this to be removed from Symbol Definition .- Raju 05/05/99.

' CMCache custom method to cache the input argument into a parameter contend and the reverse conversion
'It is up to you to find a way to convert your reference data object to a parameter content
Public Sub CMCacheForPart(pInputCM As Object, bArgToCache As Boolean, pToConvert As Object, ByRef pOutput As Object)

 If bArgToCache Then
       
        'Need to convert the graphic input pToConvert into a Parameter ( pOutput)
        Dim oPC As IJDParameterContent
        Set oPC = New DParameterContent
        
        Dim oPart As IJDPart
        Set oPart = pToConvert
        'MsgBox "Partnumber :" & oPart.PartNumber
        
        ' Create a PC whose value is an identifier of the input
        ' Raju,
        ' the property Part_Number must be retrieved form from the pToConvert argument.
        ' I am hardcoding it
        '
        oPC.Type = igString
        oPC.String = oPart.PartNumber

        Set oPart = Nothing

        ' Always return this PC
        Set pOutput = oPC
        Set oPC = Nothing
 Else
        'Need to convert the cached Parameter pToConvert into your reference data object pOutput
        Dim oPCout As IJDParameterContent
        Dim oPCCache As IJDParameterContent
        Set oPCout = New DParameterContent
        
        ' Here there is three options
        ' o Return a parameter contents, that containts the part_number stored by the pToConvert argument
        '   In this case the edit command will have to retrieve the Part object when needed.
        '   Note : It is better to return a copy of the cached object instead the cached object itself.
        '          This allow to avoid the edition of the cached object.
        ' o Retrieve the Part_number and retrieve your part object with it.
        '   With this solution you can have assoc assertion while when this method is called you are in
        '   in a compute process.
        ' o Get the symbol (or equipment) from the pInputCM, query for the IJDReferencesArg interface,
        '   then get the argument at index 1 it is the SiteProxy.
        '   With this last method, the part has to be passed by reference to the symbol,
        '   but that is what you doing with your design.
        
        ' returning NULL means that the cache method is unable to resolve the cache.
        ' as of now anyway the part is passed as an input argument. this issue of caching
        ' has to be resolved in cycle2.
        
        Set pOutput = Nothing
        
        ' Here is the implementation of the option 1
'        Set oPCCache = pToConvert
'        oPCout.Type = oPCCache.Type
'        oPCout.String = oPCCache.String
'        Set pOutput = oPCout
'        Set oPCout = Nothing
        
        ' Here is the implementation of the option 2
'        Dim oRefDB As GSCADRefDataServices.IJDRefDBGlobal
'        Dim oPart As IJDPart

'        Set oRefDB = New GSCADRefDataServices.RefDBGlobal
'        Set oPart = oRefDB.GetCatalogPart("Storage Tanks", oPCCache.String)
'        Set pOutput = oRefDB
'        Set oRefDB = Nothing
'        Set oPart = Nothing
End If

End Sub

