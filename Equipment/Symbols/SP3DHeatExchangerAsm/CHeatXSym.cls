VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHeatXSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:00:09 AM"
'******************************************************************
' Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'File
'    CHeatXSym.cls
'
'Author
'    RM
'
'Description
'    Definition of Heat Exchanger Symbol
'
'Notes
'
'    <notes>
'
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  20.Dec.2007     VRK TR-CP·118167  Wrong dimensions for Support Height of the Heat Exchanger HEXCH 001A-E
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection
Private Const E_FAIL = -2147467259
Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    FeedHeatXDefinition pSymbolDefinition

End Sub

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHeatXAsm.CHeatXSym"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
  Set oSymbolFactory = Nothing
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHeatXAsm.CHeatXSym"
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
    Err.Raise E_FAIL
End Function



Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
On Error GoTo ErrorHandler
    Set m_outputColl = outputcoll
    If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
    End If
Exit Sub
ErrorHandler:
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Err.Raise E_FAIL, "IJDUserSymbolServices_InvokeRepresentation", Err.Description, Err.HelpFile, Err.HelpContext
End Sub


' Feed HeatX Definition
' Inputs:
'           1. "Part"  ( Catalog Part )
'           2. "Length"
'           3. "Diameter"
'           4. "SupportLength"( Length of support plate )
'           5. "SupportHeight" ( Height of support plate )
'           6. "SupportThick" ( Thickness of support plate )
'           7. "NozzleLength" ( Length of the nozzles )
' The representation "Physical" is set to the definition

Private Sub FeedHeatXDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
  
  On Error GoTo ErrorHandler
  
  ' Create a new input by new operator
  Dim Inputs(1 To 5) As IMSSymbolEntities.IJDInput
  
  ' Create a defaultValue
  Dim PC As IMSSymbolEntities.DParameterContent
  Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
  PC.Type = igValue

  Dim Index As Integer

  Dim InputsIf As IMSSymbolEntities.IJDInputs
  Set InputsIf = pSymbolDefinition

  Set Inputs(1) = New IMSSymbolEntities.DInput
  Inputs(1).Name = "VesselLength"
  Inputs(1).Description = "Length of the HeatX"
  Inputs(1).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 3.5
  Inputs(1).DefaultParameterValue = PC
    
  Set Inputs(2) = New IMSSymbolEntities.DInput
  Inputs(2).Name = "VesselDiameter"
  Inputs(2).Description = "Diameter of the HeatX"
  Inputs(2).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.7
  Inputs(2).DefaultParameterValue = PC
  
  For Index = 1 To 3
     Set Inputs(Index + 2) = New IMSSymbolEntities.DInput
     Inputs(Index + 2).Description = "Dim" + Str(Index) + " of the Block of HeatX"
     Inputs(Index + 2).Properties = igINPUT_IS_A_PARAMETER
     
     If (Index = 1) Then
        PC.UomValue = 0.6
        Inputs(Index + 2).Name = "SupportLength"
     Else
        If (Index = 2) Then
            PC.UomValue = 0.6
            Inputs(Index + 2).Name = "SupportHeight"
        Else
            PC.UomValue = 0.15
            Inputs(Index + 2).Name = "SupportThickness"
        End If
     End If
     Inputs(Index + 2).DefaultParameterValue = PC
  Next

  For Index = 1 To 5
    InputsIf.SetInput Inputs(Index), Index
    Set Inputs(Index) = Nothing
  Next

  Dim oSymbolCache As New CustomCache
  oSymbolCache.SetupCustomCache pSymbolDefinition, 6
  
  ' Create the output
  Dim O1 As IMSSymbolEntities.IJDOutput
  Set O1 = New IMSSymbolEntities.DOutput
  
  O1.Name = "HEBodyCylinder"
  O1.Description = "HeatX body"
  O1.Properties = 0
  
  Dim O2 As IMSSymbolEntities.IJDOutput
  Set O2 = New IMSSymbolEntities.DOutput
    
  O2.Name = "HeatExchangerControlPoint"
  O2.Description = "Control Point of Heat Exchanger Skid"
  O2.Properties = 0
  
  Dim Support1 As IMSSymbolEntities.IJDOutput
  Set Support1 = New IMSSymbolEntities.DOutput
  Support1.Name = "Support1"
  Support1.Description = "Support 1 of the Heat Exchanger"
  Support1.Properties = 0

  Dim Support2 As IMSSymbolEntities.IJDOutput
  Set Support2 = New IMSSymbolEntities.DOutput
  Support2.Name = "Support2"
  Support2.Description = "Support 2 of the Heat Exchanger"
  Support2.Properties = 0

  Dim DefaultSurface1 As IMSSymbolEntities.IJDOutput
  Set DefaultSurface1 = New IMSSymbolEntities.DOutput
  DefaultSurface1.Name = "DefaultSurface1"
  DefaultSurface1.Description = "Default plane one for Relations"
  DefaultSurface1.Properties = 0
  
  Dim DefaultSurface2 As IMSSymbolEntities.IJDOutput
  Set DefaultSurface2 = New IMSSymbolEntities.DOutput
  DefaultSurface2.Name = "DefaultSurface2"
  DefaultSurface2.Description = "Default Plane two for Relations"
  DefaultSurface2.Properties = 0
  
  Dim Head1 As IMSSymbolEntities.IJDOutput
  Set Head1 = New IMSSymbolEntities.DOutput
  Head1.Name = "Head1"
  Head1.Description = "HE Head"
  Head1.Properties = 0
  
  Dim Head2 As IMSSymbolEntities.IJDOutput
  Set Head2 = New IMSSymbolEntities.DOutput
  Head2.Name = "Head2"
  Head2.Description = "HE Head"
  Head2.Properties = 0


' Define the representation "Physical"
  Dim repPhysical As IMSSymbolEntities.IJDRepresentation
  Set repPhysical = New IMSSymbolEntities.DRepresentation
  
  repPhysical.Name = "Physical"
  repPhysical.Description = "Physical Representation of HeatExchanger"
  repPhysical.Properties = igREPRESENTATION_ISVBFUNCTION
  'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
  'different repIDs available.
  repPhysical.RepresentationId = SimplePhysical

  Dim oRepPhysicalOutputs As IMSSymbolEntities.IJDOutputs
  Set oRepPhysicalOutputs = repPhysical

  ' Set the output
  oRepPhysicalOutputs.SetOutput O1
  oRepPhysicalOutputs.SetOutput O2

  oRepPhysicalOutputs.SetOutput Support1
  Set Support1 = Nothing
  
  oRepPhysicalOutputs.SetOutput Support2
  Set Support2 = Nothing

  oRepPhysicalOutputs.SetOutput DefaultSurface1
  Set DefaultSurface1 = Nothing

  oRepPhysicalOutputs.SetOutput DefaultSurface2
  Set DefaultSurface2 = Nothing
  
  oRepPhysicalOutputs.SetOutput Head1
  Set Head1 = Nothing

  oRepPhysicalOutputs.SetOutput Head2
  Set Head2 = Nothing

  ' ---------Set the representation to definition
  Dim RepsIf As IMSSymbolEntities.IJDRepresentations
  Set RepsIf = pSymbolDefinition
        
  RepsIf.SetRepresentation repPhysical
 
  Set RepsIf = Nothing
  Set repPhysical = Nothing
  Set oRepPhysicalOutputs = Nothing
  
  ' ----------Definition of the function identification corresponding
  ' to the representation named "Physical" --
  Dim PhysicalRepEval As IJDRepresentationEvaluation
  Set PhysicalRepEval = New DRepresentationEvaluation

  PhysicalRepEval.Name = "Physical"   ' Note : The identification name of the evaluation function
                                      ' should be the same as the representation name.
  PhysicalRepEval.Description = "Physical representation"
  PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
  PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
  PhysicalRepEval.ProgId = "SP3DHeatXAsm.CHeatXSym" '

  ' ----------Set function identification into the symbol definition
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

'=========================================================================
'CREATION OF Physical REPRESENTATION OF HEAT EXCHANGER
'=========================================================================
Private Sub Physical(ByRef arrayOfInputs())
On Error GoTo ErrorLabel
    
    Dim X1 As Double, Y1 As Double, Z1 As Double
    Dim X2 As Double, Y2 As Double, Z2 As Double
    Dim length As Double, diameter As Double, supportLength As Double
    Dim supportHeight As Double, supportThk As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    
    length = arrayOfInputs(1)
    diameter = arrayOfInputs(2)
    supportLength = arrayOfInputs(3)
    supportHeight = arrayOfInputs(4)
    supportThk = arrayOfInputs(5)

    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(6)
    
    ' Origin is at the center of the cylinder(x,y) and at elevation of support(z)
    ' Creation of  HeatX body cylinder
    Dim dDelta As Double
    dDelta = Sqr((0.25 * diameter * diameter) - (supportLength * supportLength * 0.25))
    Dim circleCenterX As Double, circleCenterY As Double, circleCenterZ As Double
    circleCenterX = 0#
    circleCenterY = 0#
    circleCenterZ = supportHeight + dDelta

    Dim circleNormalX As Double, circleNormalY As Double, circleNormalZ As Double
    circleNormalX = 1#
    circleNormalY = 0#
    circleNormalZ = 0#

    Dim circlePointVecX As Double, circlePointVecY As Double, circlePointVecZ As Double
    circlePointVecX = 0#
    circlePointVecY = diameter * 0.5
    circlePointVecZ = 0#

    Dim axesRatio As Double
    axesRatio = 1#

    Dim ellipse As IngrGeom3D.Ellipse3d
    Set ellipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(m_outputColl.ResourceManager, _
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
                                    projVecX, projVecY, projVecZ, length, True)
    m_outputColl.AddOutput "HEBodyCylinder", projection
    Set projection = Nothing

    'Delete ellipse which was created for creating projection.
    Dim gelem As IJDObject
    Set gelem = ellipse
    Debug.Assert Not gelem Is Nothing
    Set ellipse = Nothing
    gelem.Remove
    
    
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector

    
    
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_outputColl.ResourceManager, 0, 0, 0, 0, , , True)
    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_outputColl.AddOutput "HeatExchangerControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    'BUILD SUPPORT PLATE1
    Dim ii As Integer
    Dim coords(1 To 6) As Double
    coords(1) = 0.7 - supportThk * 0.5
    coords(2) = -0.5 * supportLength
    coords(3) = 0#
    coords(4) = coords(1) + supportThk
    coords(5) = 0.5 * supportLength
    coords(6) = supportHeight

    ' Get or create the definition
    Dim defColl As IJDDefinitionCollection
    Set defColl = oSymbolFactory.DefinitionCollection(m_outputColl.ResourceManager)
    Dim boxDef As IJDSymbolDefinition
    Dim definitionParams As Variant
    definitionParams = ""
    Set boxDef = defColl.GetDefinitionByProgId(True, "Box.BoxServices", vbNullString, definitionParams)

    Dim Support1 As IMSSymbolEntities.DSymbol
    Set Support1 = oSymbolFactory.PlaceSymbol(boxDef, m_outputColl.ResourceManager)

    Dim newEnumArg1 As IJDEnumArgument
    Dim IJEditJDArg1 As IJDEditJDArgument
    Set newEnumArg1 = New DEnumArgument
    Set IJEditJDArg1 = newEnumArg1.IJDEditJDArgument

    Dim PC1 As IJDParameterContent
    Dim argument1 As IJDArgument
    For ii = 1 To 6
        Set PC1 = New DParameterContent
        Set argument1 = New DArgument

        PC1.UomValue = coords(ii)
        PC1.Type = igValue
        PC1.UomType = 1
        ' Feed the Argument
        argument1.Index = ii
        argument1.Entity = PC1
        ' Add the argument to the arg collection
        IJEditJDArg1.SetArg argument1
        Set PC1 = Nothing
        Set argument1 = Nothing
    Next

    Support1.IJDValuesArg.SetValues newEnumArg1
    Dim IJDInputsArg1 As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg1 = Support1
    IJDInputsArg1.Update
    Set IJDInputsArg1 = Nothing
    Set IJEditJDArg1 = Nothing
    Set newEnumArg1 = Nothing

    m_outputColl.AddOutput "Support1", Support1

    'BUILD SUPPORT PLATE2
    coords(1) = 2.8 - supportThk * 0.5
    coords(2) = -0.5 * supportLength
    coords(3) = 0#
    coords(4) = coords(1) + supportThk
    coords(5) = 0.5 * supportLength
    coords(6) = supportHeight

    Dim Support2 As IMSSymbolEntities.DSymbol
    Set Support2 = oSymbolFactory.PlaceSymbol(boxDef, m_outputColl.ResourceManager)

    Dim newEnumArg2 As IJDEnumArgument
    Dim IJEditJDArg2 As IJDEditJDArgument
    Set newEnumArg2 = New DEnumArgument
    Set IJEditJDArg2 = newEnumArg2.IJDEditJDArgument

    Dim PC2 As IJDParameterContent
    Dim argument2 As IJDArgument
    For ii = 1 To 6
        Set PC2 = New DParameterContent
        Set argument2 = New DArgument

        PC2.UomValue = coords(ii)
        PC2.Type = igValue
        PC2.UomType = 1
'         Feed the Argument
        argument2.Index = ii
        argument2.Entity = PC2
'         Add the argument to the arg collection
        IJEditJDArg2.SetArg argument2
        Set PC2 = Nothing
        Set argument2 = Nothing
    Next

    Support2.IJDValuesArg.SetValues newEnumArg2
    Dim IJDInputsArg2 As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg2 = Support2
    IJDInputsArg2.Update
    Set IJDInputsArg2 = Nothing
    Set IJEditJDArg2 = Nothing
    Set newEnumArg2 = Nothing

    m_outputColl.AddOutput "Support2", Support2


    Dim plane As IngrGeom3D.Plane3d
    Dim Points(0 To 11) As Double

    Points(0) = 0.7 - supportThk * 0.5
    Points(1) = -0.5 * supportLength
    Points(2) = 0#
    Points(3) = Points(0)
    Points(4) = Points(1) + supportLength
    Points(5) = 0#
    Points(6) = Points(0) + supportThk
    Points(7) = Points(1) + supportLength
    Points(8) = 0#
    Points(9) = Points(6)
    Points(10) = Points(1)
    Points(11) = 0#
    Set plane = geomFactory.Planes3d.CreateByPoints(m_outputColl.ResourceManager, 4, Points)
    m_outputColl.AddOutput "DefaultSurface1", plane
    Set plane = Nothing

    Points(0) = Points(0) + (2.8 - 0.7)
    Points(3) = Points(3) + 2.1
    Points(6) = Points(6) + 2.1
    Points(9) = Points(9) + 2.1
    
    Set plane = geomFactory.Planes3d.CreateByPoints(m_outputColl.ResourceManager, 4, Points)
    m_outputColl.AddOutput "DefaultSurface2", plane
    Set plane = Nothing

    ' Insert your code for output Head1(Elliptical Head 1)
    Dim ObjEndCap As Object
    Dim ellipseCenter As AutoMath.DPosition
    Dim ellipseNormal As AutoMath.DPosition
    Dim majorAxis As AutoMath.DPosition
    Dim PI As Double
          PI = 4 * Atn(1)

    Set majorAxis = New AutoMath.DPosition
    Set ellipseNormal = New AutoMath.DPosition
    Set ellipseCenter = New AutoMath.DPosition
    
    ellipseCenter.Set length, 0, supportHeight + dDelta
    majorAxis.Set 0, 0, diameter / 2
    ellipseNormal.Set 0, -1, 0
    axesRatio = 0.5
    Dim H1ellipse As Object
  
    Set H1ellipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                ellipseCenter.x, ellipseCenter.y, ellipseCenter.z, _
                                ellipseNormal.x, ellipseNormal.y, ellipseNormal.z, _
                                majorAxis.x, majorAxis.y, majorAxis.z, _
                                axesRatio, 1.5 * PI, PI / 2)
    Dim axisCenterX As Double, axisCenterY As Double, axisCenterZ As Double
    axisCenterX = 0#
    axisCenterY = 0#
    axisCenterZ = supportHeight + dDelta

    Dim axisVecX As Double, axisVecY As Double, axisVecZ As Double
    axisVecX = 1#
    axisVecY = 0#
    axisVecZ = 0#
    Set ObjEndCap = geomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, H1ellipse, _
                                    axisVecX, axisVecY, axisVecZ, axisCenterX, axisCenterY, axisCenterZ, _
                                    2 * PI, False)
   
' Set the output
    m_outputColl.AddOutput "Head1", ObjEndCap
    Set ObjEndCap = Nothing
    
    
    ' Insert your code for output Head2(Elliptical Head2)
    'Elliptical head minor radius is 0.5 times the Vessel Radius
    
    ellipseCenter.Set 0, 0, supportHeight + dDelta
    majorAxis.Set 0, 0, diameter / 2
    ellipseNormal.Set 0, 1, 0
    axesRatio = 0.5
    Set H1ellipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                ellipseCenter.x, ellipseCenter.y, ellipseCenter.z, _
                                ellipseNormal.x, ellipseNormal.y, ellipseNormal.z, _
                                majorAxis.x, majorAxis.y, majorAxis.z, _
                                axesRatio, 1.5 * PI, PI / 2)
    Set ObjEndCap = geomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, H1ellipse, _
                                    axisVecX, axisVecY, axisVecZ, axisCenterX, axisCenterY, axisCenterZ, _
                                    2 * PI, False)
' Set the output
    m_outputColl.AddOutput "Head2", ObjEndCap
    Set ObjEndCap = Nothing
    
    
    Set majorAxis = Nothing
    Set ellipseNormal = Nothing
    Set ellipseCenter = Nothing
    

    Set boxDef = Nothing
    Set defColl = Nothing
    
    Set pos = Nothing
    Set dir = Nothing
    
    Set oPart = Nothing
    Set oSymbolFactory = Nothing
    Set geomFactory = Nothing

    
Exit Sub
ErrorLabel:
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "HeatExhanger definition"
    Debug.Assert False
    Err.Raise Err.Number, "Physical", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

