VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGateValve300Asym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   GateValve300 - Solid wedge gate valve for 300 lbs pressure rate
'   06/29/99
'   LK
'   revision 07/21/99
'   7/15/99 - multiple representations
'   7/21/99 - Put user's code in class initialization
'   revision 20010201 NN constructed an asymmetric Drain Vent / Instrument Root Gate Valve SW X FT 800 # with graphics similar to the Class 300.
'   ProgId: SP3DGateValveAsym.CGateValve300Asym
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox
'  1.NOV.2007       RRK                    CR-123952 Updated the symbol to use retrive parameters function in basgeom3d.bas

Option Explicit

Private Const MODULE = "GateValve300Asym" 'Used for error messages

Private m_SymbolFactory     As IMSSymbolEntities.DSymbolEntitiesFactory
Private m_OutputColl        As Object
Private m_ProgId            As String
Private gscadElem           As IJDObject

Private geomFactory         As IngrGeom3D.GeometryFactory
Private objBspline          As IngrGeom3D.BSplineCurve3d
Private objComplexString    As IngrGeom3D.ComplexString3d

Private discDiam        As Double
Private bodyVDiam1      As Double
Private bodyVDiam2      As Double
Private bonnetWidth     As Double
Private actWidth        As Double
Private actDiam         As Double
Private stemDiam        As Double
Private currentWidth    As Double

Private InputTypes()            As InputType
Private OutputTypes()           As OutputType
Private RepresentationTypes()   As RepresentationType
Private nRepresentations        As Integer
Private nInputs                 As Integer
Private nOutputs                As Integer
    
Private PI As Double

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Dim iCount As Integer
    
    m_ProgId = "SP3DGateValveAsym.CGateValve300Asym"

    PI = 4 * Atn(1)
    
' Inputs
    nInputs = 4
    ReDim InputTypes(1 To nInputs) As InputType
    
    iCount = 1
    InputTypes(iCount).name = "Face1toCenter"
    InputTypes(iCount).description = "Face1 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.06
    
    iCount = 1 + iCount
    InputTypes(iCount).name = "Face2toCenter"
    InputTypes(iCount).description = "Face2 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.08

    iCount = 1 + iCount
    InputTypes(iCount).name = "OperatorHeight"
    InputTypes(iCount).description = "Height of Operator"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.614
        
    iCount = 1 + iCount
    InputTypes(iCount).name = "OperatorDiameter"
    InputTypes(iCount).description = "Diameter of Operator Wheel"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.25
        
' Outputs
    'nOutputs = 12
    nOutputs = 11
    
    ReDim OutputTypes(1 To nOutputs) As OutputType
    iCount = 1
    OutputTypes(iCount).name = "BodyH"
    OutputTypes(iCount).description = "Horisontal Body of GateValve"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "BodyV"
    OutputTypes(iCount).description = "Vertical Body of GateValve; Bottom Part"
    OutputTypes(iCount).representation = SimplePhysical

    iCount = 1 + iCount
    OutputTypes(iCount).name = "BodyVFlange"
    OutputTypes(iCount).description = "Vertical Body Flange of GateValve"
    OutputTypes(iCount).representation = SimplePhysical

    iCount = 1 + iCount
    OutputTypes(iCount).name = "BonnetFlange"
    OutputTypes(iCount).description = "Bonnet Flange of GateValve"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "Bonnet"
    OutputTypes(iCount).description = "Bonnet of GateValve"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "Stem"
    OutputTypes(iCount).description = "Stem of GateValve"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "Actuator"
    OutputTypes(iCount).description = "Actuator of GateValve"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "ActuatorSpike1"
    OutputTypes(iCount).description = "Spike 1 of Actuator wheel"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "ActuatorSpike2"
    OutputTypes(iCount).description = "Spike 2 of Actuator wheel"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    'OutputTypes(iCount).name = "PipePort1"
    OutputTypes(iCount).name = "VNoz1"
    OutputTypes(iCount).description = "PipingPort1 of Gate valve"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    'OutputTypes(iCount).name = "PipePort2"
    OutputTypes(iCount).name = "VNoz2"
    OutputTypes(iCount).description = "PipingPort2 of Gate valve"
    OutputTypes(iCount).representation = SimplePhysical
    
'''    iCount = 1 + iCount
'''    OutputTypes(iCount).name = "GateValveMaintenanceBox"
'''    OutputTypes(iCount).description = "Gate Valve Maintenance space"
'''    OutputTypes(iCount).representation = Maintenance
    
' Representations
    'nRepresentations = 2
    nRepresentations = 1
    ReDim RepresentationTypes(1 To nRepresentations) As RepresentationType
    
    iCount = 1
    RepresentationTypes(iCount).name = "Physical"
    RepresentationTypes(iCount).description = "Physical representation"
    RepresentationTypes(iCount).properties = igREPRESENTATION_ISVBFUNCTION
    RepresentationTypes(iCount).representationId = SimplePhysical
    
'''    iCount = iCount + 1
'''    RepresentationTypes(iCount).name = "Maintenance"
'''    RepresentationTypes(iCount).description = "Maintenance Represntation of the GateValve"
'''    RepresentationTypes(iCount).properties = igREPRESENTATION_ISVBFUNCTION
'''    RepresentationTypes(iCount).representationId = Maintenance
    
    Exit Sub
    
ErrorHandler:
  ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub Class_Terminate()
    
    Set m_SymbolFactory = Nothing

End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition( _
                ByVal CodeBase As String, _
                ByVal defParameters As Variant, _
                ByVal ActiveConnection As Object) As Object
    
    ' This method is in charge of the creation of the symbol definition object
    ' You can keep the current design unchanged
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition:"
    On Error GoTo ErrorHandler
    
    Dim oSymbolDefinition  As IMSSymbolEntities.IJDSymbolDefinition
    
    ' Create the definition from the symbol factory
    Set m_SymbolFactory = New IMSSymbolEntities.DSymbolEntitiesFactory
    Set oSymbolDefinition = m_SymbolFactory.CreateEntity(Definition, ActiveConnection)
    Set m_SymbolFactory = Nothing
    ' Initialize it
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = m_ProgId
    oSymbolDefinition.CodeBase = CodeBase

    ' Give a unique name to the symbol definition
    oSymbolDefinition.name = oSymbolDefinition.ProgId
   
    'returned symbol definition
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    Debug.Assert False

End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  
  IJDUserSymbolServices_GetDefinitionName = m_ProgId

End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
    
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"
    On Error GoTo ErrorHandler
   
    Dim oPart   As IMSSymbolEntities.DInput
    
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
    
   Dim oSymbolCache As New CustomCache
   oSymbolCache.SetupCustomCache pSymbolDefinition

    
' ----------Parameter Input --------------
    Dim Inputs()    As IMSSymbolEntities.IJDInput
    ReDim Inputs(1 To nInputs) As IMSSymbolEntities.IJDInput
    Dim iCount      As Integer
' Create a default parameter
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent
    PC.Type = igValue
    For iCount = 1 To nInputs
        Set Inputs(iCount) = New IMSSymbolEntities.DInput
        Inputs(iCount).name = InputTypes(iCount).name
        Inputs(iCount).description = InputTypes(iCount).description
        Inputs(iCount).properties = InputTypes(iCount).properties
        PC.uomValue = InputTypes(iCount).uomValue
' Associate a default parameter to the input
        Inputs(iCount).DefaultParameterValue = PC
' ---------Set Input into the definition-
        InputsIf.SetInput Inputs(iCount), iCount + 1
        Set Inputs(iCount) = Nothing
    Next iCount
        
' Define output
    Dim symbRepresentation As IMSSymbolEntities.IJDRepresentation
    Set symbRepresentation = New IMSSymbolEntities.DRepresentation
' ---------Set the representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    Dim Output As IMSSymbolEntities.IJDOutput
    Dim iOutput      As Integer
    Dim oSymbAspect As IMSSymbolEntities.IJDOutputs
    
    For iCount = 1 To nRepresentations
        symbRepresentation.name = RepresentationTypes(iCount).name
        symbRepresentation.description = RepresentationTypes(iCount).description
        symbRepresentation.properties = RepresentationTypes(iCount).properties
        symbRepresentation.representationId = RepresentationTypes(iCount).representationId
        'symbRepresentation.RemoveAllOutput
    
    Set oSymbAspect = symbRepresentation
        
        Set Output = New IMSSymbolEntities.DOutput
        For iOutput = 1 To nOutputs
            If OutputTypes(iOutput).representation And RepresentationTypes(iCount).representationId Then
                Output.name = OutputTypes(iOutput).name
                Output.description = OutputTypes(iOutput).description
                Output.properties = 0            ' OutputTypes(iOutput).Properties
                oSymbAspect.SetOutput Output
            End If
        Next iOutput
        RepsIf.SetRepresentation symbRepresentation
        Set Output = Nothing
    Next iCount
    

    ' ----------Definition of the function identification corresponding
    Dim oRepEval     As IJDRepresentationEvaluation
    Dim RepsEvalsIf     As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepsEvalsIf = pSymbolDefinition
    
    For iCount = 1 To nRepresentations
        Set oRepEval = New DRepresentationEvaluation
        oRepEval.name = RepresentationTypes(iCount).name
        oRepEval.description = RepresentationTypes(iCount).description
        oRepEval.properties = igREPRESENTATION_HIDDEN
        oRepEval.Type = igREPRESENTATION_VBFUNCTION
        oRepEval.ProgId = m_ProgId
        RepsEvalsIf.AddRepresentationEvaluation oRepEval
        Set oRepEval = Nothing
    Next iCount
    Set RepsEvalsIf = Nothing
    
'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
        
  Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    Debug.Assert False

End Sub

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                    ByVal repName As String, _
                    ByVal outputcoll As Object, _
                    ByRef arrayOfInputs())
    
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler
    
    Dim arrayOfOutputs()    As String
    Dim iOutputs            As Integer
    Dim iCount              As Integer
    
    Set m_OutputColl = outputcoll
    
    Select Case repName
    
    Case "Physical"
        'Msgbox "in case physical nOutputs=" & nOutputs
        
        iOutputs = 1
        For iCount = 1 To nOutputs
            'Msgbox "rep type=" & OutputTypes(iCount).representation & ", simple=" & SimplePhysical
            If OutputTypes(iCount).representation = SimplePhysical Then
                ReDim Preserve arrayOfOutputs(1 To iOutputs) As String
                arrayOfOutputs(iOutputs) = OutputTypes(iCount).name
                iOutputs = iOutputs + 1
            End If
        Next iCount
        'Msgbox "number of outputs counted = " & iOutputs
        getPhysical arrayOfInputs, arrayOfOutputs
    
'''    Case "Maintenance"
'''        iOutputs = 1
'''        For iCount = 1 To nOutputs
'''            If OutputTypes(iCount).representation = Maintenance Then
'''                ReDim Preserve arrayOfOutputs(1 To iOutputs) As String
'''                arrayOfOutputs(iOutputs) = OutputTypes(iCount).name
'''                iOutputs = iOutputs + 1
'''            End If
'''        Next iCount
'''        getMaintenance arrayOfInputs, arrayOfOutputs
    
    Case Else
'        MsgBox "Representation not found"
    
    End Select
    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    
    ' The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False

End Function

Private Sub getMaintenance(ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "getMaintenance:"
    On Error GoTo ErrorHandler
    
    Dim lValveWidth     As Double
    Dim lvalveHeight    As Double
    Dim lactDiam        As Double
    Dim mntnBoxWidth    As Double
    Dim mntnBoxHeight   As Double
    Dim mntnBoxDepth    As Double
    Dim Coords(1 To 6)  As Double
    
    lValveWidth = arrayOfInputs(2)
    lvalveHeight = arrayOfInputs(4)
    lactDiam = arrayOfInputs(5)
    
    mntnBoxWidth = 1.6 * lValveWidth
    If mntnBoxWidth < 2 * lactDiam Then mntnBoxWidth = 2 * lactDiam
    mntnBoxHeight = 2 * lvalveHeight
    mntnBoxDepth = 1.2 * lValveWidth
        
    Coords(1) = -mntnBoxWidth / 2
    Coords(2) = -0.2 * mntnBoxHeight
    Coords(3) = -mntnBoxDepth / 2
    Coords(4) = mntnBoxWidth / 2
    Coords(5) = 0.8 * mntnBoxHeight
    Coords(6) = mntnBoxDepth / 2

    ' Get or create the definition
    Dim defColl             As IJDDefinitionCollection
    Dim boxDef              As IJDSymbolDefinition
    Dim definitionParams    As Variant
    Dim oEnv                As IMSSymbolEntities.DSymbol
    Dim newEnumArg          As IJDEnumArgument
    Dim IJEditJDArg         As IJDEditJDArgument
    Dim PC                  As IJDParameterContent
    Dim argument            As IJDArgument
    Dim iCount              As Integer
    
    Set m_SymbolFactory = New IMSSymbolEntities.DSymbolEntitiesFactory
    Set defColl = m_SymbolFactory.DefinitionCollection(m_OutputColl.ResourceManager)
    definitionParams = ""
    Set boxDef = defColl.GetDefinitionByProgId(True, "Box.BoxServices", vbNullString, definitionParams)

    Set oEnv = m_SymbolFactory.PlaceSymbol(boxDef, m_OutputColl.ResourceManager)
    Set m_SymbolFactory = Nothing
    Set boxDef = Nothing
    Set defColl = Nothing

    Set newEnumArg = New DEnumArgument
    Set IJEditJDArg = newEnumArg.IJDEditJDArgument

    For iCount = 1 To 6
        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = Coords(iCount)
        PC.Type = igValue
        PC.UomType = 1
        ' Feed the Argument
        argument.index = iCount
        argument.Entity = PC
        ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
    Next

    oEnv.IJDValuesArg.SetValues newEnumArg
    Dim IJDInputsArg As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg = oEnv
    IJDInputsArg.Update
    Set IJDInputsArg = Nothing
    Set IJEditJDArg = Nothing
    Set newEnumArg = Nothing

    m_OutputColl.AddOutput arrayOfOutputs, oEnv
    Set oEnv = Nothing

    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    Debug.Assert False

End Sub

Private Sub getPhysical(ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "getPhysical:"
    On Error GoTo ErrorHandler
        
    Dim oPipePort       As GSCADNozzleEntities.IJDPipePort
    Dim nozzleIndex     As Integer
    Dim circleCenter    As AutoMath.DPosition
    Dim circleNormal    As AutoMath.DVector
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim projVec         As AutoMath.DVector
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim ValveL          As Double
    Dim NozzleFactory   As GSCADNozzleEntities.NozzleFactory
    Dim oNozzle         As GSCADNozzleEntities.IJDNozzle
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim cptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim bodyDiam        As Double
    Dim valveWidth      As Double
    Dim flangeThick     As Double
    Dim cptOffset            As Double
    Dim flangeDiam      As Double
    Dim depth         As Double
'    Dim discWidth       As Double
    Dim stemFlWidth     As Double
    Dim stemFlDiam      As Double
    
    Dim face1c          As Double
    Dim face2c          As Double
    Dim faceToFace      As Double
    
    faceToFace = face1c + face2c
    
    Dim ValveL1c        As Double
    Dim ValveL2c        As Double
    
    Set oPartFclt = arrayOfInputs(1)
    Set geomFactory = New IngrGeom3D.GeometryFactory

' Use nozzle 1 as pipe port for flange and pipe parameters
    RetrieveParameters 1, oPartFclt, m_OutputColl, bodyDiam, flangeThick, flangeDiam, cptOffset, depth
    
    face1c = arrayOfInputs(2)
    face2c = arrayOfInputs(3)
    valveWidth = arrayOfInputs(4)
    actDiam = arrayOfInputs(5)
    
'    ValveL = discWidth - (2 * flangeThick)
    ValveL1c = face1c - flangeThick
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, cptOffset2, depth2
    ValveL2c = face2c - flangeThick2
    ValveL = ValveL1c + ValveL2c
'   keep oNozzle for nozzle1 cretation later

    discDiam = (bodyDiam + flangeDiam) / 2.2
    bodyVDiam1 = bodyDiam
    bodyVDiam2 = discDiam
    stemFlWidth = flangeThick
    
    stemFlDiam = flangeDiam
    bonnetWidth = valveWidth * 0.15
    stemDiam = bodyDiam * 0.4
    actWidth = bodyDiam * 0.25
    
'----------------------- Horisontal Body
' Construction of Horisontal Body Cylinder
    Set circleCenter = New AutoMath.DPosition
    circleCenter.Set -ValveL1c, 0, 0
    Set circleNormal = New AutoMath.DVector
    circleNormal.Set 1, 0, 0
    
' Construct a circle that will be used to project the disc
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        discDiam * 0.5)
    
    Set projVec = New AutoMath.DVector
    projVec.Set 1, 0, 0

' Project the disc of body
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                        objCircle, _
                                                        projVec.x, projVec.y, projVec.z, _
                                                        ValveL, False)
    
' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(1), objProjection
    Set objProjection = Nothing
    
'Delete the circle that was created for the projection
    Set gscadElem = objCircle
    Set objCircle = Nothing
    gscadElem.Remove
    
'----------------------- Body Vertical
 'Msgbox "about to create the bodyV"
   
    CreateBodyV m_OutputColl, arrayOfOutputs(2)
    currentWidth = discDiam / 2 + Sqr(bodyVDiam2 * _
                bodyVDiam2 - bodyVDiam1 * bodyVDiam1) / 2

'----------------------- Body Vertical Flange
' Construction of Body Vertical Flange Cylinder, starting where the Body Vertical left off
    circleCenter.Set 0, currentWidth, 0
    
' Construct a circle that will be used to project the Flange
    circleNormal.Set 0, 1, 0
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                        circleCenter.x, circleCenter.y, circleCenter.z, _
                                                        circleNormal.x, circleNormal.y, circleNormal.z, _
                                                        stemFlDiam * 0.5)
' Project the Flange
    projVec.Set 0, 1, 0
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                            objCircle, _
                                                            projVec.x, projVec.y, projVec.z, _
                                                            stemFlWidth, True)
    
' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(3), objProjection
    Set objProjection = Nothing
    currentWidth = currentWidth + stemFlWidth
'Delete the circle that was created for the projection

    Set gscadElem = objCircle
    Set objCircle = Nothing
    gscadElem.Remove
    
     
'----------------------- bonnet flange
' Construction of bonnet flange Cylinder, starting where the vertical body left off
    circleCenter.Set 0, currentWidth + 0.002, 0
    
    ' Construct a circle that will be used to project the actuator
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                        circleCenter.x, circleCenter.y, circleCenter.z, _
                                                        circleNormal.x, circleNormal.y, circleNormal.z, _
                                                        stemFlDiam * 0.5)
' Project the actuator body
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                    objCircle, _
                                                    projVec.x, projVec.y, projVec.z, _
                                                    stemFlWidth, True)
    
' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(4), objProjection
    Set objProjection = Nothing
    
    currentWidth = currentWidth + stemFlWidth + 0.002

'Delete the circle that was created for the projection
    Set gscadElem = objCircle
    Set objCircle = Nothing
    gscadElem.Remove
    
    '----------------------- bonnet

    CreateBonnet m_OutputColl, arrayOfOutputs(5), currentWidth
    currentWidth = currentWidth + bonnetWidth
    
    '----------------------- stem body
    ' Construction of stem Cylinder, starting where the top of the disc
    circleCenter.Set 0, currentWidth, 0
    circleNormal.Set 0, 1, 0
    
    ' Construct a circle that will be used to project the stem
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        stemDiam * 0.5)

    ' Project the Stem body
    projVec.Set 0, 1, 0
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                        objCircle, _
                        projVec.x, projVec.y, projVec.z, _
                        valveWidth - currentWidth, True)
    
    ' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(6), objProjection
    Set objProjection = Nothing
    currentWidth = valveWidth - 3 * actWidth
    
'Delete the circle that was created for the projection
    Set gscadElem = objCircle
    Set objCircle = Nothing
    gscadElem.Remove
    
'----------------------- actuator body
' Creation of actuator torus

    CreateActuator m_OutputColl, arrayOfOutputs(7), currentWidth

' Construction of actuator spike1 (cylinder)
    circleCenter.Set -actDiam / 2, currentWidth, 0
    circleNormal.Set 1, 0, 0
' Construct a circle that will be used to project the spike
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        actWidth * 0.25)
' Project the Spike body
    projVec.Set 1, 0, 0
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                        objCircle, _
                        projVec.x, projVec.y, projVec.z, _
                        actDiam, True)
' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(8), objProjection

' Construction of actuator spike2
    circleCenter.Set 0, currentWidth, -actDiam / 2
    circleNormal.Set 0, 0, 1
' Construct a circle that will be used to project the spike
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        actWidth * 0.25)
                        
' Project the Spike body
    projVec.Set 0, 0, 1
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                        objCircle, _
                        projVec.x, projVec.y, projVec.z, _
                        actDiam, True)
    
' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(9), objProjection
    Set objProjection = Nothing
    
'Delete the circle that was created for the projection
    Set gscadElem = objCircle
    Set objCircle = Nothing
    gscadElem.Remove
    
''===========================
''Construction of nozzle1
''===========================

    Dim pos                 As New AutoMath.DPosition
    Dim dir                 As New AutoMath.DVector
    Dim xctrpt              As Double

    dir.Set -1, 0, 0
    'Position of the nozzle should be the connect point of the nozzle
    xctrpt = face1c - (depth - cptOffset)
    pos.Set -xctrpt, 0, 0

    Set oNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, dir, pos)
    
'Set the Output
    m_OutputColl.AddOutput arrayOfOutputs(10), oNozzle
    Set oNozzle = Nothing

'======================
' Creation of nozzle2
'======================
    
    dir.Set 1, 0, 0
    
    'Position of the nozzle should be the connect point of the nozzle
    xctrpt = face2c - (depth2 - cptOffset2)
    pos.Set xctrpt, 0, 0
    
    Set oNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, dir, pos)
'Set the output
    m_OutputColl.AddOutput arrayOfOutputs(11), oNozzle
    Set oNozzle = Nothing
    Set geomFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing
    
    Set oPartFclt = Nothing
    
    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    Debug.Assert False

End Sub

Private Sub CreateBodyV(m_OutputColl As Object, outname As String)

' Construction of revolution object, starting where the top of the disc
    Const METHOD = "CreateBodyV:"
    On Error GoTo ErrorHandler
        
    Dim arcX1   As Double
    Dim arcY1   As Double
    Dim arcZ1   As Double
    Dim arcX2   As Double
    Dim arcY2   As Double
    Dim arcZ2   As Double
    Dim arcX3   As Double
    Dim arcY3   As Double
    Dim arcZ3   As Double
    Dim lineX1  As Double
    Dim lineY1  As Double
    Dim lineZ1  As Double
    Dim lineX2  As Double
    Dim lineY2  As Double
    Dim lineZ2  As Double
    Dim bodyVH  As Double
    
    Dim objArc              As IngrGeom3D.Arc3d
    Dim objRevolution       As IngrGeom3D.Revolution3d
    
    bodyVH = Sqr(bodyVDiam2 * bodyVDiam2 - bodyVDiam1 * bodyVDiam1) / 2

    arcY1 = ((discDiam / 2) + bodyVH) / 2
    arcX1 = -(0.5 * Sqr((discDiam * discDiam) - (4 * arcY1 * arcY1)))
    arcZ1 = 0
    arcX2 = -bodyVDiam1 / 2
    arcY2 = discDiam / 2
    arcZ2 = 0
    
    arcX3 = -bodyVDiam2 / 2
    arcY3 = (discDiam / 2) + bodyVH
    arcZ3 = 0

  

    Set objArc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                                    arcX1, arcY1, arcZ1, _
                                                    arcX2, arcY2, arcZ2, _
                                                    arcX3, arcY3, arcZ3)

'' Construct complex string from line and arc that will be used to construct the surface
'    Set curvesColl = New UntransactedMiddleElems
'    curvesColl.Add objArc
'    curvesColl.Add objLine
'    Set objComplexString = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, _
                                            curvesColl)
' Construct revolution from arc
    Dim axis    As New AutoMath.DVector
    Dim center  As New AutoMath.DPosition
    
    axis.Set 0, 1, 0
    center.Set 0, arcY3, 0
    Set objRevolution = geomFactory.Revolutions3d.CreateByCurve( _
                                                    m_OutputColl.ResourceManager, _
                                                    objArc, _
                                                    axis.x, axis.y, axis.z, _
                                                    center.x, center.y, center.z, _
                                                    2 * PI, False)

' Set the output
    m_OutputColl.AddOutput outname, objRevolution
    Set objRevolution = Nothing

' Release objects
    Set gscadElem = objArc
    Set objArc = Nothing
    gscadElem.Remove
    
    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    Debug.Assert False

End Sub

Private Sub CreateBonnet(m_OutputColl As Object, outname As String, currentWidth As Double)
    
' Construction of bonnet cone, starting where the bonnet flange left off
    Const METHOD = "CreateBonnet:"
    On Error GoTo ErrorHandler
        
    Dim centerBase  As New AutoMath.DPosition
    Dim centerTop   As New AutoMath.DPosition
    Dim startBase   As New AutoMath.DPosition
    Dim startTop    As New AutoMath.DPosition
    
    Dim objCone     As IngrGeom3D.Cone3d
    
    centerBase.Set 0, currentWidth, 0
    centerTop.Set 0, currentWidth + bonnetWidth, 0
    startBase.Set bodyVDiam2 / 2, currentWidth, 0
    startTop.Set stemDiam / 2, currentWidth, 0
    Set objCone = geomFactory.Cones3d.CreateBy4Pts(m_OutputColl.ResourceManager, _
                                                    centerBase.x, centerBase.y, centerBase.z, _
                                                    centerTop.x, centerTop.y, centerTop.z, _
                                                    startBase.x, startBase.y, startBase.z, _
                                                    startTop.x, startTop.y, startTop.z, _
                                                    False)
' Set the output
    m_OutputColl.AddOutput outname, objCone

' Release objects
    Set gscadElem = objCone
    Set objCone = Nothing
    Set objCone = Nothing
    
    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    Debug.Assert False

End Sub

Private Sub CreateActuator(m_OutputColl As Object, outname As String, currentWidth As Double)

' Construction of actuator torus
    
    Const METHOD = "CreateActuator:"
    On Error GoTo ErrorHandler
        
    Dim centerPoint     As New AutoMath.DPosition
    Dim vNormal         As New AutoMath.DVector
    Dim originDir       As New AutoMath.DVector
    Dim minorRadius     As Double
    
    Dim objTorus     As IngrGeom3D.Torus3d
    
    centerPoint.Set 0, currentWidth, 0
    vNormal.Set 0, 1, 0
    originDir.Set 1, 0, 0
    minorRadius = actWidth / 2
    Set objTorus = geomFactory.Tori3d.CreateByAxisMajorMinorRadius(m_OutputColl.ResourceManager, _
                                        centerPoint.x, centerPoint.y, centerPoint.z, _
                                        vNormal.x, vNormal.y, vNormal.z, _
                                        originDir.x, originDir.y, originDir.z, _
                                        actDiam / 2, minorRadius, False)
    
' Set the output
    m_OutputColl.AddOutput outname, objTorus

' Release objects
    Set gscadElem = objTorus
    Set objTorus = Nothing
    Set objTorus = Nothing

    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    Debug.Assert False

End Sub
