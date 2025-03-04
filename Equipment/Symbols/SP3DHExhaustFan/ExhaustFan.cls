VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExhaustFan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:00:03 AM"
'******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    ExhaustFan.cls
'
'Author
'    GW, used SP3DAirDistribAssemblyAsm from RM
'
'Description
'    Definition of Exhaust Fans  ,  Ventilator.
'
'   dd.mmm.yyyy     who                 change description
'   -----------     ---                 ------------------
'   03.Nov.2005     kkk     CR-87366  Create TROX HVAC parts and symbols
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Const NEGLIGIBLE_THICKNESS = 0.0001

'GW.
Private m_GeomFactory As IngrGeom3D.GeometryFactory

Private m_lngParameterIndex As Long

Private Const mc_lngNParameters = 18

Implements IJDUserSymbolServices

Private Sub Class_Initialize()

Set m_GeomFactory = New IngrGeom3D.GeometryFactory

End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHExhaustFan.ExhaustFan"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  LogErrorGer "IJDUserSymbolServices_InstanciateDefinition"
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHExhaustFan.ExhaustFan"
  oSymbolDefinition.CodeBase = CodeBase
  
  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId
  
  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  Set oSymbolDefinition = Nothing
  
  Exit Function

ErrorHandler:
    LogErrorGer "IJDUserSymbolServices_InstanciateDefinition " & Err.Description
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Function


Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    
LogErrorGer "IJDUserSymbolServices_InitializeSymbolDefinition"
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
    Dim Inputs(1 To mc_lngNParameters) As IMSSymbolEntities.IJDInput
    Dim index As Integer
  
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue

    Set Inputs(1) = New IMSSymbolEntities.DInput
    Inputs(1).Name = "BasicAsmWidth"
    Inputs(1).Description = "Width of the Basic Assembly"
    Inputs(1).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.3048
    Inputs(1).DefaultParameterValue = PC
    
    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).Name = "BasicAsmHeight"
    Inputs(2).Description = "Height of the basic Assembly"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.2032
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).Name = "BasicAsmLength"
    Inputs(3).Description = "Length of the basic assembly"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.3239
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).Name = "HeatingCoilLength"
    Inputs(4).Description = "Length of the heating coil"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Inputs(4).DefaultParameterValue = PC
    
    Set Inputs(5) = New IMSSymbolEntities.DInput
    Inputs(5).Name = "MultiOutletLength"
    Inputs(5).Description = "Length of the multi outlet"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.9
    Inputs(5).DefaultParameterValue = PC
    
    Set Inputs(6) = New IMSSymbolEntities.DInput
    Inputs(6).Name = "InletDia"
    Inputs(6).Description = "Inlet Diameter"
    Inputs(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.0984
    Inputs(6).DefaultParameterValue = PC

    Set Inputs(7) = New IMSSymbolEntities.DInput
    Inputs(7).Name = "MultiOutletDia"
    Inputs(7).Description = "MultiOutlet Diameter"
    Inputs(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1492
    Inputs(7).DefaultParameterValue = PC
    
    m_lngParameterIndex = 8
    Call AddParameter(Inputs, "SieSide", 1)
    Call AddParameter(Inputs, "SieBEZ")
    Call AddParameter(Inputs, "SieRadius", 1)
    Call AddParameter(Inputs, "SieRadius2", 1)
    Call AddParameter(Inputs, "SieL", 1)
    Call AddParameter(Inputs, "SieFlanschB", 1)
    Call AddParameter(Inputs, "SieFlanschD", 1)
    Call AddParameter(Inputs, "SieL2", 1)
    Call AddParameter(Inputs, "SieL3", 1)
    Call AddParameter(Inputs, "SieWidth", 1)
    Call AddParameter(Inputs, "SieDepth", 1)
    
    If m_lngParameterIndex <> mc_lngNParameters + 1 Then
        Debug.Print "Invalid number of Parameters specified"
        Debug.Assert False
    End If
    

    Set Inputs(8) = New IMSSymbolEntities.DInput
    Inputs(8).Name = "SieSide"
    Inputs(8).Description = "SieSide"
    Inputs(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 1
    Inputs(8).DefaultParameterValue = PC
    
    Set Inputs(9) = New IMSSymbolEntities.DInput
    Inputs(9).Name = "SieBEZ"
    Inputs(9).Description = "SieBEZ"
    Inputs(9).Properties = igINPUT_IS_A_PARAMETER
    
    PC.Type = igString
    PC.String = "25"
    Inputs(9).DefaultParameterValue = PC
    
    For index = 1 To mc_lngNParameters
        InputsIf.SetInput Inputs(index), index + 1
        Set Inputs(index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 2) As IMSSymbolEntities.IJDOutput
    
    For index = 1 To 2
        Set O(index) = New IMSSymbolEntities.DOutput
        O(index).Properties = 0
    Next

    O(1).Name = "BasicAsmHVACNozzle"
    O(1).Description = "HvacPort on the basic assembly"


    O(2).Name = "Feature1"
    O(2).Description = "Feature n"
    O(2).Properties = igCOLLECTION_VARIABLE
 
    
    'Define the representation "Physical"
    Dim rep1 As IMSSymbolEntities.IJDRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation
  
    rep1.Name = "Physical"
    rep1.Description = "Physical Represntation of the Air Distrib Assembly"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim oRep1Outputs As IMSSymbolEntities.IJDOutputs
    Set oRep1Outputs = rep1

    oRep1Outputs.Property = igCOLLECTION_VARIABLE
    For index = 1 To 2
        oRep1Outputs.SetOutput O(index)
        Set O(index) = Nothing
    Next

    'Set the representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1
    
'===================================================================
' DEFINE MAINTENANCE ENVELOPE REPRESENTATION WITH DIFFERENT OUTPUTS
'===================================================================
    ' Redefine outputs
    oRep1Outputs.RemoveAllOutput
      
    ' define output
    Set O(1) = New IMSSymbolEntities.DOutput
    
    O(1).Name = "VVADAEnvelope"
    O(1).Description = "VVADA's Maintenance envelope"
    O(1).Properties = 0
    

    oRep1Outputs.SetOutput O(1)
    Set O(1) = Nothing

    ' Define representation
    rep1.Name = "MaintenanceEnvelope"
    rep1.Description = "Maintenance Envelope Representation of the VVADA"
    'Set the repID to Maintenance. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = Maintenance
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION

    ' Set representation
    RepsIf.SetRepresentation rep1
    
    Set rep1 = Nothing
    Set RepsIf = Nothing
    Set oRep1Outputs = Nothing

    'Define evaluation for Physical representation
    Dim PhysicalRepEval As IJDRepresentationEvaluation
    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "script for the Physical representation"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "SP3DHExhaustFan.ExhaustFan"

    'Define the evaluation associated to the Envelope representation
    Dim EnvelopeRepEval As DRepresentationEvaluation
    Set EnvelopeRepEval = New DRepresentationEvaluation
    EnvelopeRepEval.Name = "MaintenanceEnvelope"
    EnvelopeRepEval.Description = "Maintenance Envelope representation"
    EnvelopeRepEval.Properties = igREPRESENTATION_HIDDEN
    EnvelopeRepEval.Type = igREPRESENTATION_VBFUNCTION
    EnvelopeRepEval.ProgId = "SP3DHExhaustFan.ExhaustFan"

    'Set the evaluations for the Physical representation on the definition
    Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    RepEvalsIf.AddRepresentationEvaluation EnvelopeRepEval
    Set PhysicalRepEval = Nothing
    Set EnvelopeRepEval = Nothing
    Set RepEvalsIf = Nothing

'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================

    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

    LogErrorGer "IJDUserSymbolServices_InitializeSymbolDefinition finished."
  Exit Sub

ErrorHandler:
LogErrorGer "IJDUserSymbolServices_InitializeSymbolDefinition " & Err.Description
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False
   Exit Sub
   Resume
   
End Sub


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    
LogErrorGer "IJDUserSymbolServices_InvokeRepresentation" & repName
    Set m_outputColl = outputcoll
    If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
    ElseIf StrComp(repName, "MaintenanceEnvelope") = 0 Then
'        MaintenanceEnvelopeRep arrayOfInputs
    End If

End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF VVADA

'
'  Motor              Spiralgehäuse:
'
'
'
'                     ------------
'                    /            \
'       O            |             |
'                    \             /
'                      -----------
'
'
'
'                       ^
'                       |
'                       |
'                       |
'                       |
'                       |
'                       x
'    <--------------y

Private Sub Physical(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    
    Dim ii As Integer
    
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim BasicAsmWidth As Double, BasicAsmHeight As Double
    Dim BasicAsmLength As Double, HeatingCoilLength As Double
    Dim MultiOutletLength As Double
    Dim InletDia As Double
    Dim MultiOutletDia As Double
'    Dim NozzleLength1 As Double, NozzleLength2 As Double, NozzleLength3 As Double
    Dim Coords(1 To 6) As Double
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim SieSide As Long
    Dim SieBEZ As String
    Dim SieRadius As Double
    Dim SieRadius2 As Double
    Dim SieL As Double
    Dim SieFlanschB As Double
    Dim SieFlanschD As Double
    Dim SieL2 As Double
    Dim SieL3 As Double
    Dim SieWidth As Double
    Dim SieDepth As Double
    
    Dim i As Long
    
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)

    'assign to meaningful variables from the input array
    BasicAsmWidth = arrayOfInputs(2)
    BasicAsmHeight = arrayOfInputs(3)
    BasicAsmLength = arrayOfInputs(4)
    HeatingCoilLength = arrayOfInputs(5)
    MultiOutletLength = arrayOfInputs(6)
    InletDia = arrayOfInputs(7)
    MultiOutletDia = arrayOfInputs(8)
    SieSide = arrayOfInputs(9)
    SieBEZ = arrayOfInputs(10)
    SieRadius = arrayOfInputs(11)
    SieRadius2 = arrayOfInputs(12)
    SieL = arrayOfInputs(13)
    SieFlanschB = arrayOfInputs(14)
    SieFlanschD = arrayOfInputs(15)
    SieL2 = arrayOfInputs(16)
    SieL3 = arrayOfInputs(17)
    SieWidth = arrayOfInputs(18)
    SieDepth = arrayOfInputs(19)
         
    ' Construct Spiralgehäuse:
    Dim objStart As New AutoMath.DPosition
    Dim objP(0 To 3) As AutoMath.DPosition
    Dim objC(0 To 3) As AutoMath.DPosition
    Dim objDir As New AutoMath.DVector
    Dim objTruncCylinder  As Object
    
    objStart.Set 0, 0, 0
    objDir.Set SieWidth, 0, 0
    
    For i = LBound(objC) To UBound(objC)
        Set objC(i) = New AutoMath.DPosition
    Next i
    For i = LBound(objP) To UBound(objP)
        Set objP(i) = New AutoMath.DPosition
    Next i
    
    objP(0).Set 0, 0, -SieRadius
    objP(1).Set 0, 0, -SieRadius - SieL3
    objP(2).Set 0, -SieRadius, -SieRadius - SieL3
    objP(3).Set 0, -SieRadius, 0
    
    Set objC(0) = objP(3).Clone
    objC(1).Set 0, 0, SieRadius
    objC(2).Set 0, SieRadius, 0
    Set objC(3) = objP(0).Clone
    
    Call Spiralgehaeuse(m_outputColl, "Feature_", objStart, objDir, objP, objC, True)
    
    
    ' construct Motor
    objStart.Set 0, SieL2, -SieRadius + SieRadius2
    objDir.Set 1, 0, 0


    Set objTruncCylinder = PlaceTruncatedCylinderByHeights(m_outputColl, objStart, SieL, SieL, objDir, 0, 2# * SieRadius2, True)
    
    m_outputColl.AddOutput "Feature_", objTruncCylinder
    Set objTruncCylinder = Nothing
    
    ' Construct kasten auf Motor
    Dim oBox    As Object
    Dim dblKl As Double
    Dim dblKB As Double
    Dim dblKH As Double
    
    dblKl = SieL * 0.5
    dblKB = SieRadius2
    dblKH = 0.2
    
    objP(0).Set 0.5 * SieL - 0.5 * dblKl, SieL2 + 0.5 * dblKB, -SieRadius + SieRadius2 + SieRadius2
    objP(1).Set 0.5 * SieL + 0.5 * dblKl, SieL2 - 0.5 * dblKB, -SieRadius + SieRadius2 + SieRadius2 + dblKH

    Set oBox = PlaceBox(m_outputColl, objP(0), objP(1))
    m_outputColl.AddOutput "Feature_", oBox
    Set oBox = Nothing

    objStart.Set -0.2, SieL2, -SieRadius + SieRadius2
    objDir.Set -1, 0, 0
    Set objTruncCylinder = PlaceTruncatedCylinderByHeights(m_outputColl, objStart, _
                    0.2, 0.2, objDir, 0, 2# * SieRadius2 * 0.8, True)
    
    m_outputColl.AddOutput "Feature_", objTruncCylinder
    Set objTruncCylinder = Nothing
    
    objStart.Set -0.2, 0, 0
    objDir.Set -1, 0, 0
    Set objTruncCylinder = PlaceTruncatedCylinderByHeights(m_outputColl, objStart, _
                    0.2, 0.2, objDir, 0, 2# * SieRadius * 0.5, True)
    
    m_outputColl.AddOutput "Feature_", objTruncCylinder
    Set objTruncCylinder = Nothing
    
    Dim objPs(0 To 7) As AutoMath.DPosition
    Dim dblRZ As Double
    Dim dblRbreite As Double
    
    For i = LBound(objPs) To UBound(objPs)
        Set objPs(i) = New AutoMath.DPosition
    Next i

    dblRbreite = SieWidth
    dblRZ = -SieRadius - SieL3 - 0.005
    objDir.Set 0, SieRadius + SieRadius2 + SieL2, 0
    
    objPs(0).Set 0, -SieRadius, dblRZ
    objPs(1).Set 0.1, -SieRadius, dblRZ
    objPs(2).Set 0.1, -SieRadius, dblRZ - 0.005
    objPs(3).Set 0.005, -SieRadius, dblRZ - 0.005
    objPs(4).Set 0.005, -SieRadius, dblRZ - 0.1 + 0.005
    objPs(5).Set 0.1, -SieRadius, dblRZ - 0.1 + 0.005
    objPs(6).Set 0.1, -SieRadius, dblRZ - 0.1
    objPs(7).Set 0, -SieRadius, dblRZ - 0.1
    
    Call createProjectLineString(m_outputColl, "Feature_", -1, objPs, objDir, True)
    
    objPs(0).Set dblRbreite - 0, -SieRadius, dblRZ
    objPs(1).Set dblRbreite - 0.1, -SieRadius, dblRZ
    objPs(2).Set dblRbreite - 0.1, -SieRadius, dblRZ - 0.005
    objPs(3).Set dblRbreite - 0.005, -SieRadius, dblRZ - 0.005
    objPs(4).Set dblRbreite - 0.005, -SieRadius, dblRZ - 0.1 + 0.005
    objPs(5).Set dblRbreite - 0.1, -SieRadius, dblRZ - 0.1 + 0.005
    objPs(6).Set dblRbreite - 0.1, -SieRadius, dblRZ - 0.1
    objPs(7).Set dblRbreite - 0, -SieRadius, dblRZ - 0.1
 
    
    Call createProjectLineString(m_outputColl, "Feature_", -1, objPs, objDir, True)

    objDir.Set SieWidth, 0, 0
    
    objPs(0).Set 0, SieRadius2 + SieL2, dblRZ
    objPs(1).Set 0, SieRadius2 + SieL2 - 0.1, dblRZ
    objPs(2).Set 0, SieRadius2 + SieL2 - 0.1, dblRZ - 0.005
    objPs(3).Set 0, SieRadius2 + SieL2 - 0.005, dblRZ - 0.005
    objPs(4).Set 0, SieRadius2 + SieL2 - 0.005, dblRZ - 0.1 + 0.005
    objPs(5).Set 0, SieRadius2 + SieL2 - 0.1, dblRZ - 0.1 + 0.005
    objPs(6).Set 0, SieRadius2 + SieL2 - 0.1, dblRZ - 0.1
    objPs(7).Set 0, SieRadius2 + SieL2, dblRZ - 0.1
    
    Call createProjectLineString(m_outputColl, "Feature_", -1, objPs, objDir, True)

    objDir.Set SieWidth, 0, 0
    
    objPs(0).Set 0, -SieRadius, dblRZ
    objPs(1).Set 0, -SieRadius + 0.1, dblRZ
    objPs(2).Set 0, -SieRadius + 0.1, dblRZ - 0.005
    objPs(3).Set 0, -SieRadius + 0.005, dblRZ - 0.005
    objPs(4).Set 0, -SieRadius + 0.005, dblRZ - 0.1 + 0.005
    objPs(5).Set 0, -SieRadius + 0.1, dblRZ - 0.1 + 0.005
    objPs(6).Set 0, -SieRadius + 0.1, dblRZ - 0.1
    objPs(7).Set 0, -SieRadius, dblRZ - 0.1
    
    Call createProjectLineString(m_outputColl, "Feature_", -1, objPs, objDir, True)


    Set oPart = Nothing
    Set oSymbolFactory = Nothing
    
    LogErrorGer "Physical ended"
    Exit Sub
    
ErrorLabel:
    LogErrorGer "Physical " & Err.Description
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "VVADA definition"
        Debug.Assert False
    Exit Sub
    Resume
    
End Sub

'=========================================================================
'CREATION OF MAINTENANCE ENVELOPE REPRESENTATION OF VVADA
'=========================================================================

Private Sub MaintenanceEnvelopeRep(ByRef arrayOfInputs())
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim BasicAsmWidth As Double, BasicAsmHeight As Double
    Dim BasicAsmLength As Double, HeatingCoilLength As Double
    Dim MultiOutletLength As Double
    Dim InletDia As Double
    Dim MultiOutletDia As Double
    Dim dNozLen1 As Double, dNozLen2 As Double, dNozLen3 As Double
    Dim Coords(1 To 6) As Double
    Dim ii As Integer
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
   
    On Error GoTo ErrorHandler
    
    'assign to meaningful variables from the input array
    BasicAsmWidth = arrayOfInputs(2)
    BasicAsmHeight = arrayOfInputs(3)
    BasicAsmLength = arrayOfInputs(4)
    HeatingCoilLength = arrayOfInputs(5)
    MultiOutletLength = arrayOfInputs(6)
    InletDia = arrayOfInputs(7)
    MultiOutletDia = arrayOfInputs(8)
    
'   Assumption: Nozzle length (Hvac and piping) is taken as EquipmentComponent length /1.7
    dNozLen1 = BasicAsmLength / 1.7
    dNozLen2 = BasicAsmLength / 1.7
    dNozLen3 = BasicAsmLength / 1.7
 
    ' Xmin, Ymin, Zmin and Xmax, Ymax, Zmax coordinates of the envelope box
    
    ' Box dimensions set to (0.1 * 2) times greater than the overall dimensions
    
    Dim totalLength As Double, totalWidth As Double, totalHeight As Double
    
    totalLength = BasicAsmLength + HeatingCoilLength + MultiOutletLength + _
                  dNozLen1 + dNozLen3
    totalWidth = BasicAsmWidth + dNozLen3 * 2#
    totalHeight = BasicAsmHeight
    
    Coords(1) = -(dNozLen1 + totalLength * 0.1)
    Coords(2) = -(dNozLen3 + totalWidth * 0.1)
    Coords(3) = -(totalHeight * 0.1)
    Coords(4) = totalLength - dNozLen1 + totalLength * 0.1
    Coords(5) = totalWidth - dNozLen3 + totalWidth * 0.1
    Coords(6) = totalHeight * 1.1

    ' Get or create the definition
    Dim defColl As IJDDefinitionCollection
    Set defColl = oSymbolFactory.DefinitionCollection(m_outputColl.ResourceManager)
    Dim boxDef As IJDSymbolDefinition
    Dim definitionParams As Variant
    definitionParams = ""
    Set boxDef = defColl.GetDefinitionByProgId(True, "Box.BoxServices", vbNullString, definitionParams)

    Dim env As IMSSymbolEntities.DSymbol
    Set env = oSymbolFactory.PlaceSymbol(boxDef, m_outputColl.ResourceManager)
    Set boxDef = Nothing
    Set defColl = Nothing

    Dim newEnumArg As IJDEnumArgument
    Dim IJEditJDArg As IJDEditJDArgument
    Set newEnumArg = New DEnumArgument
    Set IJEditJDArg = newEnumArg.IJDEditJDArgument

    Dim PC As IJDParameterContent
    Dim argument As IJDArgument
    For ii = 1 To 6
        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.UomValue = Coords(ii)
        PC.Type = igValue
        PC.UomType = 1
        ' Feed the Argument
        argument.index = ii
        argument.Entity = PC
        ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
    Next

    env.IJDValuesArg.SetValues newEnumArg
    Dim IJDInputsArg As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg = env
    IJDInputsArg.Update
    Set IJDInputsArg = Nothing
    Set IJEditJDArg = Nothing
    Set newEnumArg = Nothing

    
    m_outputColl.AddOutput "VVADAEnvelope", env
    Set env = Nothing
    Set oSymbolFactory = Nothing

    Exit Sub
    
ErrorHandler:
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False

End Sub


Private Function createProjectLineString(ByVal objOutputColl As Object, _
                    strName As String, _
                    lngIndex As Long, _
                    objP() As AutoMath.DPosition, _
                    objVec As AutoMath.DVector, _
                    Optional dblClosed As Boolean = True) As Long

    Dim oLine As IngrGeom3D.Line3d
    Dim iElements As IJElements
    Dim complex As IngrGeom3D.ComplexString3d
    Dim Projection As IJProjection
    Dim i As Long
    Dim j As Long
    
    i = LBound(objP)
    
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP(i).x, objP(i).y, objP(i).z, objP(i + 1).x, objP(i + 1).y, objP(i + 1).z)
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    iElements.Add oLine
    Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
    
    For i = i + 1 To UBound(objP)
        If i = UBound(objP) Then
            j = LBound(objP)
        Else
            j = i + 1
        End If
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP(i).x, objP(i).y, objP(i).z, objP(j).x, objP(j).y, objP(j).z)
        complex.AddCurve oLine, True
    Next i
    
    Set Projection = m_GeomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                    complex, objVec.x, objVec.y, objVec.z, objVec.Length, dblClosed)
                                                    
    If lngIndex >= 0 Then
        lngIndex = lngIndex + 1
        objOutputColl.AddOutput strName & Trim$(Str$(lngIndex)), Projection
    Else
        objOutputColl.AddOutput strName, Projection
    End If
    
         
         
End Function

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
    
    If objP4 Is Nothing Then
        
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP3.x, objP3.y, objP3.z, objP1.x, objP1.y, objP1.z)
        complex.AddCurve oLine, True
        
    Else
    
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP3.x, objP3.y, objP3.z, objP4.x, objP4.y, objP4.z)
        complex.AddCurve oLine, True
        
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP4.x, objP4.y, objP4.z, objP1.x, objP1.y, objP1.z)
        complex.AddCurve oLine, True
    
    End If
    
    Set Projection = m_GeomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                    complex, objVec.x, objVec.y, objVec.z, objVec.Length, dblClosed)
                                                    
    If lngIndex >= 0 Then
        lngIndex = lngIndex + 1
        objOutputColl.AddOutput strName & Trim$(Str$(lngIndex)), Projection
    Else
        objOutputColl.AddOutput strName, Projection
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
                    DirDown As AutoMath.DVector) As Long
                    
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
        objP1, objP2, objP2I, objP1I, DirDown)
        
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP2, objP3, objP3I, objP2I, DirDown)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP3, objP4, objP4I, objP3I, DirDown)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP4, objP1, objP1I, objP4I, DirDown)

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


Private Sub CreateVentilator(ByVal objOutputColl As Object, _
                    strName As String, _
                    objP1 As AutoMath.DPosition, _
                    dblRadius As Double, _
                    dblHeight As Double)

    Dim objDir As New AutoMath.DVector
    Dim objTruncCylinder  As Object
    Dim objQ1 As AutoMath.DPosition
    Dim objQ2 As AutoMath.DPosition
    Dim objQ3 As AutoMath.DPosition
    Dim objQ4 As AutoMath.DPosition
 
    Dim dirUp As New AutoMath.DVector
    Dim dirQuer1 As New AutoMath.DVector
    Dim dirQuer2 As New AutoMath.DVector
    Dim i As Long
    
    
    Dim dblVentHeight As Double
    Dim dblVentQuer As Double

    
    objDir.Set 0, 0, 1
    
    ' erzeuge Zylinder
    Set objTruncCylinder = PlaceTruncatedCylinderByHeights(objOutputColl, objP1, dblHeight, dblHeight, objDir, 0, 2# * dblRadius, True)
    
    objOutputColl.AddOutput strName, objTruncCylinder
    Set objTruncCylinder = Nothing
    
    'Place kreuz.
    
    dblVentHeight = 0.005
    dblVentQuer = 0.02 ' 2cm
    
    dirUp.Set 0, 0, 1
    
    For i = 1 To 2
    
        
        If i = 1 Then
            dirQuer1.Set 1, 0, 0
            dirQuer2.Set 0, 1, 0
        Else
            dirQuer1.Set 0, 1, 0
            dirQuer2.Set 1, 0, 0
        End If
        
    
        
        Set objQ1 = vecDir2(objP1, dirUp, dblHeight, dirQuer1, -dblRadius + 0.005, dirQuer2, dblVentQuer * 0.5)
        Set objQ2 = vecDir2(objP1, dirUp, dblHeight, dirQuer1, -dblRadius + 0.005, dirQuer2, -dblVentQuer * 0.5)
        Set objQ3 = vecDir2(objP1, dirUp, dblHeight + dblVentHeight, dirQuer1, -dblRadius + 0.005, dirQuer2, -dblVentQuer * 0.5)
        Set objQ4 = vecDir2(objP1, dirUp, dblHeight + dblVentHeight, dirQuer1, -dblRadius + 0.005, dirQuer2, dblVentQuer * 0.5)
        
        dirQuer1.Length = 2 * (dblRadius - 0.01)
        Call createProjectedLines(objOutputColl, strName, -1, _
            objQ1, objQ2, objQ3, objQ4, dirQuer1, True)
            
    Next i
        
End Sub


Private Sub createLamelle(ByVal objOutputColl As Object, _
                    strName As String, _
                    objStart As AutoMath.DPosition, _
                    dblX As Double, dblY As Double, dblZ As Double)
                    
                    
        Dim ObjBodyA As Object
        Dim objP2 As New AutoMath.DPosition
        
        objP2.Set objStart.x + dblX, objStart.y + dblY, objStart.z + dblZ
        
        Set ObjBodyA = PlaceBox(objOutputColl, objStart, objP2)
        
        m_outputColl.AddOutput "Feature_", ObjBodyA
        
        Set ObjBodyA = Nothing
End Sub

Private Sub Spiralgehaeuse(ByVal objOutputColl As Object, _
                    strName As String, _
                    objStart As AutoMath.DPosition, _
                    objVec As AutoMath.DVector, _
                    objP() As AutoMath.DPosition, _
                    objC() As AutoMath.DPosition, _
                    dblClosed As Boolean)
                    


    Dim oLine As IngrGeom3D.Line3d
    Dim oCurve As IngrGeom3D.Arc3d
    Dim iElements As IJElements
    Dim complex As IngrGeom3D.ComplexString3d
    Dim Projection As IJProjection
    Dim i As Long
    Dim j As Long
    Dim lb As Long
    Dim ub As Long
   
    
    lb = LBound(objP)
    ub = UBound(objP)
    i = lb
    
    Set complex = Nothing
    
    For i = LBound(objP) To UBound(objP) - 1
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP(i).x, objP(i).y, objP(i).z, objP(i + 1).x, objP(i + 1).y, objP(i + 1).z)
        
        If complex Is Nothing Then
            Set iElements = New JObjectCollection ' IMSElements.DynElements
            iElements.Add oLine
            Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
            Set iElements = Nothing
        Else
            complex.AddCurve oLine, True
        End If
    Next i
    
    ' add the arcs
    For i = LBound(objC) To UBound(objC) - 1
    
        Set oCurve = m_GeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                objStart.x, objStart.y, objStart.z, _
                objC(i).x, objC(i).y, objC(i).z, _
                objC(i + 1).x, objC(i + 1).y, objC(i + 1).z)
                
                
        If complex Is Nothing Then
            Set iElements = New JObjectCollection ' IMSElements.DynElements
            iElements.Add oCurve
            Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
            Set iElements = Nothing
                
        Else
            complex.AddCurve oCurve, True
        End If
    Next i
    
    Set Projection = m_GeomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                    complex, objVec.x, objVec.y, objVec.z, objVec.Length, dblClosed)
                                                    
 
    objOutputColl.AddOutput strName, Projection
    
End Sub


Private Sub AddParameter(Inputs() As IMSSymbolEntities.IJDInput, strName As String, Optional varVal)
    
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
    
    Set Inputs(m_lngParameterIndex) = New IMSSymbolEntities.DInput
    Inputs(m_lngParameterIndex).Name = strName
    Inputs(m_lngParameterIndex).Description = strName
    Inputs(m_lngParameterIndex).Properties = igINPUT_IS_A_PARAMETER
    If Not IsMissing(varVal) Then
        PC.UomValue = varVal
        Inputs(m_lngParameterIndex).DefaultParameterValue = PC
    End If
    
    m_lngParameterIndex = m_lngParameterIndex + 1
    
End Sub
