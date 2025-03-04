VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BoxServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    BoxServices.cls
'
'Author
'    RM
'
'Description
'    Definition of Box Symbol
'    The planes are constrcuted in such a way that normals point outwards
'
'Notes
'
'    <notes>
'8 March 2002 Added Operation Aspect JY and NN
'5 September 2002 Added Insulation Aspect MAS
    ' The following subprogram's have been modified
    ' (1) IJDUserSymbolServices_InvokeRepresentation
    ' (2) FeedBoxDefinition
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_outputColl As IJDOutputCollection

Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

  ' Remove all previous Symbol Definition information
  pSymbolDefinition.IJDInputs.RemoveAllInput
  pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
  pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
  
  'To have a nested box symbol support only aspects it's parent supports
  'Leave this line uncommented otherwise comment it out
  'Fix for TR-CP-186212
  pSymbolDefinition.NestedAspectOption = igSYMBOL_NESTED_ASPECTS
  
  'To have a nested box symbol revert to supporting all aspects defined for the box symbol
  'Leave this line uncommented otherwise comment it out
  'After putting this line back in you must also run synchronize
  'pSymbolDefinition.NestedAspectOption = 0
    
  FeedBoxDefinition pSymbolDefinition

End Sub

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "Box.BoxServices"
  End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
  Set oSymbolFactory = Nothing
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  

  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "Box.BoxServices"
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

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    If StrComp(repName, "MaintenanceEnvelope") = 0 Then
        Physical arrayOfInputs
    ElseIf StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
    ElseIf StrComp(repName, "OperationEnvelope") = 0 Then
        Physical arrayOfInputs
    ElseIf StrComp(repName, "Insulation") = 0 Then
        Physical arrayOfInputs
    End If
End Sub

Sub Physical(ByRef arrayOfInputs())

    'The planes are constrcuted in such a way that normals point inwards
    
    Dim Xmin As Double, Ymin As Double, Zmin As Double
    Dim Xmax As Double, Ymax As Double, Zmax As Double
    Dim Points(0 To 11) As Double

    Xmin = arrayOfInputs(1)
    Ymin = arrayOfInputs(2)
    Zmin = arrayOfInputs(3)
    Xmax = arrayOfInputs(4)
    Ymax = arrayOfInputs(5)
    Zmax = arrayOfInputs(6)

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
    Set plane = geomFactory.Planes3d.CreateByPoints(Nothing, 4, Points)
    m_outputColl.AddOutput "Plane1", plane
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
    
    Set plane = geomFactory.Planes3d.CreateByPoints(Nothing, 4, Points)
    m_outputColl.AddOutput "Plane2", plane
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
    Set plane = geomFactory.Planes3d.CreateByPoints(Nothing, 4, Points)
    m_outputColl.AddOutput "Plane3", plane
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
    Set plane = geomFactory.Planes3d.CreateByPoints(Nothing, 4, Points)
    m_outputColl.AddOutput "Plane4", plane
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
    Set plane = geomFactory.Planes3d.CreateByPoints(Nothing, 4, Points)
    m_outputColl.AddOutput "Plane5", plane
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
    Set plane = geomFactory.Planes3d.CreateByPoints(Nothing, 4, Points)
    m_outputColl.AddOutput "Plane6", plane
    Set plane = Nothing
    
    Set geomFactory = Nothing

End Sub

' Feed Box Definition
' Inputs:
'           "Xmin" = 0.1
'           "Ymin" = 0.1
'           "Zmin" = 0.1
'           "Xmax" = 0.2
'           "Ymax" = 0.2
'           "Zmax" = 0.2
' The representations "Physical" and "MaintenaceEnvelop" are set to the definition
' Physical :
'           - Plane1
'           - Plane2
'           - Plane3
'           - Plane4
'           - Plane5
'           - Plane6
' MaintenaceEnvelop :
'           - Plane1
'           - Plane2
'           - Plane3
'           - Plane4
'           - Plane5
'           - Plane6
Private Sub FeedBoxDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

  On Error GoTo ErrorHandler
  
  ' Create a new input by new operator
  Dim Coord(1 To 6) As IMSSymbolEntities.IJDInput
  Dim ii As Integer
  For ii = 1 To 6
      Set Coord(ii) = New IMSSymbolEntities.DInput
  Next
  
  Coord(1).Name = "Xmin"
  Coord(1).Description = "Minimum X Coordinate of the Box"
  Coord(1).Properties = igINPUT_IS_A_PARAMETER
  
  ' Create a defaultValue
  Dim PC As IMSSymbolEntities.IJDParameterContent
  Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
  PC.Type = igValue
  PC.UomValue = 0.1
  Coord(1).DefaultParameterValue = PC
  
  Coord(2).Name = "Ymin"
  Coord(2).Description = "Minimum Y Coordinate of the Box"
  Coord(2).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.1
  Coord(2).DefaultParameterValue = PC

  Coord(3).Name = "Zmin"
  Coord(3).Description = "Minimum Z Coordinate of the Box"
  Coord(3).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.1
  Coord(3).DefaultParameterValue = PC
  
  Coord(4).Name = "Xmax"
  Coord(4).Description = "Maximum X Coordinate of the Box"
  Coord(4).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.2
  Coord(4).DefaultParameterValue = PC
  
  Coord(5).Name = "Ymax"
  Coord(5).Description = "Maximum Y Coordinate of the Box"
  Coord(5).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.2
  Coord(5).DefaultParameterValue = PC
  
  Coord(6).Name = "Zmax"
  Coord(6).Description = "Maximum Z Coordinate of the Box"
  Coord(6).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.2
  Coord(6).DefaultParameterValue = PC

  ' Set the input to the definition
  Dim InputsIf As IMSSymbolEntities.IJDInputs
  Set InputsIf = pSymbolDefinition

  For ii = 1 To 6
    InputsIf.SetInput Coord(ii), ii
  Next

  ' Create the output
  Dim O1 As IMSSymbolEntities.IJDOutput
  Set O1 = New IMSSymbolEntities.DOutput
  
  Dim O2 As IMSSymbolEntities.IJDOutput
  Set O2 = New IMSSymbolEntities.DOutput
  
  Dim O3 As IMSSymbolEntities.IJDOutput
  Set O3 = New IMSSymbolEntities.DOutput
  
  Dim O4 As IMSSymbolEntities.IJDOutput
  Set O4 = New IMSSymbolEntities.DOutput
  
  Dim O5 As IMSSymbolEntities.IJDOutput
  Set O5 = New IMSSymbolEntities.DOutput
  
  Dim O6 As IMSSymbolEntities.IJDOutput
  Set O6 = New IMSSymbolEntities.DOutput
 
  O1.Name = "Plane1"
  O1.Description = "Plane1 of Box"
  O1.Properties = 0
  
  O2.Name = "Plane2"
  O2.Description = "Plane2 of Box"
  O2.Properties = 0
  
  O3.Name = "Plane3"
  O3.Description = "Plane3 of Box"
  O3.Properties = 0
  
  O4.Name = "Plane4"
  O4.Description = "Plane4 of Box"
  O4.Properties = 0
  
  O5.Name = "Plane5"
  O5.Description = "Plane5 of Box"
  O5.Properties = 0
  
  O6.Name = "Plane6"
  O6.Description = "Plane6 of Box"
  O6.Properties = 0
  
  'Define the representation "Physical"
  Dim rep1 As IMSSymbolEntities.IJDRepresentation
  Set rep1 = New IMSSymbolEntities.DRepresentation

  rep1.Name = "Physical"
  rep1.Description = "Physical Representation of Box"
  rep1.Properties = igREPRESENTATION_ISVBFUNCTION
  'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
  'different repIDs available.
  rep1.RepresentationId = SimplePhysical

  Dim oRep1Outputs As IMSSymbolEntities.IJDOutputs
  Set oRep1Outputs = rep1

  ' Set the outputs
  oRep1Outputs.SetOutput O1
  oRep1Outputs.SetOutput O2
  oRep1Outputs.SetOutput O3
  oRep1Outputs.SetOutput O4
  oRep1Outputs.SetOutput O5
  oRep1Outputs.SetOutput O6
  
  ' Set the representation to definition
  Dim RepsIf As IMSSymbolEntities.IJDRepresentations
  Set RepsIf = pSymbolDefinition
  RepsIf.SetRepresentation rep1
  
  
  Dim PhysicalRepEval As IJDRepresentationEvaluation
  Set PhysicalRepEval = New DRepresentationEvaluation
  PhysicalRepEval.Name = "Physical"
  PhysicalRepEval.Description = "Physical representation"
  PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
  PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
  PhysicalRepEval.ProgId = "Box.BoxServices"
  
  Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
  Set RepEvalsIf = pSymbolDefinition

  RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
  
  'Define another representation "MaintenanceEnvelope"
  rep1.Name = "MaintenanceEnvelope"
  rep1.Description = "MaintenanceEnvelope Represntation of the Box"
  rep1.Properties = igREPRESENTATION_ISVBFUNCTION
  rep1.RepresentationId = Maintenance
  
  RepsIf.SetRepresentation rep1
  
  'Define another representation "OperationEnvelope"
  rep1.Name = "OperationEnvelope"
  rep1.Description = "OperationEnvelope Represntation of the Box"
  rep1.Properties = igREPRESENTATION_ISVBFUNCTION
  rep1.RepresentationId = Operation
  
  RepsIf.SetRepresentation rep1
  
  'Define another representation "InsulationEnvelope"
  rep1.Name = "Insulation"
  rep1.Description = "InsulationEnvelope Representation of the Box"
  rep1.Properties = igREPRESENTATION_ISVBFUNCTION
  rep1.RepresentationId = Insulation
  
  RepsIf.SetRepresentation rep1
  
  Set rep1 = Nothing
  Set RepsIf = Nothing
  Set oRep1Outputs = Nothing
  
  ' Set the script associated to the MaintenanceEnvelope representation
  Dim EnvelopeRepEval As IJDRepresentationEvaluation
  
  Set EnvelopeRepEval = New DRepresentationEvaluation
  EnvelopeRepEval.Name = "MaintenanceEnvelope"
  EnvelopeRepEval.Description = "MaintenanceEnvelope representation"
  EnvelopeRepEval.Properties = igREPRESENTATION_HIDDEN
  EnvelopeRepEval.Type = igREPRESENTATION_VBFUNCTION
  EnvelopeRepEval.ProgId = "Box.BoxServices"
  
  RepEvalsIf.AddRepresentationEvaluation EnvelopeRepEval
  
  ' Set the script associated to the OperationEnvelope representation
  Dim OperationRepEval As IJDRepresentationEvaluation
  
  Set OperationRepEval = New DRepresentationEvaluation
  OperationRepEval.Name = "OperationEnvelope"
  OperationRepEval.Description = "OperationEnvelope representation"
  OperationRepEval.Properties = igREPRESENTATION_HIDDEN
  OperationRepEval.Type = igREPRESENTATION_VBFUNCTION
  OperationRepEval.ProgId = "Box.BoxServices"
  
  RepEvalsIf.AddRepresentationEvaluation OperationRepEval

  ' Set the script associated to the InsulationEnvelope representation
  Dim InsulationRepEval As IJDRepresentationEvaluation
  
  Set InsulationRepEval = New DRepresentationEvaluation
  InsulationRepEval.Name = "Insulation"
  InsulationRepEval.Description = "InsulationEnvelope representation"
  InsulationRepEval.Properties = igREPRESENTATION_HIDDEN
  InsulationRepEval.Type = igREPRESENTATION_VBFUNCTION
  InsulationRepEval.ProgId = "Box.BoxServices"
  
  RepEvalsIf.AddRepresentationEvaluation InsulationRepEval
  
  Set PhysicalRepEval = Nothing
  Set EnvelopeRepEval = Nothing
  Set OperationRepEval = Nothing
  Set InsulationRepEval = Nothing
  Set RepEvalsIf = Nothing
  
  Exit Sub

ErrorHandler:
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False

End Sub

