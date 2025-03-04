VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCTChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    CCTChannel.cls
'   ProgID:         SP3DCableTrayProfile.CCTChannel
'Author
'    NN
'
'Description
'    Definition of Cable Tray Channel section Profile Symbol
'
'Notes
'
'    <notes>
'
'   History
'
'         Similar to HVAC rectangular cross section
'
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
    
    FeedNozzleDefinition pSymbolDefinition

End Sub

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DCableTrayProfile.CCTChannel"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler

  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  Set oSymbolFactory = Nothing
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DCableTrayProfile.CCTChannel"
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
    If StrComp(repName, "Physical") = 0 Then
        Symbolic arrayOfInputs
    End If
End Sub

'Draw a Detailed Cable Tray Channel section
Private Sub Detailed(ByRef arrayOfInputs())
    
End Sub

Private Sub Symbolic(ByRef arrayOfInputs())

    Dim parActualWidth As Double
    Dim parActualDepth As Double
    parActualWidth = arrayOfInputs(1)
    parActualDepth = arrayOfInputs(2)
    
        '*************************************************************************
        '           Cable Tray Channel section
        '*************************************************************************
        ' Construction of Cable Tray Channel section
        '                       Y
        '                       ^
        '                       |
        '            4|         |          |1
        '             |         |          |
        ' ActualDepth |         o------> Z |
        '             |     (0,0,0)        |
        '            3 -------------------- 2
        '               ActualWidth
        ' The Channel section origin is in the middle of the section 0,0,0 of the symbol
        ' Y axis is the radial direction
        ' Channel Depth is along the positive Y axis
        ' Channel Width is parallel to the Z axis
        ' Bottom of Channel section is on the negative side of Y axis
        ' Channel section is normal to the X axis which is the Direction Vector of routing the straight section
        
    Dim CP As New AutoMath.DPosition 'arc center point
    Dim HDepth              As Double
    Dim HWidth              As Double
    Dim dblPnts(0 To 11)  As Double
    Dim ObjCableTrayProfile As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    CP.Set 0, 0, 0
    HDepth = parActualDepth / 2
    HWidth = parActualWidth / 2
'''''   Point 1
''''    dblPnts(0) = CP.x
''''    dblPnts(1) = CP.y + HDepth
''''    dblPnts(2) = CP.z + HWidth
'''''   Point 2
''''    dblPnts(3) = CP.x
''''    dblPnts(4) = CP.y - HDepth
''''    dblPnts(5) = CP.z + HWidth
'''''   Point3
''''    dblPnts(6) = CP.x
''''    dblPnts(7) = CP.y - HDepth
''''    dblPnts(8) = CP.z - HWidth
'''''   Point 4
''''    dblPnts(9) = CP.x
''''    dblPnts(10) = CP.y + HDepth
''''    dblPnts(11) = CP.z - HWidth

'   Point 1
    dblPnts(2) = CP.x
    dblPnts(1) = CP.y + HDepth
    dblPnts(0) = CP.z + HWidth
'   Point 2
    dblPnts(5) = CP.x
    dblPnts(4) = CP.y - HDepth
    dblPnts(3) = CP.z + HWidth
'   Point3
    dblPnts(8) = CP.x
    dblPnts(7) = CP.y - HDepth
    dblPnts(6) = CP.z - HWidth
'   Point 4
    dblPnts(11) = CP.x
    dblPnts(10) = CP.y + HDepth
    dblPnts(9) = CP.z - HWidth

    Set ObjCableTrayProfile = geomFactory.LineStrings3d.CreateByPoints(m_outputColl.ResourceManager, 4, dblPnts)
    Set geomFactory = Nothing
                
        
        m_outputColl.AddOutput "CTChannel", ObjCableTrayProfile
        Set ObjCableTrayProfile = Nothing
            
End Sub
Private Sub MaintenanceCTChannel(ByRef arrayOfInputs())

    Dim parActualWidth As Double
    Dim parActualDepth As Double
    parActualWidth = arrayOfInputs(1)
    parActualDepth = arrayOfInputs(2)
    
        '*************************************************************************
        '           Closed contour of Channel section cable tray
        '*************************************************************************
        ' Construction of Cable Tray CLOSED contour of a rectangular profile
        '                                  5
        '            4 -------------------- 1           Y
        '             |                    |            |
        '             |        .(0,0,0)    |            |
        '             |                    |            O--------->Z
        '             |                    |
        '            3 -------------------- 2
        '
        '
        ' Y axis is the radial direction
        
    Dim CP As New AutoMath.DPosition 'arc center point
    Dim HDepth              As Double
    Dim HWidth              As Double
    Dim dblPnts(0 To 14)  As Double
    Dim ObjCableTrayProfile As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    CP.Set 0, 0, 0
    HDepth = parActualDepth / 2
    HWidth = parActualWidth / 2
'   Point 1
    dblPnts(0) = CP.x
    dblPnts(1) = CP.y + HDepth
    dblPnts(2) = CP.z + HWidth
'   Point 2
    dblPnts(3) = CP.x
    dblPnts(4) = CP.y - HDepth
    dblPnts(5) = CP.z + HWidth
'   Point3
    dblPnts(6) = CP.x
    dblPnts(7) = CP.y - HDepth
    dblPnts(8) = CP.z - HWidth
'   Point 4
    dblPnts(9) = CP.x
    dblPnts(10) = CP.y + HDepth
    dblPnts(11) = CP.z - HWidth
'   Point 5
    dblPnts(12) = CP.x
    dblPnts(13) = CP.y + HDepth
    dblPnts(14) = CP.z + HWidth

    Set ObjCableTrayProfile = geomFactory.LineStrings3d.CreateByPoints(m_outputColl.ResourceManager, 5, dblPnts)
    Set geomFactory = Nothing
                
        
        m_outputColl.AddOutput "CTChannel", ObjCableTrayProfile
        Set ObjCableTrayProfile = Nothing
            
End Sub

' Feed Cable Tray Port Definition
' Inputs:
'           "Actual Width and Depth"
' The representations "Symbolic" and "Detailed" are set to the definition
' Symbolic :
' Detailed :

Sub FeedNozzleDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

  On Error GoTo ErrorHandler
  
  ' Create a new input by new operator
  Dim Inputs(0 To 1) As IMSSymbolEntities.IJDInput
  
  ' Create a defaultValue
  Dim PC As IMSSymbolEntities.DParameterContent
  Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
  PC.Type = igValue

  
  Set Inputs(0) = New IMSSymbolEntities.DInput
  Inputs(0).Name = "ActualWidth"
  Inputs(0).Description = "Actual Width of the Cable Tray"
  Inputs(0).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.5
  Inputs(0).DefaultParameterValue = PC
  
  Set Inputs(1) = New IMSSymbolEntities.DInput
  Inputs(1).Name = "ActualDepth"
  Inputs(1).Description = "Actual Depth of the Cable Tray"
  Inputs(1).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.1
  Inputs(1).DefaultParameterValue = PC
  
  'Set the input to the definition
  Dim InputsIf As IMSSymbolEntities.IJDInputs
  Set InputsIf = pSymbolDefinition

  Dim Index As Integer
  For Index = 0 To 1
    InputsIf.SetInput Inputs(Index), Index + 1
  Next

  'Create the output
  Dim O1 As IMSSymbolEntities.IJDOutput
  Set O1 = New IMSSymbolEntities.DOutput
  
  O1.Name = "CTChannel"
  O1.Description = "Cable Tray Channel section"
  O1.Properties = 0
  

  'Define the representation "Symbolic"
  Dim rep1 As IMSSymbolEntities.IJDRepresentation
  Set rep1 = New IMSSymbolEntities.DRepresentation

  rep1.Name = "Physical"
  rep1.Description = "Representation of Cable Tray Channel section is Symbolic"
  rep1.Properties = igREPRESENTATION_ISVBFUNCTION
  rep1.RepresentationId = SimplePhysical

  Dim oRep1Outputs As IMSSymbolEntities.IJDOutputs
  Set oRep1Outputs = rep1

  'Set the outputs
  oRep1Outputs.SetOutput O1
  Set O1 = Nothing

  'Set the 'symbolic' representation to definition
  Dim RepsIf As IMSSymbolEntities.IJDRepresentations
  Set RepsIf = pSymbolDefinition
  RepsIf.SetRepresentation rep1

  'Define another representation "MaintenanceEnvelope" and set it on the definition
  rep1.Name = "MaintenanceEnvelope"
  rep1.Description = "MaintenanceEnvelope Represntation of the Cable Tray Channel section"
  rep1.Properties = igREPRESENTATION_ISVBFUNCTION
  rep1.RepresentationId = Maintenance

  RepsIf.SetRepresentation rep1

  Set rep1 = Nothing
  Set RepsIf = Nothing
  Set oRep1Outputs = Nothing

  'Define the evaluation for the Symbolic representation
  Dim SymbolicRepEval As IJDRepresentationEvaluation
  Set SymbolicRepEval = New DRepresentationEvaluation
  SymbolicRepEval.Name = "Physical"
  SymbolicRepEval.Description = "Symbolic representation of Cable Tray Channel section"
  SymbolicRepEval.Properties = igREPRESENTATION_HIDDEN
  SymbolicRepEval.Type = igREPRESENTATION_VBFUNCTION
  SymbolicRepEval.ProgId = "SP3DCableTrayProfile.CCTChannel"

  'Define the evaluation for the Detailed representation
  Dim DetailedRepEval As IJDRepresentationEvaluation
  Set DetailedRepEval = New DRepresentationEvaluation
  DetailedRepEval.Name = "MaintenanceEnvelope"
  DetailedRepEval.Description = "Detailed representation of Cable Tray Channel section"
  DetailedRepEval.Properties = igREPRESENTATION_HIDDEN
  DetailedRepEval.Type = igREPRESENTATION_VBFUNCTION
  DetailedRepEval.ProgId = "SP3DCableTrayProfile.CCTChannel"

  'Set the evaluations on symbol definition
  Dim RepsEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
  Set RepsEvalsIf = pSymbolDefinition

  RepsEvalsIf.AddRepresentationEvaluation SymbolicRepEval
  RepsEvalsIf.AddRepresentationEvaluation DetailedRepEval

  Set RepsEvalsIf = Nothing
  Set SymbolicRepEval = Nothing
  Set DetailedRepEval = Nothing

  Exit Sub

ErrorHandler:
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False

End Sub


