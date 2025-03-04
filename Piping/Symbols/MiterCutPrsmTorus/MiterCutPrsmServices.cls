VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MiterCutPrsmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   Orientation
'   ProgID:         MiterCutPrsmTorus.MiterCutPrsmServices
'   Author:         JG
'   Creation Date:  Tuesday, Jan 15 2003
'   Description:
'           Creates a segmented torus by creating 2 linestrings in the
'           shape of the profile.  It rotates one by sweepAngle/NumOfSegments
'           and then creates a ruled surface between the two.  Then the other
'           face is rotated to match the first.  This is repeated NumOfSegments
'           times.
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------     ---         ------------------
'   15.Jan.2003     JG  Took code base from boxservices and adapted
'                               the inputs, outputs and physical sections to
'                               create a segmented torus
'
'   20.Jan.2003     JG  Switched the order in which the matrix and
'                               verticies are loaded into the array.  This
'                               is to allow for a future n vertice prismatic
'                               segmented torus.
'
'   23.Jan.2003     JG  Changed the way the the physical object is
'                               created.  Originally the verticies fell along
'                               the curve.  Now the midpoints of the straights
'                               fall along the curve.
'
'   14.Feb.2003     JG  Changed the default orientation of the torus.
'
'   24.Feb.2003     JG  Changed name from MiterCutTorus to MCutPrismTorus
'                               this helps diferentiate it from the circular torus.
'
'   26.Mar.2003     JG  Added ability to specify if the output is to be capped
'                               or not.
'
'   ******TODO******
'   1:  Make inputs variable.  IE.  number of verticies = n vs 4
'   2:  Add error handing
'   3:  Add documentation
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  19.DEC.2007     PK            Undone the changes made in TR-132021
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_outputColl As IJDOutputCollection

Const MODULE = "MCutPrismServices"
Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

  ' Remove all previous Symbol Definition information
  pSymbolDefinition.IJDInputs.RemoveAllInput
  pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
  pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

  FeedBoxDefinition pSymbolDefinition

End Sub

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "MiterCutPrsmTorus.MiterCutPrsmServices"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  Set oSymbolFactory = Nothing
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  

  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "MiterCutPrsmTorus.MiterCutPrsmServices"
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
    End If
End Sub

Sub Physical(ByRef arrayOfInputs())

 Const METHOD = "PHYSICAL"
 On Error GoTo ErrorLabel

    'The planes are constructed in such a way that normals point inwards

    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    'This will be used to rotate the linestrings to
    'create the individual ruled surfaces.
    Dim vecRotation As AutoMath.DVector
    Set vecRotation = New DVector
    vecRotation.Set 0, 0, -1
    Dim tmxMatrix1 As IJDT4x4
    Set tmxMatrix1 = New DT4x4
    tmxMatrix1.LoadIdentity
    
    'Get sweep angle and number of segments
    Dim dblSweep As Double
    dblSweep = arrayOfInputs(1)

    Dim intSeg As Double
    intSeg = arrayOfInputs(2)

    'If sweep angle is 360 degrees we must add one more segment
    If dblSweep >= 6.2831 Then intSeg = intSeg + 1

    'Reconstruct the transformation matrix from the array
    Dim i As Integer
    Dim j As Integer
    Dim tmxMatrix As IJDT4x4
    Set tmxMatrix = New DT4x4
    j = 0
    For i = 3 To 18
        tmxMatrix.IndexValue(j) = arrayOfInputs(i)
        j = j + 1 'Counter into the new matrix
    Next i
    tmxMatrix.Rotate dblSweep, vecRotation 'Move the object to its 45 degree pos
    
    'Get the array of points to create the line string
    Dim arrayPoints(0 To 14) As Double
    j = 0
 
    For i = 19 To 30
        arrayPoints(j) = arrayOfInputs(i)
        j = j + 1 'Counter into the new array
    Next i

    'add the first point to the array to close the loop
    For i = 19 To 21
        arrayPoints(j) = arrayOfInputs(i)
        j = j + 1 'Counter into the new array
    Next i
       
    'Determine if the output is to be capped or not
    Dim blnIsCapped As Boolean
    If arrayOfInputs(31) = 1 Then
        blnIsCapped = True
    Else
        blnIsCapped = False
    End If

    'Create the line strings that will be used for the profile
    'of the torus.
    Dim oLinestr As IngrGeom3D.LineString3d
    Set oLinestr = New IngrGeom3D.LineString3d
    Dim oLinestr2 As IngrGeom3D.LineString3d
    Set oLinestr2 = New IngrGeom3D.LineString3d
    Dim oLinestr3 As IngrGeom3D.LineString3d
    Set oLinestr3 = New IngrGeom3D.LineString3d
    Dim ruledSurface As IngrGeom3D.RuledSurface3d

    'Create the face that is used for both ends.
    'This is the outside face on the minor segments.
    Set oLinestr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, arrayPoints)
 
    'Triangleate the points for the angled peice
    Dim majorAngle As Double
    majorAngle = dblSweep / (intSeg - 1)
   
    Dim dblLengthInside As Double
    Dim dblRadiusIns As Double
    Dim dblRadiusOut As Double
    
    dblRadiusIns = arrayPoints(1)
    dblRadiusOut = arrayPoints(7)
    dblLengthInside = Tan(majorAngle / 2) * dblRadiusIns

    arrayPoints(0) = arrayPoints(0) + dblLengthInside
    arrayPoints(3) = arrayPoints(3) + dblLengthInside
    arrayPoints(12) = arrayPoints(12) + dblLengthInside
    
    Dim dlbLengthOutside As Double
    dlbLengthOutside = (Tan(majorAngle / 2) * (dblRadiusOut - dblRadiusIns)) + dblLengthInside
    arrayPoints(6) = arrayPoints(6) + dlbLengthOutside
    arrayPoints(9) = arrayPoints(9) + dlbLengthOutside
    
    'Build 2 reference faces to be used to create the major segments.
    Set oLinestr2 = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, arrayPoints)
    Set oLinestr3 = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, arrayPoints)

    'Create the first 1/2 or minor segment
    Set ruledSurface = geomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, oLinestr, oLinestr2, blnIsCapped)
   
    ruledSurface.Transform tmxMatrix 'Move the peice into place
    'Add it to the output collection
    i = 0
    m_outputColl.AddOutput "Plane" & i, ruledSurface
     
    'Set oLinestr3 = oLinestr2
    tmxMatrix1.LoadIdentity
    tmxMatrix1.Rotate majorAngle, vecRotation
    
    'Create all the major inside segments
    For i = 1 To intSeg - 2
        
        oLinestr2.Transform tmxMatrix1
        Set ruledSurface = geomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, oLinestr3, oLinestr2, False)
        ruledSurface.Transform tmxMatrix 'Move the peice into place
        oLinestr3.Transform tmxMatrix1
        m_outputColl.AddOutput "Plane_" & i, ruledSurface
    
    Next i
       
    tmxMatrix1.LoadIdentity
    tmxMatrix1.Rotate dblSweep, vecRotation
    oLinestr.Transform tmxMatrix1
    
    'Create the second 1/2 or minor segment
    Set ruledSurface = geomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, oLinestr, oLinestr2, blnIsCapped)
    ruledSurface.Transform tmxMatrix 'Move the piece into place.
    i = i + 1
     m_outputColl.AddOutput "Plane_" & i, ruledSurface
                  
    Set oLinestr = Nothing
    Set oLinestr2 = Nothing
    Set tmxMatrix1 = Nothing
    Set ruledSurface = Nothing
    Set geomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    
   ' ReportUnanticipatedError MODULE, METHOD
    
End Sub

Private Sub FeedBoxDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

  On Error GoTo ErrorHandler
  
  ' Create a new input by new operator
  Dim Coord(1 To 31) As IMSSymbolEntities.IJDInput
  Dim ii As Integer
  For ii = 1 To 31
      Set Coord(ii) = New IMSSymbolEntities.DInput
  Next
  
  ' Add Sweep Angle
  Coord(1).Name = "Sweep"
  Coord(1).Description = "Sweep angle of the torus"
  Coord(1).Properties = igINPUT_IS_A_PARAMETER
  
  ' Create a defaultValue
  Dim PC As IMSSymbolEntities.IJDParameterContent
  Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
  PC.Type = igValue
  PC.UomValue = 0.1
  Coord(1).DefaultParameterValue = PC
  
  ' Add the number of segments
  Coord(2).Name = "NumOfSegs"
  Coord(2).Description = "Number of segments"
  Coord(2).Properties = igINPUT_IS_A_PARAMETER
  PC.Type = igValue
  PC.UomValue = 0.1
  Coord(2).DefaultParameterValue = PC
  
  ' Add the values of the transformation matrix
  For ii = 3 To 18
     
    Coord(ii).Name = "matrix" & ii
    Coord(ii).Description = "matrix" & ii
    Coord(ii).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Coord(ii).DefaultParameterValue = PC
    
  Next ii

  ' Add the coordinates to create the original profile.
  For ii = 19 To 30
  
    Coord(ii).Name = "Cord" & ii
    Coord(ii).Description = "Cord" & ii
    Coord(ii).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Coord(ii).DefaultParameterValue = PC
  
  Next ii

  ' Add the number of segments
  Coord(31).Name = "blnIsCapped"
  Coord(31).Description = "cap the torus?"
  Coord(31).Properties = igINPUT_IS_A_PARAMETER
  PC.Type = igValue
  PC.UomValue = 0.1
  Coord(31).DefaultParameterValue = PC

  ' Set the input to the definition
  Dim InputsIf As IMSSymbolEntities.IJDInputs
  Set InputsIf = pSymbolDefinition

  For ii = 1 To 31
    InputsIf.SetInput Coord(ii), ii
  Next
  
Dim outputTest() As IMSSymbolEntities.IJDOutput
 ReDim outputTest(0 To 2)

  For ii = 0 To 2
   
    Set outputTest(ii) = New IMSSymbolEntities.DOutput
    
    outputTest(ii).Name = "Plane" & ii
    outputTest(ii).Description = "Plane" & ii & " of Box"
    outputTest(ii).Properties = 0
  
  Next ii
   
  'Define the representation "Physical"
  Dim rep1 As IMSSymbolEntities.IJDRepresentation
  Set rep1 = New IMSSymbolEntities.DRepresentation

  rep1.Name = "Physical"
  rep1.Description = "Physical Representation of Box"
  rep1.Properties = igCOLLECTION_VARIABLE ' = igREPRESENTATION_ISVBFUNCTION
  'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
  'different repIDs available.
  rep1.RepresentationId = SimplePhysical

  Dim oRep1Outputs As IMSSymbolEntities.IJDOutputs
  Set oRep1Outputs = rep1

  ' Set the outputs
  For ii = 0 To 2
    oRep1Outputs.SetOutput outputTest(ii)
  Next ii
  
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
  PhysicalRepEval.ProgId = "MiterCutPrsmTorus.MiterCutPrsmServices"
  
  Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
  Set RepEvalsIf = pSymbolDefinition

  RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
  
  'Define another representation "MaintenanceEnvelope"
  rep1.Name = "MaintenanceEnvelope"
  rep1.Description = "MaintenanceEnvelope Represntation of the Box"
  rep1.Properties = igCOLLECTION_VARIABLE '= igREPRESENTATION_ISVBFUNCTION
  rep1.RepresentationId = Maintenance
  
  RepsIf.SetRepresentation rep1
  
  'Define another representation "OperationEnvelope"
  rep1.Name = "OperationEnvelope"
  rep1.Description = "OperationEnvelope Represntation of the Box"
  rep1.Properties = igCOLLECTION_VARIABLE '= igREPRESENTATION_ISVBFUNCTION
  rep1.RepresentationId = Operation
  
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
  EnvelopeRepEval.ProgId = "MiterCutPrsmTorus.MiterCutPrsmServices"
  
  RepEvalsIf.AddRepresentationEvaluation EnvelopeRepEval
  
  ' Set the script associated to the OperationEnvelope representation
  Dim OperationRepEval As IJDRepresentationEvaluation
  
  Set OperationRepEval = New DRepresentationEvaluation
  OperationRepEval.Name = "OperationEnvelope"
  OperationRepEval.Description = "OperationEnvelope representation"
  OperationRepEval.Properties = igREPRESENTATION_HIDDEN
  OperationRepEval.Type = igREPRESENTATION_VBFUNCTION
  OperationRepEval.ProgId = "MiterCutPrsmTorus.MiterCutPrsmServices"
  
  RepEvalsIf.AddRepresentationEvaluation OperationRepEval

  Set PhysicalRepEval = Nothing
  Set EnvelopeRepEval = Nothing
  Set OperationRepEval = Nothing
  Set RepEvalsIf = Nothing
  
  Exit Sub

ErrorHandler:
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False

End Sub


