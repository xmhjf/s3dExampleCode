VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlopBotCylServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   Orientation
'   ProgID:         SlopedBotCylinder.SlopBotCylServices
'   Author:         JG
'   Creation Date:  Tuesday, Jan 15
'
'           Creates a sloped bottom cylinder by first calculating the shape of
'           the elipses on the end.  Then creates a Cone3D object with equal height
'           ends.  Then it removes the boundaries from the cone/cylinder and replaces
'           then with the previously generated elipeses.  If the cylinder is to be
'           capped it creates 2 planes by outer boundary useing the elipeses.
'           All the ojects are then moved into place useing the transformation matrix
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------     ---         ------------------
'   19.Feb.2003     JG  Created service
'
'   ******TODO******
'   1:  Add error handing
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_outputColl As IJDOutputCollection

Const MODULE = "SlopedBotCylinder"
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
  IJDUserSymbolServices_GetDefinitionName = "SlopedBotCylinder.SlopBotCylServices"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  Set oSymbolFactory = Nothing
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  

  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SlopedBotCylinder.SlopBotCylServices"
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

Private Function radianToDegree(Radians As Double) As Double


   Dim PI As Double
   
   PI = 4 * Atn(1)
   radianToDegree = 180 / PI * Radians

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   degreeToRadian
'   Author:     JG
'   Inputs:
'               Degree
'
'   Outputs:
'               Radian
'
'   Description:
'               This function converts degrees into radians
'
'   Example of call:
'               Dim dblDegree As Double
'               Dim dblRadian As Double
'               dblRadian = degreeToRadian(dblDegree)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   20.Dec.2002     JG      Created this function.
'   19.Jan.2003     JG      Copied function here to use internally
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function degreeToRadian(degree As Double) As Double

    Const METHOD = "degreeToRadian:"
    On Error GoTo ErrorHandler
    
    Dim PI As Double
    PI = 4 * Atn(1)
    degreeToRadian = (degree * PI) / 180
    
    Exit Function
    
ErrorHandler:
  '  ReportUnanticipatedError MODULE, METHOD
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   aCos
'   Author:     Doug Hempel
'   Inputs:
'               dblInputValue  => value for which arccosine is computed.
'
'   Outputs:
'               Double - The arccosine of the input value. It is an angle in radians.
'
'   Description:
'               Computes the arccosine of the input value.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'   28.Jan.2003     Doug Hempel     Revised the function to handle input of 0 correctly.
'   19.Jan.2003     JG      Copied function here to use internally
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function ACos(ByVal dblInputValue As Double) As Double
  
    Select Case dblInputValue
        Case 0
            ACos = degreeToRadian(90)
        Case -1
            ACos = degreeToRadian(180)
        Case 1
            ACos = 0
        Case Else
            ACos = Atn(-dblInputValue / ((-dblInputValue * dblInputValue + 1) ^ 0.5)) + 2 * Atn(1)
    End Select

End Function

Sub Physical(ByRef arrayOfInputs())

 Const METHOD = "PHYSICAL"
 Dim errMsg As String
 On Error GoTo ErrorLabel
    
    'Used to carry errors to the error handler.
    Dim strError As String

    'Objects that are used to create the sloped
    'bottom cylinder.  We need 2 elipses, 1 ruled surface,
    'and the geometryfactory to build them.
    Dim oElipseWest     As IngrGeom3D.Ellipse3d
    Dim oElipseEast     As IngrGeom3D.Ellipse3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'These are used to build the ellipses
    Dim vecMajAxis      As IJDVector
    Set vecMajAxis = New DVector
    Dim MMRatio         As Double
    Dim dblRadius       As Double
    'These are used to orientate the ellipse
    Dim dblXRotAng      As Double
    Dim dblYRotAng      As Double
   
   'These are used for reference axis
    Dim vecXAxis As IJDVector
    Dim vecYAxis As IJDVector
    Set vecXAxis = New DVector
    Set vecYAxis = New DVector
    vecXAxis.Set 1, 0, 0
    vecYAxis.Set 0, 1, 0
   
   
    'These are the inputs passed in.
    Dim posStartPoint As AutoMath.DPosition
    Set posStartPoint = New DPosition
    
    Dim vecStartVector As AutoMath.DVector
    Set vecStartVector = New DVector
    
    Dim posEndPoint As AutoMath.DPosition
    Set posEndPoint = New DPosition
    
    Dim vecEndVector As AutoMath.DVector
    Set vecEndVector = New DVector
    
    Dim dblDiameter As Double
    Dim bolIsCapped As Boolean
    
    Dim tmxMatrix As IngrGeom3D.IJDT4x4
    Set tmxMatrix = New DT4x4
    
    Dim tmxElipOri As IngrGeom3D.IJDT4x4
    Set tmxElipOri = New DT4x4
    tmxElipOri.LoadIdentity
    
    'Parse the arrayofinputs and set all the variables.
    posStartPoint.Set arrayOfInputs(1), arrayOfInputs(2), arrayOfInputs(3)
    vecStartVector.Set arrayOfInputs(4), arrayOfInputs(5), arrayOfInputs(6)
    posEndPoint.Set arrayOfInputs(7), arrayOfInputs(8), arrayOfInputs(9)
    vecEndVector.Set arrayOfInputs(10), arrayOfInputs(11), arrayOfInputs(12)
    
    dblDiameter = arrayOfInputs(13)
    dblRadius = dblDiameter / 2
    If (arrayOfInputs(14) = 1) Then
        bolIsCapped = True
    Else
        bolIsCapped = False
    End If
    
    Dim i As Integer
    Dim j As Integer
    j = 0
    For i = 15 To 30
        tmxMatrix.IndexValue(j) = arrayOfInputs(i)
        j = j + 1
    Next i
    
    'Start building the sloped bottom cylinder
    
    'Find the angle of the cut.  This is based off the
    'Negative X axis and the normal vector of the elipse.
    Dim vecXInvAxis As IJDVector
    Set vecXInvAxis = New DVector
    vecXInvAxis.Set -1, 0, 0
    'Formula        theta = Acos((vecA . vecB) / |vecA| * |vecB|)
    dblYRotAng = ACos(vecStartVector.Dot(vecXInvAxis) / (Abs(vecStartVector.Length) * Abs(vecXInvAxis.Length)))
    Set vecXInvAxis = Nothing 'We are done with this so go ahead and clean up.
    
    'Find the rotation of the elipse about the x axis
    'starting on the Z axis
    'This orientates the normal of the elipse to its proper
    'Place.  We must check all the cases in which division by/of
    '0 may occure.
    If (vecStartVector.y <> 0 And vecStartVector.z <> 0) Then
        '-1 * tan(y/z)  -1 is to counteract right hand rule.
        'y = opposite z = adjacent
        dblXRotAng = -1 * Atn(vecStartVector.y / vecStartVector.z)
    ElseIf (vecStartVector.y <> 0 And vecStartVector.z = 0) Then
        If vecStartVector.y > 0 Then
            '+Y/0  Z = 0 result = positive Y axis
            dblXRotAng = degreeToRadian(-90) 'Rotate neg due to rt hand rule
        Else
            '-Y/0 Z = 0 result = negative Y axis
            dblXRotAng = degreeToRadian(90)
        End If
    ElseIf (vecStartVector.y = 0 And vecStartVector.z <> 0) Then
        If vecStartVector.z > 0 Then
            '0/+Z Y = 0 result = positive Z axis
            dblXRotAng = degreeToRadian(0)
        Else
            '0/-Z Y = 0 result = negative Z axis
            dblXRotAng = degreeToRadian(180)
        End If
    Else
        'Handle errors cleanly.
        #If INDEBUG Then
            MsgBox "Sloped bottom cylinder: west elipse rotation about X could not be determined."
        #End If
        dblXRotAng = 0
    End If
    
    tmxElipOri.Rotate dblXRotAng, vecXAxis
    tmxElipOri.Rotate dblYRotAng, vecYAxis
                       
    'The Z axis will be our major axis for the elipse
    'And we must include the major axis length with the vector
    'vecMajAxis.Set oriElipAxis.ZAxis.x, oriElipAxis.ZAxis.y, oriElipAxis.ZAxis.z
    vecMajAxis.Set tmxElipOri.IndexValue(8), tmxElipOri.IndexValue(9), tmxElipOri.IndexValue(10)
    vecMajAxis.Length = (dblRadius / Cos(Abs(dblYRotAng)))

    'Minor/Major ratio
    MMRatio = dblRadius / vecMajAxis.Length

    'Create the first elipse.
    Set oElipseWest = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                        posStartPoint.x, posStartPoint.y, posStartPoint.z, _
                                                                        -1 * vecStartVector.x, -1 * vecStartVector.y, -1 * vecStartVector.z, _
                                                                        vecMajAxis.x, vecMajAxis.y, vecMajAxis.z, _
                                                                        MMRatio)
    
'*******************************************************************************************************************
'END ELIPSE 1
'BEGIN ELIPSE 2
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.  
'*******************************************************************************************************************
    'Reset the axis for the 2nd elipse
    tmxElipOri.LoadIdentity
    
    'Find the angle of the cut.  This is based off the
    'positive X axis and the normal vector of the elipse.
    'Formula        theta = Acos((vecA . vecB) / |vecA| * |vecB|)
    dblYRotAng = -1 * ACos(vecEndVector.Dot(vecXAxis) / (Abs(vecEndVector.Length) * Abs(vecXAxis.Length)))
    
    'Find the rotation of the elipse about the x axis
    'starting on the Z axis.
    'This orientates the normal of the elipse to its proper
    'Place.  We must check all the cases in which division by/of
    '0 may occure.
    If (vecEndVector.y <> 0 And vecEndVector.z <> 0) Then
        '-1 * tan(y/z)  -1 is to counteract right hand rule.
        'y = opposite z = adjacent
        dblXRotAng = -1 * Atn(vecEndVector.y / vecEndVector.z)
    ElseIf (vecEndVector.y <> 0 And vecEndVector.z = 0) Then
        If vecEndVector.y > 0 Then
            '+Y/0  Z = 0 result = positive Y axis
            dblXRotAng = degreeToRadian(-90)
        Else
            '-Y/0 Z = 0 result = negative Y axis
            dblXRotAng = degreeToRadian(90)
        End If
    ElseIf (vecEndVector.y = 0 And vecEndVector.z <> 0) Then
        If vecEndVector.z > 0 Then
            '0/+Z Y = 0 result = positive Z axis
            dblXRotAng = degreeToRadian(0)
        Else
            '0/-Z Y = 0 result = negative Z axis
            dblXRotAng = degreeToRadian(180)
        End If
    Else
        'Handle errors cleanly.
        #If INDEBUG Then
            MsgBox "Sloped bottom cylinder: East elipse rotation about X could not be determined."
        #End If
        dblXRotAng = 0
    End If
    
    tmxElipOri.Rotate dblXRotAng, vecXAxis
    tmxElipOri.Rotate dblYRotAng, vecYAxis
    
    'The Z axis will be our major axis for the elipse
    'And we must include the major axis length with the vector
    vecMajAxis.Set tmxElipOri.IndexValue(8), tmxElipOri.IndexValue(9), tmxElipOri.IndexValue(10)
    vecMajAxis.Length = (dblRadius / Cos(Abs(dblYRotAng)))
    
    'Minor/Major ratio
    MMRatio = dblRadius / vecMajAxis.Length

    'Create the first elipse.
    Set oElipseEast = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                        posEndPoint.x, posEndPoint.y, posEndPoint.z, _
                                                                        -1 * vecEndVector.x, -1 * vecEndVector.y, -1 * vecEndVector.z, _
                                                                        vecMajAxis.x, vecMajAxis.y, vecMajAxis.z, _
                                                                        MMRatio)
 
    Dim oSBCY As Cone3d
    Set oSBCY = New Cone3d
    Set oSBCY = geomFactory.Cones3d.CreateBy4Pts(m_outputColl.ResourceManager, _
                                                posStartPoint.x, posStartPoint.y, posStartPoint.z, _
                                                posEndPoint.x, posEndPoint.y, posEndPoint.z, _
                                                posStartPoint.x, posStartPoint.y, dblRadius, _
                                                posEndPoint.x, posEndPoint.y, dblRadius, True)
 
    'Replace the boundaries of the cone with the new boundaries
    
    'Set the east boundary
    Dim oEBoundary As ComplexString3d
    Set oEBoundary = New ComplexString3d
    oEBoundary.AddCurve oElipseEast, True
    oSBCY.SetBoundary 1, oEBoundary

    'Set the west boundary
    Dim oWBoundary As ComplexString3d
    Set oWBoundary = New ComplexString3d
    oWBoundary.AddCurve oElipseWest, False
    oSBCY.SetBoundary 2, oWBoundary
    
    'Move the completed cylinder to its final location
    oSBCY.Transform tmxMatrix

    'Add the cylinder to the outputcollection
    m_outputColl.AddOutput "SBCY", oSBCY
    
    'if the object is capped then we need to add the ends
    If bolIsCapped = True Then
        Dim oCap As IngrGeom3D.Plane3d
        Set oCap = New Plane3d
        
        'Need to switch the normal back to point out
        oElipseWest.SetNormal vecStartVector.x, vecStartVector.y, vecStartVector.z
        
        Set oCap = geomFactory.Planes3d.CreateByOuterBdry(m_outputColl.ResourceManager, _
                                                               oElipseWest)
        
        'Move the cap to the cylinder
        oCap.Transform tmxMatrix
        
        'Add the cap to the output collection
        m_outputColl.AddOutput "WestCap", oCap
                                                               
        'Build the east cap
        
        'Need to switch the normal back to point out.
        oElipseEast.SetNormal vecEndVector.x, vecEndVector.y, vecEndVector.z
        
        Set oCap = geomFactory.Planes3d.CreateByOuterBdry(m_outputColl.ResourceManager, _
                                                               oElipseEast)
                                                                       
                  
        'Move the cap to the cylinder
        oCap.Transform tmxMatrix
        
        'Add the cap to the output collection
        m_outputColl.AddOutput "EastCap", oCap
        
    End If
    
    Exit Sub
    
    'Clean out all the used objects.
    Set oElipseWest = Nothing
    Set oElipseEast = Nothing
    Set geomFactory = Nothing
    Set vecMajAxis = Nothing
    Set vecXAxis = Nothing
    Set vecYAxis = Nothing
    Set posStartPoint = Nothing
    Set posEndPoint = Nothing
    Set vecStartVector = Nothing
    Set vecEndVector = Nothing
    Set tmxMatrix = Nothing
    Set tmxElipOri = Nothing
    Set oSBCY = Nothing
    Set oEBoundary = Nothing
    Set oWBoundary = Nothing
    Set oCap = Nothing
    
    
ErrorLabel:
    
    MsgBox "ERROR: " & errMsg
    
End Sub

Private Sub FeedBoxDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

  On Error GoTo ErrorHandler
  
  ' Create a defaultValue
  Dim PC As IMSSymbolEntities.IJDParameterContent
  Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
  
  ' Create a new input by new operator
  Dim Coord(1 To 30) As IMSSymbolEntities.IJDInput
  Dim ii As Integer
  For ii = 1 To 30
      Set Coord(ii) = New IMSSymbolEntities.DInput
  Next
  
  For ii = 1 To 3

    Coord(ii).Name = "startPos" & ii
    Coord(ii).Description = "startPos" & ii
    Coord(ii).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Coord(ii).DefaultParameterValue = PC

  Next ii
  
  For ii = 4 To 6

    Coord(ii).Name = "startVector" & ii
    Coord(ii).Description = "startVector" & ii
    Coord(ii).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Coord(ii).DefaultParameterValue = PC

  Next ii
  
  For ii = 7 To 9
  
    Coord(ii).Name = "endPos" & ii
    Coord(ii).Description = "endPos" & ii
    Coord(ii).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Coord(ii).DefaultParameterValue = PC

  Next ii
  
  For ii = 10 To 12
    
    Coord(ii).Name = "endVector" & ii
    Coord(ii).Description = "endVector" & ii
    Coord(ii).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Coord(ii).DefaultParameterValue = PC

  Next ii
  
  Coord(13).Name = "Diameter"
  Coord(13).Description = "Diameter"
  Coord(13).Properties = igINPUT_IS_A_PARAMETER
  PC.Type = igValue
  PC.UomValue = 0.1
  Coord(13).DefaultParameterValue = PC
  
  Coord(14).Name = "isCapped"
  Coord(14).Description = "isCapped"
  Coord(14).Properties = igINPUT_IS_A_PARAMETER
  PC.Type = igValue
  PC.UomValue = 0.1
  Coord(14).DefaultParameterValue = PC
 
  ' Add the values of the transformation matrix
  For ii = 15 To 30

    Coord(ii).Name = "matrix" & ii
    Coord(ii).Description = "matrix" & ii
    Coord(ii).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Coord(ii).DefaultParameterValue = PC

  Next ii

  ' Set the input to the definition
  Dim InputsIf As IMSSymbolEntities.IJDInputs
  Set InputsIf = pSymbolDefinition

  For ii = 1 To 30
    InputsIf.SetInput Coord(ii), ii
  Next

  Dim outputTest(1 To 3) As IMSSymbolEntities.IJDOutput

  Set outputTest(1) = New IMSSymbolEntities.DOutput
  outputTest(1).Name = "SBCY"
  outputTest(1).Description = "Body of the SBCY"
  outputTest(1).Properties = 0

  Set outputTest(2) = New IMSSymbolEntities.DOutput
  outputTest(2).Name = "EastCap"
  outputTest(2).Description = "Cap for the east end"
  outputTest(2).Properties = 0

  Set outputTest(3) = New IMSSymbolEntities.DOutput
  outputTest(3).Name = "WestCap"
  outputTest(3).Description = "Cap for the west end"
  outputTest(3).Properties = 0
   
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
  For ii = 1 To 3
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
  PhysicalRepEval.ProgId = "SlopedBotCylinder.SlopBotCylServices"
  
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
  EnvelopeRepEval.ProgId = "SlopedBotCylinder.SlopBotCylServices"
  
  RepEvalsIf.AddRepresentationEvaluation EnvelopeRepEval
  
  ' Set the script associated to the OperationEnvelope representation
  Dim OperationRepEval As IJDRepresentationEvaluation
  
  Set OperationRepEval = New DRepresentationEvaluation
  OperationRepEval.Name = "OperationEnvelope"
  OperationRepEval.Description = "OperationEnvelope representation"
  OperationRepEval.Properties = igREPRESENTATION_HIDDEN
  OperationRepEval.Type = igREPRESENTATION_VBFUNCTION
  OperationRepEval.ProgId = "SlopedBotCylinder.SlopBotCylServices"
  
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


