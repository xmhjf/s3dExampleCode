VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MiterCutCircServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   Orientation
'   ProgID:         MiterCutCircTorus.MiterCutCircServices
'   Author:         JG
'   Creation Date:  Tuesday, Jan 15
'           Creates a segmented circular torus by creating 1 circle and 1 elipse
'           shape.  The position of the first elipse is
'           (SweepAngle / NumberOfSegments) / 2 a ruled surface between the two
'           creates the first Minor Segement.  Then the elipse is rotated
'           (SweepAngle / NumberOfSegments) and a ruledsurface is made to create
'           a major segment.  This is done NumberOfSegments - 2 (end segments) times.
'           Then the last minor segment is created by rotating the circle
'           by SweepAngle and createing a ruled surface from the last elipse and it.
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------     ---         ------------------
'   17.Feb.2003     JG  Created service
'
'   ******TODO******
'   1:  Add error handing
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  19.DEC.2007     PK            Undone the changes made in TR-132021
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_outputColl As IJDOutputCollection

Const MODULE = "Miter"
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
  IJDUserSymbolServices_GetDefinitionName = "MiterCutCircTorus.MiterCutCircServices"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  Set oSymbolFactory = Nothing
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  

  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "MiterCutCircTorus.MiterCutCircServices"
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

Sub Physical(ByRef arrayOfInputs())

 Const METHOD = "PHYSICAL"
 Dim errMsg As String
 On Error GoTo ErrorLabel

    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim oCircle As IngrGeom3D.Circle3d
    Dim oElipse As IngrGeom3D.Ellipse3d
    Dim oElipseRef As IngrGeom3D.Ellipse3d
    
    Set oCircle = New IngrGeom3D.Circle3d
    Set oElipse = New Ellipse3d
    Set oElipseRef = New Ellipse3d
    
    'These are the vars that are either pulled out of the
    'arrayofinputs or calculated from those values
    'They define the creation of the torus.
    Dim dblSweep As Double
    Dim intSeg As Double
    Dim dblRIns As Double
    Dim dblROut As Double
    Dim dblDiameter As Double
    Dim dblRadius As Double
    Dim dblScalingFactor As Double
    Dim dblPartialScaleFactor As Double
    Dim tmxFinalLoc As IJDT4x4
    Set tmxFinalLoc = New DT4x4
    
    'Get sweep angle,number of segments, diameter, and scaleratio
    dblSweep = arrayOfInputs(1)
    intSeg = arrayOfInputs(2)
    dblRIns = arrayOfInputs(3)
    dblROut = arrayOfInputs(4)
    dblScalingFactor = arrayOfInputs(5)
    'We want to break the ratio up into the smaller peices
    'for each segment.  But the first and last segments
    'created are only 1/2 the length of the middle segments
    'so we start by reducing the first ratio by 1/2
    dblPartialScaleFactor = (dblScalingFactor / intSeg) * 0.5
    
    'If the reduction goes to a point we need a very
    'small end peice.
    If dblScalingFactor = 1 Then
        dblScalingFactor = 0.9999 'This is as small as I can make it!
    ElseIf dblScalingFactor > 1 Then
        'Need to switch the sign so the equations will work
        'to increase the size of the mitered torus.
        dblScalingFactor = dblScalingFactor * -1
        dblPartialScaleFactor = dblPartialScaleFactor * -1
    End If
    
    dblDiameter = dblROut - dblRIns
    dblRadius = dblDiameter / 2
    
    'If sweep angle is a multiple of 360 degrees we must add one more segment
    Dim PI As Double
    PI = 4 * Atn(1)
    
    'Createing our own mod function due to
    'VB's function is to accurate.
    Dim intWholeNumber As Integer
    intWholeNumber = (dblSweep / (PI * 2))
    'Now we get the decimal part and check to see if it is 0
    If ((dblSweep / (PI * 2) - intWholeNumber) = 0) Then intSeg = intSeg + 1

    'Retreive the transformation matrix
    Dim i As Integer
    Dim j As Integer
    j = 0
    
    For i = 6 To 21
        tmxFinalLoc.IndexValue(j) = arrayOfInputs(i)
        j = j + 1
    Next i

    'Check wheather to cap the torus or not
    Dim bolIsCapped As Boolean
    If arrayOfInputs(22) = 1 Then
        bolIsCapped = True
    Else
        bolIsCapped = False
    End If
        
    'Create the first circle.
    Dim posCenter As IJDPosition
    Set posCenter = New DPosition
        
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                    0, -1 * (dblRIns + dblRadius), 0, _
                                                                    1, 0, 0, dblRadius)
    
    'Create the first elipse
        
    'Triangleate the points for the angled peice
    Dim dblMajorAngle As Double
    Dim dblMinorAngle As Double
   
    dblMajorAngle = dblSweep / (intSeg - 1)
    dblMinorAngle = dblMajorAngle / 2
    
    'Get the center point of the elipse
    Dim posElipCent As IJDPosition
    Set posElipCent = New DPosition
    Dim dblHypot As Double  'Used for triangulation of points
        
    'Get Center point
    dblHypot = (dblRIns + dblRadius) / Cos(dblMinorAngle)
    posElipCent.Set -1 * (Sin(dblMinorAngle) * dblHypot), -1 * (dblRIns + dblRadius), 0
    
    'Find all the necessary vectors.
    'vecNormal = Normal to the face of the elipse
    'vecFace = the fector that follows the longer part of the elipse
    'vecZAxis = A referece vector to find vecNormal
    Dim vecNormal As IJDVector
    Dim vecFace As IJDVector
    Dim vecZAxis As IJDVector
    
    Set vecNormal = New DVector
    Set vecFace = New DVector
    Set vecZAxis = New DVector

    vecZAxis.Set 0, 0, 1
    vecFace.Set posElipCent.x, posElipCent.y, posElipCent.z
    vecFace.Length = 1  'Get the unit vector
    Set vecNormal = vecFace.Cross(vecZAxis)  'get the normal to the face.

    'Set the magnitude of the face vector = to the radius of the major length.
    vecFace.Length = ((dblDiameter / Cos(dblMinorAngle)) / 2)
    vecFace.Length = vecFace.Length - (vecFace.Length * dblPartialScaleFactor)
    dblDiameter = dblDiameter - (dblDiameter * dblPartialScaleFactor)
    
               
    'We now need the MMRatio.
    Dim MMRatio As Double
    MMRatio = (dblDiameter / (dblDiameter / Cos(dblMinorAngle)))
           
    'Create the elipse's.
    'oElipse:   is used as the lead elipse for the ruled surface.
    'oElipseRef:is rotated one step later and acts as the starting
    '           point for the rulled surfaces.
    
    Set oElipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                        posElipCent.x, posElipCent.y, posElipCent.z, _
                                                                        vecNormal.x, vecNormal.y, vecNormal.z, _
                                                                        vecFace.x, vecFace.y, vecFace.z, _
                                                                        MMRatio)
    
    Set oElipseRef = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                        posElipCent.x, posElipCent.y, posElipCent.z, _
                                                                        vecNormal.x, vecNormal.y, vecNormal.z, _
                                                                        vecFace.x, vecFace.y, vecFace.z, _
                                                                        MMRatio)
                                                                        
    'Build the ruled surface.
    Dim oRuledSurface As IngrGeom3D.RuledSurface3d
    Set oRuledSurface = geomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, oCircle, _
                                                                    oElipseRef, bolIsCapped)
    oRuledSurface.Transform tmxFinalLoc  'Move the peice to its final location
    m_outputColl.AddOutput "Plane0", oRuledSurface
        
    'The following are used to rotate the oElipses to create each major segment
    Dim vecRotation As IJDVector
    Set vecRotation = New DVector
    vecRotation.Set 0, 0, -1
    
    Dim tmxMatrix As IJDT4x4
    Set tmxMatrix = New DT4x4
    tmxMatrix.LoadIdentity
    tmxMatrix.Rotate dblMajorAngle, vecRotation
    'The matrix is loaded with the proper angle to create each major segment
    
    'We must increase the scale ratio due to the following
    'segments are 2 times the length of the first segment thus
    'they have a larger portion of the scale.
    dblPartialScaleFactor = dblPartialScaleFactor * 2
    
    'Create all the major inside segments
    For i = 1 To intSeg - 2

        oElipse.Transform tmxMatrix
        
        'Must scale the segment face before surefaces are created
        vecFace.Length = vecFace.Length - (vecFace.Length * dblPartialScaleFactor)
        dblDiameter = dblDiameter - (dblDiameter * dblPartialScaleFactor)
        MMRatio = (dblDiameter / (dblDiameter / Cos(dblMinorAngle)))
        
        oElipse.MajorRadius = vecFace.Length
        oElipse.MinorMajorRatio = MMRatio
        
        Set oRuledSurface = geomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, oElipseRef, oElipse, False)
        oRuledSurface.Transform tmxFinalLoc
        m_outputColl.AddOutput "Plane_" & i, oRuledSurface
        
        'Scale the referece face to match one in the new location
        oElipseRef.MajorRadius = vecFace.Length
        oElipseRef.MinorMajorRatio = MMRatio
        
        oElipseRef.Transform tmxMatrix 'Move the peice into place
        
    Next i
    
    'Rotate the original circle the full sweep angle to finish the torus.
    tmxMatrix.LoadIdentity
    tmxMatrix.Rotate dblSweep, vecRotation
    oCircle.Transform tmxMatrix
    
    'We must scale the very end circle.
    oCircle.Radius = dblRadius - (dblRadius * dblScalingFactor)
    
    'Create the final 1/2 segment
    Set oRuledSurface = geomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, oElipse, _
                                                                    oCircle, bolIsCapped)
                                                                    
    oRuledSurface.Transform tmxFinalLoc 'Move the peice to its final location
    m_outputColl.AddOutput "Plane1", oRuledSurface

    Set oElipse = Nothing
    Set oCircle = Nothing
    Set oRuledSurface = Nothing
    Set tmxFinalLoc = Nothing
    Set tmxMatrix = Nothing
    Set vecRotation = Nothing
    Set vecNormal = Nothing
    Set vecFace = Nothing
    Set vecZAxis = Nothing
    Set geomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    
'    MsgBox "ERROR: " & errMsg
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

    
End Sub

Private Sub FeedBoxDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

  On Error GoTo ErrorHandler
  
  ' Create a new input by new operator
  Dim Coord(1 To 22) As IMSSymbolEntities.IJDInput
  Dim ii As Integer
  For ii = 1 To 22
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
  
  ' Add the inside radius
  Coord(3).Name = "RadiusIn"
  Coord(3).Description = "Inside radius of the torus"
  Coord(3).Properties = igINPUT_IS_A_PARAMETER
  PC.Type = igValue
  PC.UomValue = 0.1
  Coord(3).DefaultParameterValue = PC
  
  ' Add the outside radius
  Coord(4).Name = "RadiusOut"
  Coord(4).Description = "Outside radius of the torus"
  Coord(4).Properties = igINPUT_IS_A_PARAMETER
  PC.Type = igValue
  PC.UomValue = 0.1
  Coord(4).DefaultParameterValue = PC
  
  ' Add the reduction radius
  Coord(5).Name = "RadiusOut"
  Coord(5).Description = "Outside radius of the torus"
  Coord(5).Properties = igINPUT_IS_A_PARAMETER
  PC.Type = igValue
  PC.UomValue = 0.1
  Coord(5).DefaultParameterValue = PC
  
  ' Add the values of the transformation matrix
  For ii = 6 To 21

    Coord(ii).Name = "matrix" & ii
    Coord(ii).Description = "matrix" & ii
    Coord(ii).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Coord(ii).DefaultParameterValue = PC

  Next ii
  
  Coord(22).Name = "capped"
  Coord(22).Description = "To cap or not to cap the torus"
  Coord(22).Properties = igINPUT_IS_A_PARAMETER
  PC.Type = igValue
  PC.UomValue = 0.1
  Coord(22).DefaultParameterValue = PC
  

  ' Set the input to the definition
  Dim InputsIf As IMSSymbolEntities.IJDInputs
  Set InputsIf = pSymbolDefinition

  For ii = 1 To 22
    InputsIf.SetInput Coord(ii), ii
  Next
    
Dim outputTest(0 To 2) As IMSSymbolEntities.IJDOutput
 
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
  PhysicalRepEval.ProgId = "MiterCutCircTorus.MiterCutCircServices"
  
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
  EnvelopeRepEval.ProgId = "MiterCutCircTorus.MiterCutCircServices"
  
  RepEvalsIf.AddRepresentationEvaluation EnvelopeRepEval
  
  ' Set the script associated to the OperationEnvelope representation
  Dim OperationRepEval As IJDRepresentationEvaluation
  
  Set OperationRepEval = New DRepresentationEvaluation
  OperationRepEval.Name = "OperationEnvelope"
  OperationRepEval.Description = "OperationEnvelope representation"
  OperationRepEval.Properties = igREPRESENTATION_HIDDEN
  OperationRepEval.Type = igREPRESENTATION_VBFUNCTION
  OperationRepEval.ProgId = "MiterCutCircTorus.MiterCutCircServices"
  
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


