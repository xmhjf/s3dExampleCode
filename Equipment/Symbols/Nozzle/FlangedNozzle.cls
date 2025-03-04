VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlangedNozzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 1998-2016, Intergraph Corporation. All rights reserved.
'
'File
'    FlangedNozzle.cls
'
'Author
'    RM
'
'Description
'    Definition of Pipe Nozzle Symbol
'
'Notes
'
'    <notes>
'History
'?? ??? ????       RJ           Creation
'15 Nov 2005    VS        Remove the maintenance aspect.
'
'  08.SEP.2006     KKC    DI-95670  Replace names with initials in all revision history sheets and symbols
'  18.Nov.2009     Haneef  Added insulation aspect to elbow nozzle as a part of 3D-Accessors-Sprint-2009-11
'  22.April.2016   Navateja TR-CP-292817    146 minidump(s) at 'NozzleVBsymb!FlangedNozzle::InsulateFlangeNozzle'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As Object
Private Const E_FAIL = -2147467259
Private Const E_PIPENOZZLE_SYMBOLCREATE = 2147746816#
Private Const TOLERANCE = 0.000001
Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    On Error GoTo ErrorHandler
    IJDUserSymbolServices_EditOccurence = False
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
    
    FeedNozzleDefinition pSymbolDefinition
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  On Error GoTo ErrorHandler
  IJDUserSymbolServices_GetDefinitionName = "GSCADNozzleVBSymb.FlangedNozzle"
  Exit Function
ErrorHandler:
    Err.Raise E_FAIL
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  Set oSymbolFactory = Nothing
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "GSCADNozzleVBSymb.FlangedNozzle"
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

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    On Error GoTo ErrorHandler
    Set m_outputColl = outputcoll
    If StrComp(RepName, "Physical", vbTextCompare) = 0 Then
        Symbolic arrayOfInputs
    ElseIf StrComp(RepName, "Insulation", vbTextCompare) = 0 Then
        InsulateFlangeNozzle arrayOfInputs
    End If
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

Private Sub Symbolic(ByRef arrayOfInputs())
    On Error GoTo ErrorHandler
    Dim flangeDia As Double, flangeThk As Double, pipeDia As Double, pipeLength As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim dCptOffset As Double, dDepth As Double
    
    flangeDia = arrayOfInputs(1)
    flangeThk = arrayOfInputs(2)
    pipeDia = arrayOfInputs(3)
    pipeLength = arrayOfInputs(4)
    dCptOffset = arrayOfInputs(5)
    dDepth = arrayOfInputs(6)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    oErrors.Clear
    
    If (flangeDia <= 0) Or (pipeDia <= 0) Or (flangeThk <= 0) Or (pipeLength < 0) Then
        oErrors.Add E_FAIL, "GSCADNozzleVBsymb.FlangedNozzle", "Nozzle matching dimensional data does not exist. Resetting the values.", "flangeDia <= 0 Or pipeDia <= 0"
        GoTo ErrorHandler
    End If

    'CONSTRUCTION OF  Flange Body Cylinder
    Dim circleCenterX As Double, circleCenterY As Double, circleCenterZ As Double
    circleCenterX = dCptOffset - dDepth
    circleCenterY = 0#
    circleCenterZ = 0#

    Dim circleNormalX As Double, circleNormalY As Double, circleNormalZ As Double
    circleNormalX = 1#
    circleNormalY = 0#
    circleNormalZ = 0#

    Dim circlePointVecX As Double, circlePointVecY As Double, circlePointVecZ As Double
    circlePointVecX = 0
    circlePointVecY = flangeDia * 0.5
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
                                    projVecX, projVecY, projVecZ, flangeThk, True)
    m_outputColl.AddOutput "Flange", projection
    Set projection = Nothing
    Set ellipse = Nothing
    
    'Don't have to create the Pipe Cylinder if PipeLength is zero
    If pipeLength > 0 Then
        'CONSTRUCTION OF  Pipe Cylinder
        circleCenterX = circleCenterX + flangeThk
        circlePointVecY = pipeDia * 0.5
        Set ellipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                    circleCenterX, circleCenterY, circleCenterZ, _
                                    circleNormalX, circleNormalY, circleNormalZ, _
                                    circlePointVecX, circlePointVecY, circlePointVecZ, _
                                    axesRatio)
    
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, ellipse, _
                                        projVecX, projVecY, projVecZ, pipeLength, True)
        m_outputColl.AddOutput "Pipe", projection
        Set projection = Nothing
        Set ellipse = Nothing
    End If

    Set geomFactory = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

' Feed Nozzle Definition
' Inputs:
'           "FlangeDiameter"
'           "FlangeThickness"
'           "PipeDiameter"
'           "PipeLength"
'           "CptOffset"
'           "Depth"
'           "InsulationCriteria"
'           "InsulationThickness"
' The representation "Physical" is set to the definition
' Physical :

Private Sub FeedNozzleDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
  
  On Error GoTo ErrorHandler
  
  ' Create a new input by new operator
  Dim Inputs(0 To 7) As IMSSymbolEntities.IJDInput
  
  ' Create a defaultValue
  Dim PC As IMSSymbolEntities.IJDParameterContent
  Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
  PC.Type = igValue

  Set Inputs(0) = New IMSSymbolEntities.DInput
  Inputs(0).Name = "FlangeDiameter"
  Inputs(0).Description = "Diameter of the Flange"
  Inputs(0).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.2
  Inputs(0).DefaultParameterValue = PC
  
  Set Inputs(1) = New IMSSymbolEntities.DInput
  Inputs(1).Name = "FlangeThickness"
  Inputs(1).Description = "Thickness of the Flange"
  Inputs(1).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.05
  Inputs(1).DefaultParameterValue = PC
  
  Set Inputs(2) = New IMSSymbolEntities.DInput
  Inputs(2).Name = "PipeDiameter"
  Inputs(2).Description = "Diameter of the connector pipe"
  Inputs(2).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.1
  Inputs(2).DefaultParameterValue = PC
  
  Set Inputs(3) = New IMSSymbolEntities.DInput
  Inputs(3).Name = "PipeLength"
  Inputs(3).Description = "Length of the connector pipe"
  Inputs(3).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.25
  Inputs(3).DefaultParameterValue = PC
  
  Set Inputs(4) = New IMSSymbolEntities.DInput
  Inputs(4).Name = "CptOffset"
  Inputs(4).Description = "Connect Point Offset"
  Inputs(4).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0#
  Inputs(4).DefaultParameterValue = PC
  
  Set Inputs(5) = New IMSSymbolEntities.DInput
  Inputs(5).Name = "Depth"
  Inputs(5).Description = "Connect Depth"
  Inputs(5).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0#
  Inputs(5).DefaultParameterValue = PC
  
  Set Inputs(6) = New IMSSymbolEntities.DInput
  Inputs(6).Name = "InsulationCriteria"
  Inputs(6).Description = "Insulation Criteria"
  Inputs(6).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0#
  Inputs(6).DefaultParameterValue = PC
  
  Set Inputs(7) = New IMSSymbolEntities.DInput
  Inputs(7).Name = "InsulationThickness"
  Inputs(7).Description = "Insulation Thickness"
  Inputs(7).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = -1#
  Inputs(7).DefaultParameterValue = PC
  
  'Set the input to the definition
  Dim InputsIf As IMSSymbolEntities.IJDInputs
  Set InputsIf = pSymbolDefinition

  Dim index As Integer
  For index = 0 To 7
    InputsIf.SetInput Inputs(index), index + 1
  Next
  For index = 0 To 7
    Set Inputs(index) = Nothing
  Next

  'Create the output for Simple Physical
  Dim O1 As IMSSymbolEntities.IJDOutput
  Set O1 = New IMSSymbolEntities.DOutput
  
  O1.Name = "Flange"
  O1.Description = "Flange of Nozzle"
  O1.Properties = 0
  
  Dim O2 As IMSSymbolEntities.IJDOutput
  Set O2 = New IMSSymbolEntities.DOutput

  O2.Name = "Pipe"
  O2.Description = "Pipe of Nozzle"
  O2.Properties = 0

    
  'Create the outputs for Insulation Aspect
   Dim O3 As IMSSymbolEntities.IJDOutput
   Set O3 = New IMSSymbolEntities.DOutput

    O3.Name = "InsulatedFlange"
    O3.Description = "Insulated Flange Of Nozzle"
    O3.Properties = 0
    
   Dim O4 As IMSSymbolEntities.DOutput
   Set O4 = New IMSSymbolEntities.DOutput
  
    O4.Name = "InsulatedPipe"
    O4.Description = "Insulated Pipe Of Nozzle"
    O4.Properties = 0
  
  
  'Define the representation "Symbolic"
  Dim rep1 As IMSSymbolEntities.IJDRepresentation
  Set rep1 = New IMSSymbolEntities.DRepresentation

  rep1.Name = "Physical"
  rep1.Description = "Physical Representation of Nozzle"
  rep1.Properties = igREPRESENTATION_ISVBFUNCTION
  rep1.RepresentationId = SimplePhysical

  Dim oRep1Outputs As IMSSymbolEntities.IJDOutputs
  Set oRep1Outputs = rep1

  'Set the outputs
  oRep1Outputs.SetOutput O1
  oRep1Outputs.SetOutput O2
  Set O1 = Nothing
  Set O2 = Nothing

  'Set the 'Physical' representation to definition
  Dim RepsIf As IMSSymbolEntities.IJDRepresentations
  Set RepsIf = pSymbolDefinition
  RepsIf.SetRepresentation rep1

  Set rep1 = Nothing
  Set RepsIf = Nothing
  Set oRep1Outputs = Nothing

  'Define the evaluation for the Physical representation
  Dim SymbolicRepEval As IJDRepresentationEvaluation
  Set SymbolicRepEval = New DRepresentationEvaluation
  SymbolicRepEval.Name = "Physical"
  SymbolicRepEval.Description = "Physical representation of flanged nozzle"
  SymbolicRepEval.Properties = igREPRESENTATION_HIDDEN
  SymbolicRepEval.Type = igREPRESENTATION_VBFUNCTION
  SymbolicRepEval.ProgId = "GSCADNozzleVBsymb.FlangedNozzle"

  'Set the evaluations on symbol definition
  Dim RepsEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
  Set RepsEvalsIf = pSymbolDefinition

  RepsEvalsIf.AddRepresentationEvaluation SymbolicRepEval
  
  Set RepsEvalsIf = Nothing
  Set SymbolicRepEval = Nothing
  
 '**************/* Inslulation Representation Definition Begin*/ **************
 
  'Define the representation "Insulation Symbolic"
   Dim oRep2 As IMSSymbolEntities.IJDRepresentation
   Set oRep2 = New IMSSymbolEntities.DRepresentation
    
    oRep2.Name = "Insulation"
    oRep2.Description = "Insulation Representation of Nozzle"
    oRep2.Properties = igREPRESENTATION_ISVBFUNCTION
    oRep2.RepresentationId = Insulation
    
   Dim oRep2Outputs As IMSSymbolEntities.IJDOutputs
   Set oRep2Outputs = oRep2
    ' Set the Outputs
    oRep2Outputs.SetOutput O3
    oRep2Outputs.SetOutput O4
    Set O3 = Nothing
    Set O4 = Nothing
    
   'set the "Insulation" representation to the definition
   Set RepsIf = pSymbolDefinition
   RepsIf.SetRepresentation oRep2
   
   Set oRep2 = Nothing
   Set RepsIf = Nothing
   Set oRep2Outputs = Nothing
 
   
  'Define the evaluation for the Insulation representation
   
  Set SymbolicRepEval = New DRepresentationEvaluation
  SymbolicRepEval.Name = "Insulation"
  SymbolicRepEval.Description = "Insualtion representation of flanged nozzle"
  SymbolicRepEval.Properties = igREPRESENTATION_HIDDEN
  SymbolicRepEval.Type = igREPRESENTATION_VBFUNCTION
  SymbolicRepEval.ProgId = "GSCADNozzleVBsymb.FlangedNozzle"

  'Set the evaluations on symbol definition
  Set RepsEvalsIf = pSymbolDefinition
  RepsEvalsIf.AddRepresentationEvaluation SymbolicRepEval
  

  Set RepsEvalsIf = Nothing
  Set SymbolicRepEval = Nothing
  
    '**************End Inslulation Representation Definition**************

  Exit Sub

ErrorHandler:
    Error.Raise E_FAIL
End Sub

Private Sub InsulateFlangeNozzle(ByRef arrayOfInputs())
    On Error GoTo ErrorHandler
    Dim dFd As Double  ' Flange Diameter
    Dim dFt As Double 'Flange Thickness
    Dim dPd As Double 'Pipe Diameter
    Dim dPl As Double 'Pipe Length
    Dim dCptOffset As Double
    Dim dDepth As Double ' Flange Depth Thickness
    Dim dInsulationCriteria As Double
    Dim dITh As Double 'Insulation Thickness
    
   
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim oColl As Collection
    Dim iCount As Integer
    Dim PI As Double
    Dim oLinestr As IngrGeom3D.LineString3d
    Dim oGeomServices As IGeometryServices
    
    'Getting the Out Put from the Array
    dFd = arrayOfInputs(1)
    dFt = arrayOfInputs(2)
    dPd = arrayOfInputs(3)
    dPl = arrayOfInputs(4)
    dCptOffset = arrayOfInputs(5)
    dDepth = arrayOfInputs(6)
    dInsulationCriteria = arrayOfInputs(7)
    dITh = arrayOfInputs(8)
    
    
    Set oColl = New Collection
    Set oGeomServices = geomFactory.GeometryServices
    'This is good to get the value of PI use Math call, donot use aTan_1 of 1 * 4
    if not oGeomServices is nothing then
    	PI = oGeomServices.GeomPI
	else
		PI = 3.14159265358979323846264338327950288419716939937511
	EndIf
    
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    oErrors.Clear
    
     If (dFd <= 0) Or (dPd <= 0) Or (dFt <= 0) Or (dPl < 0) Then
        oErrors.Add E_FAIL, "GSCADNozzleVBsymb.FlangedNozzle", "Nozzle matching dimensional data does not exist. Resetting the values.", "flangeDia <= 0 Or pipeDia <= 0 or dInsulationThicknes<=0"
        GoTo ErrorHandler
    End If
    
    'if there is no insulation thick ness then better we exit from here itself so that we
    'will remove all the existing outputs from the symbol machinery minimum of 0.1 mm insulation thickness
    'should exist
    If dITh <= TOLERANCE Then
        'Create a dummy Output and exit
        CreateDummyOutput ((dCptOffset - dDepth) + dFt)
        Exit Sub
    End If
    
    'we donot insulate the nozzle if the insulation criteria is set to none
    If (dInsulationCriteria < 1) Then
        'Create a dummy Output and exit
        CreateDummyOutput ((dCptOffset - dDepth) + dFt)
        Exit Sub
    End If
    
    If (dPl > TOLERANCE) Then
        ' Creating the InsulatePipe Output
        Dim dPipeCornerPts(0 To 14) As Double
        Dim oPipeInsuGeom As IngrGeom3D.Revolution3d
        
        dPipeCornerPts(0) = (dCptOffset - dDepth) + dFt
        dPipeCornerPts(1) = dPd / 2#
        dPipeCornerPts(2) = 0#
        
        dPipeCornerPts(3) = (dCptOffset - dDepth) + dFt
        dPipeCornerPts(4) = (dPd / 2#) + dITh
        dPipeCornerPts(5) = 0#
        
        dPipeCornerPts(6) = (dCptOffset - dDepth) + dFt + dPl
        dPipeCornerPts(7) = (dPd / 2#) + dITh
        dPipeCornerPts(8) = 0#
        
        dPipeCornerPts(9) = (dCptOffset - dDepth) + dFt + dPl
        dPipeCornerPts(10) = dPd / 2#
        dPipeCornerPts(11) = 0#
        
        dPipeCornerPts(12) = dPipeCornerPts(0)
        dPipeCornerPts(13) = dPipeCornerPts(1)
        dPipeCornerPts(14) = dPipeCornerPts(2)
        
        Set oLinestr = Nothing
        Set oLinestr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dPipeCornerPts)
         
        'Create the surface of Revolution using the same line string around X axis
        Set oPipeInsuGeom = geomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                    1#, 0#, 0#, 0#, 0#, 0#, (PI * 2#), True)
        
        'Add the surface to out put of insulation aspect
        m_outputColl.AddOutput "InsulatedPipe", oPipeInsuGeom
    End If
    
    If (dInsulationCriteria > 1) Then ' Insulate the flange also
     'Creating the InsulatedFlange output
      If (dFt > TOLERANCE) Then
          Dim dFlngCornerpts(0 To 20) As Double
          Dim oFlngInsulGeom As IngrGeom3D.Revolution3d
          
          dFlngCornerpts(0) = (dCptOffset - dDepth)
          dFlngCornerpts(1) = dFd / 2#
          dFlngCornerpts(2) = 0#
          
          dFlngCornerpts(3) = (dCptOffset - dDepth)
          dFlngCornerpts(4) = (dFd / 2#) + dITh
          dFlngCornerpts(5) = 0#
          
          dFlngCornerpts(6) = (dCptOffset - dDepth) + dFt + dITh
          dFlngCornerpts(7) = (dFd / 2#) + dITh
          dFlngCornerpts(8) = 0#
          
          dFlngCornerpts(9) = (dCptOffset - dDepth) + dFt + dITh
          dFlngCornerpts(10) = (dPd / 2#) + dITh
          dFlngCornerpts(11) = 0#
          
          dFlngCornerpts(12) = (dCptOffset - dDepth) + dFt
          dFlngCornerpts(13) = (dPd / 2#) + dITh
          dFlngCornerpts(14) = 0#
          
          dFlngCornerpts(15) = (dCptOffset - dDepth) + dFt
          dFlngCornerpts(16) = (dFd / 2#)
          dFlngCornerpts(17) = 0#
          
          dFlngCornerpts(18) = dFlngCornerpts(0)
          dFlngCornerpts(19) = dFlngCornerpts(1)
          dFlngCornerpts(20) = dFlngCornerpts(2)
          
         Set oLinestr = Nothing
         Set oLinestr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dFlngCornerpts)
         'Create the surface of Revolution using the same line string around X axis
         Set oFlngInsulGeom = geomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, oLinestr, _
                                                                            1#, 0#, 0#, _
                                                                         0#, 0#, 0#, (PI * 2#), True)
        'Add the surface to out put of insulation aspect
         m_outputColl.AddOutput "InsulatedFlange", oFlngInsulGeom
      End If
    End If
    Set oGeomServices = Nothing
    Set geomFactory = Nothing
    
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
    m_outputColl.AddOutput "InsulatedPipe", oPoint
Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub
