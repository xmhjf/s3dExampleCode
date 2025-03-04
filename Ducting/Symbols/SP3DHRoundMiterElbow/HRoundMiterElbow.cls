VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HRoundMiterElbow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (c) 2005, Intergraph Corporation.  All Rights Reserved.
'
' File: HRoundMiterElbow.cls
'
' Author: kkk & svsmylav
' Creation Date:  Friday, Jun 3 2005
'
' Abstract: The Elbow Symbol Class for 3DRouting HVAC Round Duct.
'
' Description:This Class inserts a Round Mitered Elbow (of Fincantieri)symbol. User can change the
' parameter of an Elbow symbol.
'
'History:
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "ElbowServices" ' These 2 variables are for error.
Private Const UNANTICIPATED = 102

Private m_outputColl As imssymbolentities.DOutputCollection

Implements IJDUserSymbolServices

'********************************************************************
' Routine: Class_Terminate
'
' Abstract:The method which is called when this class finishes.
'
' Description:
'
'********************************************************************
Private Sub Class_Terminate()
   Set m_outputColl = Nothing
End Sub

'********************************************************************
' Routine: IJDUserSymbolServices_EditOccurence
'
' Abstract:The method for user to edit parameters of an Elbow symbol
'
' Description:
'
'********************************************************************
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

'********************************************************************
' Routine: IJDUserSymbolServices_InitializeSymbolDefinition
'
' Abstract:The method to initialize symbol definition
'
' Description:This removes the previous definition and sets the new one
' (Input, Output, Representaion, RepresentaionEvaluation)
' This gives default value of its parameters elbow symbol
' using InputSub. This defines graphics to output using OutPutSub.
' It defines Evaluate representation using EvaluateSub.
'
'********************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As imssymbolentities.IJDSymbolDefinition)
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"
On Error GoTo ErrorHandler
   ' Remove all previous Symbol Definition information
   pSymbolDefinition.IJDInputs.RemoveAllInput
   pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
   pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
   
   InputSub pSymbolDefinition
   OutPutSub pSymbolDefinition
   EvaluateSub pSymbolDefinition

'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================
   pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
   
   Exit Sub

ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHRoundMiterElbow.HRoundMiterElbow"
End Function
'********************************************************************
' Routine: IJDUserSymbolServices_InstanciateDefinition
'
' Abstract:The method to instanciate an elbow class object
'
' Description:This instanciates a persistent symbol definition object
' and initialize it for the first time.
'********************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParameters As Variant, ByVal ActiveConnection As Object) As Object
  
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition:"
On Error GoTo ErrorHandler
    Dim oSymbolFactory As New imssymbolentities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As imssymbolentities.IJDSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHRoundMiterElbow.HRoundMiterElbow"
  oSymbolDefinition.CodeBase = CodeBase

  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId
  
  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  Set oSymbolDefinition = Nothing
  Set oSymbolFactory = Nothing

  Exit Function

ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Function

'********************************************************************
' Routine: IJDUserSymbolServices_InvokeRepresentation
'
' Abstract:The method to Invoke Representation
'
' Description: This computes the symbol graphics outputs given the specified
' representation and set of parameters.
'********************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
On Error GoTo ErrorHandler:
    Set m_outputColl = outputcoll
    If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
    ElseIf StrComp(repName, "Insulation") = 0 Then
'        InsulationRep arrayOfInputs
    End If

    Exit Sub

ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: Physical
'
' Abstract:The method to create a Physical graphics object
'
' Description:This creates a Physical graphic object by using
' creatnozzle and createlbow method with given values and fixed
' values.
'********************************************************************
Private Sub Physical(ByRef arrayOfInputs())
   Dim elCenter(3) As Double
   Dim circleCenter(3) As Double
   Dim nozzleCenter(3) As Double
   Dim nozzleOrient(6) As Double
   Dim NozzleLength As Double
   Dim Face1toCenter As Double
   Dim Face2toCenter As Double
   Dim Angle As Double

   Dim elbowRadius As Double
   Dim Width As Double
   Dim Throat As Double
   Dim dElbowRadiusplusThroat As Double
   Dim OutNames As String
 
    Const METHOD = "Physical:"
On Error GoTo ErrorHandler
   
   Dim oPart As PartFacelets.IJDPart
   Set oPart = arrayOfInputs(1)
   
   Dim nozzleIndex As Integer
   
   Width = arrayOfInputs(2)
   elbowRadius = Width / 2
   Throat = arrayOfInputs(3)
   Face1toCenter = arrayOfInputs(4)
   Face2toCenter = arrayOfInputs(5)
   
''Assumed Origin is at intersection of port1 and pot2 axis

   dElbowRadiusplusThroat = Throat + elbowRadius
   
   Dim PI As Double
   PI = 4 * Atn(1)
   Angle = PI / 2

'CreateNozzle 1
'  Length of Nozzle
   NozzleLength = 0

'  Set pt from 0,0,0 to Elbow Center
   elCenter(0) = 0
   elCenter(1) = 0
   elCenter(2) = 0
    
   circleCenter(0) = -dElbowRadiusplusThroat
   circleCenter(1) = 0
   circleCenter(2) = 0
   
'  Create Pipe whose shape is cylinder
   'NozzleCenter(0,0,0)
   Dim dNozzleLength1 As Double
   dNozzleLength1 = Face1toCenter - dElbowRadiusplusThroat
   nozzleCenter(0) = -(dElbowRadiusplusThroat + dNozzleLength1)
   nozzleCenter(1) = 0
   nozzleCenter(2) = 0

   nozzleOrient(0) = -1
   nozzleOrient(1) = 0
   nozzleOrient(2) = 0
   nozzleOrient(3) = 0
   nozzleOrient(4) = 0
   nozzleOrient(5) = 1
  
   OutNames = "HvacNozzle1"
   nozzleIndex = 1
      
   CreateNozzle oPart, nozzleIndex, nozzleCenter, nozzleOrient, OutNames, Width, dNozzleLength1
   
'CreateNozzle 2
   Dim dNozzleLength2 As Double
   dNozzleLength2 = Face2toCenter - dElbowRadiusplusThroat
   nozzleCenter(0) = (dElbowRadiusplusThroat + dNozzleLength2) * Cos(Angle)
   nozzleCenter(1) = (dElbowRadiusplusThroat + dNozzleLength2) * Sin(Angle)
   nozzleCenter(2) = 0

   nozzleOrient(0) = Cos(Angle)
   nozzleOrient(1) = Sin(Angle)
   nozzleOrient(2) = 0
   nozzleOrient(3) = 0
   nozzleOrient(4) = 0
   nozzleOrient(5) = 1
 
   OutNames = "HvacNozzle2"
   nozzleIndex = 2

   CreateNozzle oPart, nozzleIndex, nozzleCenter, nozzleOrient, OutNames, Width, dNozzleLength2
   
'Create Elbow
    Dim ObjRearHeadMiterTop15degBody As Object
    Dim ObjRearHeadMiterTop30degBody As Object
    Dim ObjRearHeadMiterBottom30degBody As Object
    Dim ObjRearHeadMiterBottom15degBody As Object
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New GeometryFactory
    
'''''Insert your code for output 1(Elliptical Front Head Bonnet Body)
    Dim oStPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
    Dim oCenterPoint As New AutoMath.DPosition
    Dim oAxisVec As New AutoMath.DVector
    Dim oEllipticalArc As IngrGeom3D.EllipticalArc3d

    Dim dEllipseMajorAxis As Double
    Dim dMinorToMajorRatio As Double
    Dim dPoints() As Double

    oAxisVec.Set 0, 0, 1

    Dim dMiterRadius As Double
    Dim oEllipse(1 To 6) As IngrGeom3D.Ellipse3d      'Five miter bodies need Six profiles
    Dim dCenterX   As Double
    Dim dCenterY   As Double
    Dim dCenterZ   As Double
    Dim dMajorX As Double
    Dim dMajorY As Double
    Dim dMajorZ As Double
    Dim dNormalX As Double
    Dim dNormalY As Double
    Dim dNormalZ As Double
    
''To Find out Miter Radius.
    dMiterRadius = dElbowRadiusplusThroat
' Insert your code for output 1 (Rear Head Miter Bottom 11.25 degree portion)
    dCenterX = -dMiterRadius    'Face aligned to the center of Miter Bend
    dCenterY = 0
    dCenterZ = 0
        
    dNormalX = -1
    dNormalY = 0
    dNormalZ = 0

    dMajorX = 0
    dMajorY = Width / 2
    dMajorZ = 0
    dMinorToMajorRatio = 1
    
    ''This is a Circle only.
    Set oEllipse(1) = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                        dCenterX, dCenterY, dCenterZ, _
                                                        dNormalX, dNormalY, dNormalZ, _
                                                        dMajorX, dMajorY, dMajorZ, dMinorToMajorRatio)

    dCenterX = -dMiterRadius * (1 - Sin(11.25 * PI / 180))
    dCenterY = dMiterRadius * (1 - Cos(11.25 * PI / 180))
    dCenterZ = 0
    dNormalX = -Cos(11.25 * PI / 180)
    dNormalY = -Sin(11.25 * PI / 180)
    dNormalZ = 0
    
    dEllipseMajorAxis = (0.5 * Width) / Cos(11.25 * PI / 180)
    dMajorX = -dEllipseMajorAxis * Sin(11.25 * PI / 180)
    dMajorY = dEllipseMajorAxis * Cos(11.25 * PI / 180)
    dMajorZ = 0
    dMinorToMajorRatio = Cos(11.25 * PI / 180)
    Set oEllipse(2) = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                        dCenterX, dCenterY, dCenterZ, _
                                                        dNormalX, dNormalY, dNormalZ, _
                                                        dMajorX, dMajorY, dMajorZ, dMinorToMajorRatio)

    Set ObjRearHeadMiterBottom15degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                        oEllipse(1), oEllipse(2), True)
   
'   Set the output
    m_outputColl.AddOutput "ElbowSection1", ObjRearHeadMiterBottom15degBody
    Set ObjRearHeadMiterBottom15degBody = Nothing

' Insert your code for output 9 (Rear Head Miter Bottom 22.5 degree portion)
'    Current segment angle is 22.5 + first segment's angle is 11.25 = 33.75 degrees
    dCenterX = -dMiterRadius * (1 - Sin(33.75 * PI / 180))
    dCenterY = dMiterRadius * (1 - Cos(33.75 * PI / 180))
    dCenterZ = 0
    dNormalX = -Cos(33.75 * PI / 180)
    dNormalY = -Sin(33.75 * PI / 180)
    dNormalZ = 0
    dEllipseMajorAxis = (0.5 * Width) / Cos(11.25 * PI / 180)
    dMajorX = -dEllipseMajorAxis * Sin(33.75 * PI / 180)
    dMajorY = dEllipseMajorAxis * Cos(33.75 * PI / 180)
    dMajorZ = 0
    dMinorToMajorRatio = Cos(11.25 * PI / 180)  'Current segment s angle is 22.5 degrees
    Set oEllipse(3) = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                        dCenterX, dCenterY, dCenterZ, _
                                                        dNormalX, dNormalY, dNormalZ, _
                                                        dMajorX, dMajorY, dMajorZ, dMinorToMajorRatio)

    Set ObjRearHeadMiterBottom30degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                        oEllipse(2), oEllipse(3), True)

'   Set the output
    m_outputColl.AddOutput "ElbowSection2", ObjRearHeadMiterBottom30degBody
    Set ObjRearHeadMiterBottom30degBody = Nothing

' Insert your code for output 10 (Rear Head Miter Middle 22.5 degree portion)
'    Current segment angle is 22.5 + (first two segment's angles are 11.25 + 22.5 =33.75) = 56.25 degrees
    dCenterX = -dMiterRadius * (1 - Sin(56.25 * PI / 180))
    dCenterY = dMiterRadius * (1 - Cos(56.25 * PI / 180))
    dCenterZ = 0
    dNormalX = -Cos(56.25 * PI / 180)
    dNormalY = -Sin(56.25 * PI / 180)
    dNormalZ = 0
    
    dEllipseMajorAxis = (0.5 * Width) / Cos(11.25 * PI / 180)
    dMajorX = -dEllipseMajorAxis * Sin(56.25 * PI / 180)
    dMajorY = dEllipseMajorAxis * Cos(56.25 * PI / 180)
    dMajorZ = 0
    dMinorToMajorRatio = Cos(11.25 * PI / 180) 'Current segment s angle is 22.5 degrees
    
    Set oEllipse(4) = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                        dCenterX, dCenterY, dCenterZ, _
                                                        dNormalX, dNormalY, dNormalZ, _
                                                        dMajorX, dMajorY, dMajorZ, dMinorToMajorRatio)

    Set ObjRearHeadMiterTop30degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                        oEllipse(3), oEllipse(4), True)

'   Set the output
    m_outputColl.AddOutput "ElbowSection3", ObjRearHeadMiterTop30degBody
    Set ObjRearHeadMiterTop30degBody = Nothing

' Insert your code for output 11 (Rear Head Miter Middle 22.5 degree portion)
'    Current segment angle is 22.5 + (first three segment's angles are 11.25 + 22.5 + 22.5 =56.25) = 78.75 degrees
    dCenterX = -dMiterRadius * (1 - Sin(78.75 * PI / 180))
    dCenterY = dMiterRadius * (1 - Cos(78.75 * PI / 180))
    dCenterZ = 0
    dNormalX = -Cos(78.75 * PI / 180)
    dNormalY = -Sin(78.75 * PI / 180)
    dNormalZ = 0
    
    dEllipseMajorAxis = (0.5 * Width) / Cos(11.25 * PI / 180)
    dMajorX = dEllipseMajorAxis * Sin(78.75 * PI / 180)
    dMajorY = -dEllipseMajorAxis * Cos(78.75 * PI / 180)
    dMajorZ = 0
    dMinorToMajorRatio = Cos(11.25 * PI / 180)  'Current segment s angle is 22.5 degrees
    
    Set oEllipse(5) = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                        dCenterX, dCenterY, dCenterZ, _
                                                        dNormalX, dNormalY, dNormalZ, _
                                                        dMajorX, dMajorY, dMajorZ, dMinorToMajorRatio)

    Set ObjRearHeadMiterTop30degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                        oEllipse(4), oEllipse(5), True)

'   Set the output
    m_outputColl.AddOutput "ElbowSection4", ObjRearHeadMiterTop30degBody
    Set ObjRearHeadMiterTop30degBody = Nothing

' Insert your code for output 12 (Rear Head Miter Top 11.25 degree portion)
'    Current segment angle is 90 degrees
    dCenterX = 0
    dCenterY = dMiterRadius
    dCenterZ = 0
    dNormalX = 0
    dNormalY = 1
    dNormalZ = 0
    dMajorX = Width / 2
    dMajorY = 0
    dMajorZ = 0
    dMinorToMajorRatio = 1
    
    ''This is Circle only.
    Set oEllipse(6) = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                        dCenterX, dCenterY, dCenterZ, _
                                                        dNormalX, dNormalY, dNormalZ, _
                                                        dMajorX, dMajorY, dMajorZ, dMinorToMajorRatio)

    Set ObjRearHeadMiterTop15degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, _
                        oEllipse(5), oEllipse(6), True)

'   Set the output
    m_outputColl.AddOutput "ElbowSection5", ObjRearHeadMiterTop15degBody
    Set ObjRearHeadMiterTop15degBody = Nothing
    
    '   Cleanup
    Dim i As Integer
    For i = 1 To 6
        Set oEllipse(i) = Nothing
    Next i

    Set oCenterPoint = Nothing
    Set oAxisVec = Nothing
    Set oGeomFactory = Nothing

    
   Exit Sub
ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: InputSub
'
' Abstract:The method to Inputs the defalut value to displayed parameters.
'
' Description: This defines 5 parameters user can change and set the
' default value.
'********************************************************************
Private Sub InputSub(pSymbolDefinition As IJDSymbolDefinition)
 ' Create a new input by new operator
   Dim Inputs(5) As imssymbolentities.IJDInput
   
   ' Create a defaultValue
   Dim PC As imssymbolentities.DParameterContent
   
    Const METHOD = "InputSub:"
On Error GoTo ErrorHandler
   Set PC = New imssymbolentities.DParameterContent 'not persistent PC
   
   PC.Type = igValue
   
   'Set the input to the definition
   Dim InputsIf As imssymbolentities.IJDInputs
   Set InputsIf = pSymbolDefinition
   
   Dim oSymbolCache As New CustomCache
   oSymbolCache.SetupCustomCache pSymbolDefinition

   Set Inputs(0) = New imssymbolentities.DInput
   Inputs(0).Name = "Width"
   Inputs(0).Description = "Diameter of Elbow"
   Inputs(0).Properties = igINPUT_IS_A_PARAMETER
   PC.UomValue = 0.1
   Inputs(0).DefaultParameterValue = PC
   
   Set Inputs(1) = New imssymbolentities.DInput
   Inputs(1).Name = "ElbowCurveRadius"
   Inputs(1).Description = ""
   Inputs(1).Properties = igINPUT_IS_A_PARAMETER
   PC.UomValue = 0.1
   Inputs(1).DefaultParameterValue = PC
   
   Set Inputs(2) = New imssymbolentities.DInput
   Inputs(2).Name = "Face1toCenter"
   Inputs(2).Description = "Face1toCenter"
   Inputs(2).Properties = igINPUT_IS_A_PARAMETER
   PC.UomValue = 0.86
   Inputs(2).DefaultParameterValue = PC
   
   Set Inputs(3) = New imssymbolentities.DInput
   Inputs(3).Name = "Face2toCenter"
   Inputs(3).Description = "Face2toCenter"
   Inputs(3).Properties = igINPUT_IS_A_PARAMETER
   PC.UomValue = 1.13
   Inputs(3).DefaultParameterValue = PC
      
   Set Inputs(4) = New imssymbolentities.DInput
   Inputs(4).Name = "InsulationThickness"
   Inputs(4).Description = "Insulation Thickness"
   Inputs(4).Properties = igINPUT_IS_A_PARAMETER
   PC.UomValue = 0.025
   Inputs(4).DefaultParameterValue = PC
   
   Dim Index As Integer
   For Index = 0 To 4
      InputsIf.SetInput Inputs(Index), Index + 2
   Next
   
   For Index = 0 To 4 Step 1
      Set Inputs(Index) = Nothing
   Next Index
   Set PC = Nothing
   Exit Sub

ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: OutPutSub
'
' Abstract:The method to define output graphics
'
' Description: This defines the graphics to output.Then call RepSub
' method.
'********************************************************************
Private Sub OutPutSub(pSymbolDefinition As IJDSymbolDefinition)
   Dim i As Integer
   
 'Create the output
   Dim Outs(8) As imssymbolentities.DOutput
   
    Const METHOD = "OutPutSub:"
On Error GoTo ErrorHandler
   For i = 0 To 7 Step 1
      Set Outs(i) = New imssymbolentities.DOutput
      Outs(i).Properties = 0
   Next i

   Outs(0).Name = "HvacNozzle1"
   Outs(0).Description = "PipingPort at End2"

   Outs(1).Name = "HvacNozzle2"
   Outs(1).Description = "PipingPort of End1"
   
   Outs(2).Name = "ElbowSection1"
   Outs(2).Description = "Elbow Section1"
   
   Outs(3).Name = "ElbowSection2"
   Outs(3).Description = "Elbow Section2"
   
   Outs(4).Name = "ElbowSection3"
   Outs(4).Description = "Elbow Section3"
   
   Outs(5).Name = "ElbowSection4"
   Outs(5).Description = "Elbow Section4"
   
   Outs(6).Name = "ElbowSection5"
   Outs(6).Description = "Elbow Section5"
      
   Outs(7).Name = "OutElbowIns"
   Outs(7).Description = "Outside Elbow Insulation"

'This defines Physical, insulation representations.
   RepSub pSymbolDefinition, Outs

   Exit Sub

ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: RepSub
'
' Abstract:The method to define Representation
'
' Description:  This defines 1 representation which is Physical
'
'********************************************************************
Private Sub RepSub(pSymbolDefinition As IJDSymbolDefinition, Outs() As imssymbolentities.DOutput)
   Dim rep1 As imssymbolentities.DRepresentation
   Dim RepsIf As imssymbolentities.IJDRepresentations
   Dim i As Integer

    Const METHOD = "RepSub:"
   On Error GoTo ErrorHandler

   Set rep1 = New imssymbolentities.DRepresentation

   'Define representation "Physical" and set it on the definition
   rep1.Name = "Physical"
   rep1.Description = "Physical Representation of the Elbow"
   rep1.Properties = igREPRESENTATION_ISVBFUNCTION
   'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
   'different repIDs available.
   rep1.RepresentationId = SimplePhysical
   
   Dim oRepPhysicalOutputs As imssymbolentities.IJDOutputs
   Set oRepPhysicalOutputs = rep1
        
   'Set the outputs
   For i = 0 To 6 Step 1
      oRepPhysicalOutputs.SetOutput Outs(i)
      Set Outs(i) = Nothing
   Next i
   
   Set RepsIf = pSymbolDefinition
   RepsIf.SetRepresentation rep1

' Redefine outputs
   oRepPhysicalOutputs.RemoveAllOutput
   Set oRepPhysicalOutputs = Nothing
   
'  Insulation representation
   rep1.Name = "Insulation"
   rep1.Description = "Insulation of the Elbow"
   'Set the repID to Insulation.
   rep1.RepresentationId = Insulation
   rep1.Properties = igREPRESENTATION_ISVBFUNCTION
   
   Dim oRepInsulationOutputs As imssymbolentities.IJDOutputs
   Set oRepInsulationOutputs = rep1
   
   'Set the output
   oRepInsulationOutputs.SetOutput Outs(7)
   Set Outs(7) = Nothing
   Set RepsIf = pSymbolDefinition
   RepsIf.SetRepresentation rep1
   
   Set rep1 = Nothing
   Set oRepPhysicalOutputs = Nothing
   Set oRepInsulationOutputs = Nothing
   Set RepsIf = Nothing
   Exit Sub

ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: EvaluateSub
'
' Abstract:The method to define representaion evaluation.
'
' Description: This defines representaion evaluation.
'********************************************************************
Private Sub EvaluateSub(pSymbolDefinition As IJDSymbolDefinition)
'Define the evaluation for the Physical representation
   Dim PhysicalRepEval As DRepresentationEvaluation
   Dim RepsEvalsIf As imssymbolentities.IJDRepresentationEvaluations

    Const METHOD = "EvaluateSub:"
On Error GoTo ErrorHandler
     
   Set PhysicalRepEval = New DRepresentationEvaluation
   PhysicalRepEval.Name = "Physical"
   PhysicalRepEval.Description = "Physical representation of Elbow"
   PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
   PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
   PhysicalRepEval.ProgId = "SP3DHRoundMiterElbow.HRoundMiterElbow"
   
' Define the evaluation associated to the Insulation representation
  Dim InsulationRepEval As DRepresentationEvaluation
  Set InsulationRepEval = New DRepresentationEvaluation
  InsulationRepEval.Name = "Insulation"
  InsulationRepEval.Description = "Insulation representation"
  InsulationRepEval.Properties = igREPRESENTATION_HIDDEN
  InsulationRepEval.Type = igREPRESENTATION_VBFUNCTION
  InsulationRepEval.ProgId = "SP3DHRoundMiterElbow.HRoundMiterElbow"
   
'  Set the evaluations on symbol definition
   Set RepsEvalsIf = pSymbolDefinition

   RepsEvalsIf.AddRepresentationEvaluation PhysicalRepEval
   RepsEvalsIf.AddRepresentationEvaluation InsulationRepEval
   
   Set RepsEvalsIf = Nothing
   Set PhysicalRepEval = Nothing
   Set InsulationRepEval = Nothing

   Exit Sub

ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
 
'********************************************************************
' Routine: CreateElbow
'
' Abstract:The method to create an elbow.
'
' Description: This creates two elbow surface. Each elbow is created
' by 2 steps. The first one is to create circle. The second step is
' to revolve circle 90 degree around one axis.Each circle has the same
' axis to revolve it. It has the same center point. The difference is
' only diameter.
'********************************************************************
Private Sub CreateElbow(ByRef OrientVec, ByRef elCenter, _
   ByRef OutNames, ByVal Radius As Double, ByVal Angle As Double)
   Dim oGeomFac As New IngrGeom3D.GeometryFactory
   Dim CircleNormal(3) As Double
   Dim RadiusVec(3) As Double
'''   Dim rAng, dAng As Double
   Dim circleCenter(3) As Double
   Dim oElbow As IngrGeom3D.IJRevolution
   Dim AxesRatio As Double
   Dim oEllipse As IngrGeom3D.Ellipse3d
   Dim AxCenter(3)  As Double
   Dim AxNormalVec(3) As Double
   'Dim oTrader As New Trader
   'Dim oUOM As IJUnitsOfMeasure
   Dim oGelem As IJDObject
   Dim i As Integer
   
    Const METHOD = "CreateElbow:"
On Error GoTo ErrorHandler
  
'CONSTRUCTION OF Outside Elbow
   
   For i = 0 To 2 Step 1
      circleCenter(i) = elCenter(i)
      CircleNormal(i) = OrientVec(i)
      RadiusVec(i) = OrientVec(i + 3) * Radius
'      RadiusVec(i) = RadiusVec(i)
   Next i
RadiusVec(2) = -(RadiusVec(2))
   AxesRatio = 1#

   Set oEllipse = oGeomFac.Ellipses3d.CreateByCenterNormMajAxisRatio(m_outputColl.ResourceManager, _
                              circleCenter(0), circleCenter(1), circleCenter(2), _
                              CircleNormal(0), CircleNormal(1), CircleNormal(2), _
                              RadiusVec(0), RadiusVec(1), RadiusVec(2), _
                              AxesRatio)
  
'''   dAng = Angle
   'Set oUOM = oTrader.Service("UnitsOfMeasure", "")
   'Set oTrader = Nothing
   
   'rAng = oUOM.ConvertUnitToUnit(UNIT_ANGLE, dAng, ANGLE_DEGREE, ANGLE_RADIAN)
'''   rAng = ((22#) / (7# * 180#)) * dAng
   'Set oUOM = Nothing
  
   
   AxCenter(0) = circleCenter(0)
   AxCenter(1) = -(circleCenter(0) / Tan(Angle / 2))
   AxCenter(2) = 0#
   
   
   AxNormalVec(0) = 0#
   AxNormalVec(1) = 0#
   AxNormalVec(2) = 1#


   If oElbow Is Nothing Then

      Set oElbow = oGeomFac.Revolutions3d.CreateByCurve _
                 (m_outputColl.ResourceManager, oEllipse, _
                    AxNormalVec(0), AxNormalVec(1), AxNormalVec(2), _
                    AxCenter(0), AxCenter(1), AxCenter(2), Angle, False)
   Else
   End If

   Set oGelem = oEllipse
   Set oEllipse = Nothing
   oGelem.Remove
   Set oGelem = Nothing
   m_outputColl.AddOutput OutNames, oElbow
   Set oElbow = Nothing
   Set oGeomFac = Nothing

   Exit Sub
ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
' Routine: CreateNozzle
'
' Abstract:The method to create a nozzle.
'
' Description: This creates a nozzle.  It has a flange.
'********************************************************************
Private Sub CreateNozzle(ByRef oPart As IJDPart, ByVal nozzleIndex As Integer, _
                            ByRef nozzleCenter, ByRef nozzleOrient, ByRef OutNames, _
                            ByVal Width As Double, NozzleLength As Double)
 
   ''Construction of nozzle
''===========================
    Const METHOD = "CreateNozzle:"
On Error GoTo ErrorHandler

    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oNozzle As GSCADNozzleEntities.IJDNozzle
        Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
        Dim iNozzle As GSCADNozzleEntities.IJDNozzle

    Dim iLogicalDistPort As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
'
'NN: add HVAC parameters for Nozzle
        Dim Depth As Double
        Dim Thickness As Double
        Dim FlangeWidth As Double
'        Dim NozzleLength As Double
        Dim CornerRadius As Double
        Dim PortDepth As Double
        Dim CptOffset As Double
    'assign to meaningful variables from the input array
        Depth = Width

        Dim EndPrep As Long
        Dim FlowDir As DistribFlow
        Dim PortStatus As DistribPortStatus
        Dim DimBaseOuter As Boolean
        Dim strPortID As String
'        Dim iPortIndex As Integer
'
'        iPortIndex = 1
        EndPrep = 11
        CornerRadius = 0#
        PortDepth = 0#
        CptOffset = 0#
    'NozzleLength Has to be GREATER than NozzleFlangeThickness
    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    'Assumption: FlangeWidth and Thickness values are taken approximately from Fincantieri part drawing
    FlangeWidth = 0.11
    Thickness = 0.042
    'NozzleLength = 0
        FlowDir = DistribFlow_BOTH
        PortStatus = DistribPortStatus_BASE
        strPortID = "Hvac" & CStr(nozzleIndex)
        
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(nozzleIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            strPortID, PortDepth, CptOffset, False, m_outputColl.ResourceManager)

'Original Piping Elbow nozzle
     'Set oNozzle = NozzleFactory.CreatePipeNozzleFromPart(oPart, nozzleIndex, _
                                            m_outputColl.ResourceManager)
     Set NozzleFactory = Nothing
          
'
     pos.Set nozzleCenter(0), nozzleCenter(1), nozzleCenter(2)
     'Set iLogicalDistPort = oHvacNozzle
     Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos
     Set pos = Nothing
'
'    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
'    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
     Set dir = New AutoMath.DVector
     dir.Set nozzleOrient(0), nozzleOrient(1), nozzleOrient(2)
     'dir.Set -1, 0, 0
     iDistribPort.SetDirectionVector dir
     Set dir = Nothing
'
     Set iNozzle = oHvacNozzle
     iNozzle.Length = NozzleLength
     m_outputColl.AddOutput OutNames, oHvacNozzle
     Set oHvacNozzle = Nothing
          Set iNozzle = Nothing

     'Set iLogicalDistPort = Nothing
     Set iDistribPort = Nothing

   Exit Sub

ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

'********************************************************************
'Routine:  ReportUnanticipatedError
'Abstract: Helper functions
'Description:
'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub
'********************************************************************
'Private Sub ReportUnanticipatedError(InMethod As String)
'
'    Dim ern As IMSErrorService.JWellKnownErrorNumbers
'    Dim errorService As IMSErrorLog.IJErrorService
'    Dim oTrader As New Trader
'
'    Set errorService = oTrader.Service("ErrorHandler", "")
'
'    ern = errorService.ReportError(Err.Number, MODULE & ":" & InMethod, "UNANTICIPATED", App)
'
'    Select Case ern
'      Case imsAbortApplication:
'            errorService.TerminateApp
'    End Select
'
'    Set errorService = Nothing
'    Set oTrader = Nothing
'
'End Sub

'********************************************************************
' Routine: Insulation
'
' Abstract:The method to create a Insulation graphics object
'
' Description:This creates a Insulation graphic object
'********************************************************************
Private Sub InsulationRep(ByRef arrayOfInputs())
    Dim elCenter(3) As Double
    Dim circleCenter(3) As Double
    Dim nozzleOrient(6) As Double
    Dim Angle As Double
    
    Dim elbowRadius As Double
    Dim Width As Double
    Dim Throat As Double
    Dim parElbowRadiusplusThroat As Double
    Dim OutNames As String
    
    Const METHOD = "Physical:"
    On Error GoTo ErrorHandler
    
    Dim InsulationThickness As Double
    
    Width = arrayOfInputs(2)
    elbowRadius = arrayOfInputs(2) / 2
    Throat = arrayOfInputs(3)
    Angle = arrayOfInputs(4)
    InsulationThickness = arrayOfInputs(5)
    
    Dim parFacetoCenterofTrimmedE As Double
    
    parElbowRadiusplusThroat = Throat + elbowRadius
    parFacetoCenterofTrimmedE = parElbowRadiusplusThroat * Tan(Angle / 2)
    
    'Set pt from 0,0,0 to Elbow Center
    elCenter(0) = 0
    elCenter(1) = 0
    elCenter(2) = 0
    
    circleCenter(0) = -parFacetoCenterofTrimmedE
    circleCenter(1) = 0
    circleCenter(2) = 0
    
    nozzleOrient(0) = -1
    nozzleOrient(1) = 0
    nozzleOrient(2) = 0
    nozzleOrient(3) = 0
    nozzleOrient(4) = 0
    nozzleOrient(5) = 1
    
    'Create Elbow
    OutNames = "OutElbowIns"
    
    '   Prepare insulation envelop's elbow radius
    Dim dInsElbowRadius As Double
    If Throat < InsulationThickness Then
        InsulationThickness = Throat * 0.99
    End If
    dInsElbowRadius = elbowRadius + InsulationThickness
    CreateElbow nozzleOrient, circleCenter, OutNames, dInsElbowRadius, Angle
    
   Exit Sub
ErrorHandler:
  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

