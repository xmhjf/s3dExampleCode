VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BFYLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006-07, Intergraph Corporation. All rights reserved.
'
'   BFYLP.cls
'   ProgID:         SP3DButterFlyValveL.BFYLP
'   Author:         svsmylav
'   Creation Date:  Monday, Apr 24 2006
'   Description:
'       Butter Fly Valve
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   24.Apr.2006     SymbolTeam(India)       DI-94663  New symbol is prepared from existing
'                                           GSCAD symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  1.NOV.2007      RRK  CR-123952 Updated the symbol to use retrive parameters function in basgeom3d.bas
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

' Instance of the class

Private m_OutputColl As Object
Private m_ProgId As String
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices


Private Sub Class_Initialize()
    m_ProgId = "SP3DButterFlyValveL.BFYLP"
End Sub

Private Sub Class_Terminate()
  
  Set m_OutputColl = Nothing

End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParameters As Variant, ByVal ActiveConnection As Object) As Object
    ' This method is in charge of the creation of the symbol definition object
    ' You can keep the current design unchanged
    On Error GoTo ErrorHandler
    
    ' Create the definition from the symbol factory
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
    Set oSymbolFactory = Nothing
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    ' Initialize it
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = m_ProgId
    oSymbolDefinition.CodeBase = CodeBase

    ' Give a unique name to the symbol definition
    oSymbolDefinition.Name = oSymbolDefinition.ProgId
   
    'returned symbol definition
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
    Set oSymbolDefinition = Nothing
  
    Exit Function

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Function



Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DButterFlyValveL.BFYLP"
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
    On Error GoTo ErrorHandler
   
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

        Dim InputsIf As IMSSymbolEntities.IJDInputs
        Set InputsIf = pSymbolDefinition
        
        Dim oSymbolCache As New CustomCache
        oSymbolCache.SetupCustomCache pSymbolDefinition
    
        ' ----------Parameter Input --------------
        Dim Inputs(1 To 10) As IMSSymbolEntities.IJDInput
        
        Set Inputs(1) = New IMSSymbolEntities.DInput
        Inputs(1).Name = "DiscWidth"
        Inputs(1).Description = "Width of the Disc"
        Inputs(1).Properties = igINPUT_IS_A_PARAMETER
    
        ' Create a default parameter
        Dim PC As IMSSymbolEntities.IJDParameterContent
        Set PC = New IMSSymbolEntities.DParameterContent
        PC.Type = igValue
        PC.UomValue = 0.0714
        
        ' Associate a default parameter to the input
        Inputs(1).DefaultParameterValue = PC
            
        
        Set Inputs(2) = New IMSSymbolEntities.DInput
        Inputs(2).Name = "DiscDiameter"
        Inputs(2).Description = "Diameter of the Disc"
        Inputs(2).Properties = igINPUT_IS_A_PARAMETER
        PC.UomValue = 0.243
        Inputs(2).DefaultParameterValue = PC
        
         Set Inputs(3) = New IMSSymbolEntities.DInput
        Inputs(3).Name = "StemWidth"
        Inputs(3).Description = "Width of the Stem"
        Inputs(3).Properties = igINPUT_IS_A_PARAMETER
        PC.UomValue = 0.25
        Inputs(3).DefaultParameterValue = PC
        
         Set Inputs(4) = New IMSSymbolEntities.DInput
        Inputs(4).Name = "StemDiameter"
        Inputs(4).Description = "Diameter of the Stem"
        Inputs(4).Properties = igINPUT_IS_A_PARAMETER
        PC.UomValue = 0.05
        Inputs(4).DefaultParameterValue = PC
        
        Set Inputs(5) = New IMSSymbolEntities.DInput
        Inputs(5).Name = "StemFlWidth"
        Inputs(5).Description = "Width of the Stem Flange"
        Inputs(5).Properties = igINPUT_IS_A_PARAMETER
        PC.UomValue = 0.02
        Inputs(5).DefaultParameterValue = PC
        
         Set Inputs(6) = New IMSSymbolEntities.DInput
        Inputs(6).Name = "StemFlDiameter"
        Inputs(6).Description = "Diameter of the Stem Flange"
        Inputs(6).Properties = igINPUT_IS_A_PARAMETER
        PC.UomValue = 0.152
        Inputs(6).DefaultParameterValue = PC
        
        Set Inputs(7) = New IMSSymbolEntities.DInput
        Inputs(7).Name = "ActFlWidth"
        Inputs(7).Description = "Width of the Actuator Flange"
        Inputs(7).Properties = igINPUT_IS_A_PARAMETER
        PC.UomValue = 0.02
        Inputs(7).DefaultParameterValue = PC
        
        Set Inputs(8) = New IMSSymbolEntities.DInput
        Inputs(8).Name = "ActFlDiameter"
        Inputs(8).Description = "Diameter of the Actuator Flange"
        Inputs(8).Properties = igINPUT_IS_A_PARAMETER
        PC.UomValue = 0.152
        Inputs(8).DefaultParameterValue = PC
        
        Set Inputs(9) = New IMSSymbolEntities.DInput
        Inputs(9).Name = "ActWidth"
        Inputs(9).Description = "Width of the Actuator"
        Inputs(9).Properties = igINPUT_IS_A_PARAMETER
        PC.UomValue = 0.35
        Inputs(9).DefaultParameterValue = PC
        
         Set Inputs(10) = New IMSSymbolEntities.DInput
        Inputs(10).Name = "ActDiameter"
        Inputs(10).Description = "Diameter of the Actuator"
        Inputs(10).Properties = igINPUT_IS_A_PARAMETER
        PC.UomValue = 0.1  '(Inputs(8).DefaultParameterValue) * 0.6
        Inputs(10).DefaultParameterValue = PC
              
              ' ---------Set Input into the definition-
        Dim index As Integer
        
        For index = 1 To 10
            InputsIf.SetInput Inputs(index), index + 1
            Set Inputs(index) = Nothing
        Next
        
        ' -------define Physical representation--
        Dim repPhysical As IMSSymbolEntities.IJDRepresentation
        Set repPhysical = New IMSSymbolEntities.DRepresentation
  
        repPhysical.Name = "Physical"
        repPhysical.Description = "Physical representation"
        repPhysical.Properties = igREPRESENTATION_ISVBFUNCTION
        'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
        'different repIDs available.
        repPhysical.RepresentationId = SimplePhysical

        '-----------Output Example-----------------------
        ' Output of Disc
        Dim Output1 As IMSSymbolEntities.IJDOutput
        Set Output1 = New IMSSymbolEntities.DOutput
    
        Output1.Name = "Disc"
        Output1.Description = "Disc of ButtValve"
        Output1.Properties = 0
        
        ' Output of Stem
        Dim Output2 As IMSSymbolEntities.IJDOutput
        Set Output2 = New IMSSymbolEntities.DOutput
    
        Output2.Name = "Stem"
        Output2.Description = "Stem of ButtValve"
        Output2.Properties = 0
        
        ' Output of Stem Flange
        Dim Output3 As IMSSymbolEntities.IJDOutput
        Set Output3 = New IMSSymbolEntities.DOutput
    
        Output3.Name = "StemFlange"
        Output3.Description = "Stem flange of ButtValve"
        Output3.Properties = 0
        
        ' Output of Actuator flange
        Dim Output4 As IMSSymbolEntities.IJDOutput
        Set Output4 = New IMSSymbolEntities.DOutput
    
        Output4.Name = "ActuatorFlange"
        Output4.Description = "Actuator flange of ButtValve"
        Output4.Properties = 0
        
        ' Output of Actuator
        Dim Output5 As IMSSymbolEntities.IJDOutput
        Set Output5 = New IMSSymbolEntities.DOutput
    
        Output5.Name = "Actuator"
        Output5.Description = "Actuator of ButtValve"
        Output5.Properties = 0
        
        ' Output of Nozzle
        Dim Output6 As IMSSymbolEntities.IJDOutput
        Set Output6 = New IMSSymbolEntities.DOutput
        Output6.Name = "PipePort1"
        Output6.Description = "PipingPort1 of Butterfly valve"
        Output6.Properties = 0
  
        Dim Output7 As IMSSymbolEntities.IJDOutput
        Set Output7 = New IMSSymbolEntities.DOutput
  
        Output7.Name = "PipePort2"
        Output7.Description = "PipingPort2 of Butterfly valve"
        Output7.Properties = 0
        
          
        ' ---------Set Output into the representation -----------
        Dim physicalRepOutputs As IMSSymbolEntities.IJDOutputs
        Set physicalRepOutputs = repPhysical

        physicalRepOutputs.SetOutput Output1
        physicalRepOutputs.SetOutput Output2
        physicalRepOutputs.SetOutput Output3
        physicalRepOutputs.SetOutput Output4
        physicalRepOutputs.SetOutput Output5
        physicalRepOutputs.SetOutput Output6
        physicalRepOutputs.SetOutput Output7
                
        Set Output1 = Nothing
        Set Output2 = Nothing
        Set Output3 = Nothing
        Set Output4 = Nothing
        Set Output5 = Nothing
        Set Output6 = Nothing
        Set Output7 = Nothing

        
        ' ---------Set the representation to definition
        Dim RepsIf As IMSSymbolEntities.IJDRepresentations
        Set RepsIf = pSymbolDefinition
        
        RepsIf.SetRepresentation repPhysical
        
        Set RepsIf = Nothing
        Set repPhysical = Nothing
        Set physicalRepOutputs = Nothing
        
        ' ----------Definition of the function identification corresponding
        ' to the representation named "Physical" --
        Dim PhysicalRepEval As IJDRepresentationEvaluation
        Set PhysicalRepEval = New DRepresentationEvaluation

        PhysicalRepEval.Name = "Physical"   ' Note : The identification name of the evaluation function
                                            ' should be the same as the representation name.
        PhysicalRepEval.Description = "Physical representation"
        PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
        PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
        PhysicalRepEval.ProgId = m_ProgId '

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
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    
    Set m_OutputColl = outputcoll
    If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
    End If

End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    
    ' The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False

End Function

Private Sub Physical(ByRef arrayOfInputs())

    On Error GoTo ErrorLabel
    
    Dim discWidth As Double, discDia As Double
    Dim stemWidth As Double, stemDia As Double
    Dim stemFlWidth As Double, stemFlDia As Double
    Dim actFlWidth As Double, actFlDia As Double
    Dim actWidth As Double, actDia As Double
    Dim stemInWidth As Double
    
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim outname1 As String
    Dim outname2 As String
    Dim outname3 As String
    Dim outname4 As String
    Dim outname5 As String
    Dim outname6 As String
    Dim outname7 As String
    
    outname1 = "Disc"
    outname2 = "Stem"
    outname3 = "StemFlange"
    outname4 = "ActuatorFlange"
    outname5 = "Actuator"
    outname6 = "PipePort1"
    outname7 = "PipePort2"
    
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
  
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    discWidth = arrayOfInputs(2)
    discDia = arrayOfInputs(3)
    stemWidth = arrayOfInputs(4)
    stemDia = arrayOfInputs(5)
    stemFlWidth = arrayOfInputs(6)
    stemFlDia = arrayOfInputs(7)
    actFlWidth = arrayOfInputs(8)
    actFlDia = arrayOfInputs(9)
    actWidth = arrayOfInputs(10)
    actDia = arrayOfInputs(11)
    
    '----------------------- disc body
    ' Construction of Disc Body Cylinder
    Dim circleCenterX As Double, circleCenterY As Double, circleCenterZ As Double
    circleCenterX = -(discWidth / 2)
    circleCenterY = 0
    circleCenterZ = 0
    
    Dim circleNormalX As Double, circleNormalY As Double, circleNormalZ As Double
    circleNormalX = 1
    circleNormalY = 0
    circleNormalZ = 0
    
    ' Construct a circle that will be used to project the disc
    Dim circle1 As IngrGeom3D.Circle3d
    Set circle1 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, circleCenterX, _
    circleCenterY, circleCenterZ, circleNormalX, circleNormalY, circleNormalZ, discDia * 0.5)
    
    Dim projVecX As Double, projVecY As Double, projVecZ As Double
    projVecX = 1
    projVecY = 0
    projVecZ = 0
    Dim projection As IngrGeom3D.Projection3d

    ' Project the disc of body
    Set projection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
    circle1, projVecX, projVecY, projVecZ, discWidth, True)
    
    ' Set the output
    m_OutputColl.AddOutput outname1, projection
    Set projection = Nothing
    
    'Delete the circle that was created for the projection
    Dim gscadElem As IJDObject
    Set gscadElem = circle1
    Set circle1 = Nothing
    gscadElem.Remove
    
    
    '----------------------- stem body
    ' Construction of stem Cylinder, starting where the top of the disc
    circleCenterX = 0
    circleCenterY = (discDia / 2) - stemFlWidth
    
    circleCenterZ = 0
    
    circleNormalX = 0
    circleNormalY = 1
    circleNormalZ = 0
    
    ' Construct a circle that will be used to project the stem
    Set circle1 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, circleCenterX, _
    circleCenterY, circleCenterZ, circleNormalX, circleNormalY, circleNormalZ, stemDia * 0.5)

    ' Project the Stem body
    projVecX = 0
    projVecY = 1
    projVecZ = 0
    stemInWidth = stemWidth - (discDia / 2)
    Set projection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
    circle1, projVecX, projVecY, projVecZ, stemInWidth, True)
    
    ' Set the output
    m_OutputColl.AddOutput outname2, projection
    Set projection = Nothing
    
    'Delete the circle that was created for the projection
    Set gscadElem = circle1
    Set circle1 = Nothing
    gscadElem.Remove

    
    '----------------------- stem flange
    ' Construction of stem flange Cylinder, starting where the disc body left off
    circleCenterX = 0
    circleCenterY = stemWidth - stemFlWidth
    circleCenterZ = 0
    
    ' Construct a circle that will be used to project the stem
    Set circle1 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, circleCenterX, _
    circleCenterY, circleCenterZ, circleNormalX, circleNormalY, circleNormalZ, stemFlDia * 0.5)
    
    ' Project the stem body
    Set projection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
    circle1, projVecX, projVecY, projVecZ, stemFlWidth, True)
    
    ' Set the output
    m_OutputColl.AddOutput outname3, projection
    Set projection = Nothing
    
    'Delete the circle that was created for the projection
    Set gscadElem = circle1
    Set circle1 = Nothing
    gscadElem.Remove
    
    '----------------------- actuator flange
    ' Construction of actuator flange Cylinder, starting where the disc body left off
    circleCenterX = 0
    circleCenterY = stemWidth + 0.002
    circleCenterZ = 0
    
    ' Construct a circle that will be used to project the actuator
    Set circle1 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, circleCenterX, _
    circleCenterY, circleCenterZ, circleNormalX, circleNormalY, circleNormalZ, actFlDia * 0.5)
    
    ' Project the actuator body
    Set projection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
    circle1, projVecX, projVecY, projVecZ, actFlWidth, True)
    
    ' Set the output
    m_OutputColl.AddOutput outname4, projection
    Set projection = Nothing
    
    'Delete the circle that was created for the projection
    Set gscadElem = circle1
    Set circle1 = Nothing
    gscadElem.Remove
    
    
    '----------------------- actuator body
    ' Construction of actuator flange Cylinder, starting where the disc body left off
    circleCenterX = 0
    circleCenterY = stemWidth + actFlWidth + 0.002
    circleCenterZ = 0
    
    ' Construct a circle that will be used to project the actuator
    Set circle1 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, circleCenterX, _
    circleCenterY, circleCenterZ, circleNormalX, circleNormalY, circleNormalZ, actDia * 0.5)
    
    ' Project the actuator body
    Set projection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
    circle1, projVecX, projVecY, projVecZ, actDia, True)
    
    ' Set the output
    m_OutputColl.AddOutput outname5, projection
    Set projection = Nothing
    
    'Delete the circle that was created for the projection
    Set gscadElem = circle1
    Set circle1 = Nothing
    gscadElem.Remove
    
    
    '------------------This is to generate a nozzle


''Construction of nozzle1
''===========================
    Dim oNozzle As GSCADNozzleEntities.IJDNozzle

    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
         
    dir.Set -1, 0, 0
    RetrieveParameters 1, oPart, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    
    Dim xctrpt As Double
    xctrpt = -(discWidth / 2) - cptOffset + depth
    pos.Set xctrpt, 0, 0
    Set oNozzle = CreateNozzle(1, oPart, m_OutputColl, dir, pos)

' Set the output
    m_OutputColl.AddOutput outname6, oNozzle
    Set oNozzle = Nothing
    
'======================
' Creation of nozzle2
'======================
    RetrieveParameters 2, oPart, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
     
    dir.Set 1, 0, 0
    
    xctrpt = (discWidth / 2) + cptOffset - depth
    pos.Set xctrpt, 0, 0

' Set the output
    Set oNozzle = CreateNozzle(2, oPart, m_OutputColl, dir, pos)
    m_OutputColl.AddOutput outname7, oNozzle
    Set oNozzle = Nothing
    Set geomFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing
    Set oPart = Nothing
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing default representation", vbExclamation, "BFYLP definition"
    Debug.Assert False
        Resume Next
    
End Sub
