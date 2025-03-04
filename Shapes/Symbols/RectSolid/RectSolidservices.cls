VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RectSolidservices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    RectSolidServices.cls
'
'Author
'    MU
'
'Description
'    Definition of RectSolidServices Symbol
'History
'   06 Dec  Mule Creation
'   24Jul2003   VS      Do not clear middle tier errors. This may clear out other errors too.
'   15Oct2003   VS      Added outputs for more edges and points
'
'Notes
'
'    <notes>
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection
Private Const E_FAIL = -2147467259
Implements IJDUserSymbolServices

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = "RectSolid.RectSolidServices"
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    
    On Error GoTo ErrorHandler
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    'Feed RectangularSolidDefnition
    ' Inputs:
    '          1. "Part"  ( Catalog part )
    '          2. "RectSolidLength"
    '          3. "RectSolidWidth"
    '          4. "RectSolidHeight"
    ' Representations :
    '           Physical
    
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    ' Create a new input by new operator
    Dim Inputs(1 To 10) As IMSSymbolEntities.DInput
    
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
    
    'Define inputs for RectangularSolid
    Dim Index As Integer
    For Index = 1 To 3
        Set Inputs(Index) = New IMSSymbolEntities.DInput
        Inputs(Index).Properties = igINPUT_IS_A_PARAMETER
        If (Index = 1) Then
            PC.UomValue = 0.1
            Inputs(Index).Name = "RectSolidLength"
            Inputs(Index).Description = "Length of the RectangularSolid"
        ElseIf (Index = 2) Then
            PC.UomValue = 0.06
            Inputs(Index).Name = "RectSolidWidth"
        Else
            PC.UomValue = 0.08
            Inputs(Index).Name = "RectSolidHeight"
        End If
        Inputs(Index).DefaultParameterValue = PC
    Next
    
    For Index = 1 To 3
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
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
      
      Dim O7 As IMSSymbolEntities.IJDOutput
      Set O7 = New IMSSymbolEntities.DOutput
      
      Dim O8 As IMSSymbolEntities.IJDOutput
      Set O8 = New IMSSymbolEntities.DOutput

      Dim O9 As IMSSymbolEntities.IJDOutput
      Set O9 = New IMSSymbolEntities.DOutput
      
      Dim O10 As IMSSymbolEntities.IJDOutput
      Set O10 = New IMSSymbolEntities.DOutput
      
      Dim O11 As IMSSymbolEntities.IJDOutput
      Set O11 = New IMSSymbolEntities.DOutput
            
      Dim O12 As IMSSymbolEntities.IJDOutput
      Set O12 = New IMSSymbolEntities.DOutput
      
      Dim E1 As IMSSymbolEntities.IJDOutput
      Set E1 = New IMSSymbolEntities.DOutput
      
      Dim E2 As IMSSymbolEntities.IJDOutput
      Set E2 = New IMSSymbolEntities.DOutput

      Dim E3 As IMSSymbolEntities.IJDOutput
      Set E3 = New IMSSymbolEntities.DOutput

      Dim E4 As IMSSymbolEntities.IJDOutput
      Set E4 = New IMSSymbolEntities.DOutput

      Dim E5 As IMSSymbolEntities.IJDOutput
      Set E5 = New IMSSymbolEntities.DOutput

      Dim E6 As IMSSymbolEntities.IJDOutput
      Set E6 = New IMSSymbolEntities.DOutput

      Dim E7 As IMSSymbolEntities.IJDOutput
      Set E7 = New IMSSymbolEntities.DOutput

      Dim E8 As IMSSymbolEntities.IJDOutput
      Set E8 = New IMSSymbolEntities.DOutput

      Dim E9 As IMSSymbolEntities.IJDOutput
      Set E9 = New IMSSymbolEntities.DOutput

      Dim E10 As IMSSymbolEntities.IJDOutput
      Set E10 = New IMSSymbolEntities.DOutput

      Dim E11 As IMSSymbolEntities.IJDOutput
      Set E11 = New IMSSymbolEntities.DOutput

      Dim E12 As IMSSymbolEntities.IJDOutput
      Set E12 = New IMSSymbolEntities.DOutput
      
      O1.Name = "Plane1"
      O1.Description = "Plane1 of RectangularSolid"
      O1.Properties = 0
      
      O2.Name = "Plane2"
      O2.Description = "Plane2 of RectangularSolid"
      O2.Properties = 0
      
      O3.Name = "Plane3"
      O3.Description = "Plane3 of RectangularSolid"
      O3.Properties = 0
      
      O4.Name = "Plane4"
      O4.Description = "Plane4 of RectangularSolid"
      O4.Properties = 0
      
      O5.Name = "Plane5"
      O5.Description = "Plane5 of RectangularSolid"
      O5.Properties = 0
      
      O6.Name = "Plane6"
      O6.Description = "Plane6 of RectangularSolid"
      O6.Properties = 0
      
      O7.Name = "Point1"
      O7.Description = "Point1 of RectangularSolid"
      O7.Properties = 0
      
      O8.Name = "Point2"
      O8.Description = "Point2 of RectangularSolid"
      O8.Properties = 0
      
      O9.Name = "Point3"
      O9.Description = "Point3 of RectangularSolid"
      O9.Properties = 0
      
      O10.Name = "Point4"
      O10.Description = "Point4 of RectangularSolid"
      O10.Properties = 0
      
      O11.Name = "Point5"
      O11.Description = "Point5 of RectangularSolid"
      O11.Properties = 0
      
      O12.Name = "Point6"
      O12.Description = "Point6 of RectangularSolid"
      O12.Properties = 0
      
      E1.Name = "Edge1"
      E1.Description = "Edge1 of RectangularSolid"
      E1.Properties = 0

      E2.Name = "Edge2"
      E2.Description = "Edge2 of RectangularSolid"
      E2.Properties = 0

      E3.Name = "Edge3"
      E3.Description = "Edge3 of RectangularSolid"
      E3.Properties = 0

      E4.Name = "Edge4"
      E4.Description = "Edge4 of RectangularSolid"
      E4.Properties = 0

      E5.Name = "Edge5"
      E5.Description = "Edge5 of RectangularSolid"
      E5.Properties = 0

      E6.Name = "Edge6"
      E6.Description = "Edge6 of RectangularSolid"
      E6.Properties = 0

      E7.Name = "Edge7"
      E7.Description = "Edge7 of RectangularSolid"
      E7.Properties = 0

      E8.Name = "Edge8"
      E8.Description = "Edge8 of RectangularSolid"
      E8.Properties = 0

      E9.Name = "Edge9"
      E9.Description = "Edge9 of RectangularSolid"
      E9.Properties = 0

      E10.Name = "Edge10"
      E10.Description = "Edge10 of RectangularSolid"
      E10.Properties = 0

      E11.Name = "Edge11"
      E11.Description = "Edge11 of RectangularSolid"
      E11.Properties = 0

      E12.Name = "Edge12"
      E12.Description = "Edge12 of RectangularSolid"
      E12.Properties = 0
      
      'Define the representation "Physical"
      Dim rep1 As IMSSymbolEntities.IJDRepresentation
      Set rep1 = New IMSSymbolEntities.DRepresentation
    
      rep1.Name = "Physical"
      rep1.Description = "Physical Representation of Rectangular Solid"
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
      oRep1Outputs.SetOutput O7
      oRep1Outputs.SetOutput O8
      oRep1Outputs.SetOutput O9
      oRep1Outputs.SetOutput O10
      oRep1Outputs.SetOutput O11
      oRep1Outputs.SetOutput O12
      oRep1Outputs.SetOutput E1
      oRep1Outputs.SetOutput E2
      oRep1Outputs.SetOutput E3
      oRep1Outputs.SetOutput E4
      oRep1Outputs.SetOutput E5
      oRep1Outputs.SetOutput E6
      oRep1Outputs.SetOutput E7
      oRep1Outputs.SetOutput E8
      oRep1Outputs.SetOutput E9
      oRep1Outputs.SetOutput E10
      oRep1Outputs.SetOutput E11
      oRep1Outputs.SetOutput E12
      
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
      PhysicalRepEval.ProgId = "RectSolid.RectSolidServices"
      
      Set rep1 = Nothing
      Set RepsIf = Nothing
      Set oRep1Outputs = Nothing
    
    'Set the evaluations for the Symbolic and Physical representations on the definition
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
    Err.Raise E_FAIL
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal ActiveConnection As Object) As Object
    On Error GoTo ErrorHandler
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
    Set oSymbolFactory = Nothing
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = "RectSolid.RectSolidServices"
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

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal RepName As String, ByVal OutputColl As Object, arrayOfInputs() As Variant)
    On Error GoTo ErrorHandler
    Set m_outputColl = OutputColl
    If StrComp(RepName, "Physical") = 0 Then
        Physical arrayOfInputs
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF RectangularSolid
'=========================================================================

  Private Sub Physical(ByRef arrayOfInputs())
    On Error GoTo ErrorHandler
    'The planes are constrcuted in such a way that normals point inwards
    Dim RectSolidLength As Double, RectSolidWidth As Double, RectSolidHeight As Double
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oPart As PartFacelets.IJDPart
    
    Set oPart = arrayOfInputs(1)
    RectSolidLength = arrayOfInputs(2)
    RectSolidWidth = arrayOfInputs(3)
    RectSolidHeight = arrayOfInputs(4)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    If RectSolidLength <= 0 Or RectSolidWidth <= 0 Or RectSolidHeight <= 0 Then
        oErrors.Add E_FAIL, "RectSolid.RectSolidServices", "Shape Dimensions should be greater than zero", "ZeroOrNegative"
        GoTo ErrorHandler
    End If
    
    Dim Points(0 To 11) As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim plane As IngrGeom3D.Plane3d
    Dim oPoint3d As IngrGeom3D.Point3d
    Dim line As Line3d
    
    Set oPoint3d = geomFactory.Points3d.CreateByPoint(Nothing, 0, 0, 0)
    m_outputColl.AddOutput "Point1", oPoint3d
    
    Set oPoint3d = geomFactory.Points3d.CreateByPoint(Nothing, RectSolidLength, 0, 0)
    m_outputColl.AddOutput "Point2", oPoint3d

    Set oPoint3d = geomFactory.Points3d.CreateByPoint(Nothing, RectSolidLength / 2, -RectSolidWidth / 2, 0)
    m_outputColl.AddOutput "Point3", oPoint3d

    Set oPoint3d = geomFactory.Points3d.CreateByPoint(Nothing, RectSolidLength / 2, 0, RectSolidHeight / 2)
    m_outputColl.AddOutput "Point4", oPoint3d

    Set oPoint3d = geomFactory.Points3d.CreateByPoint(Nothing, RectSolidLength / 2, RectSolidWidth / 2, 0)
    m_outputColl.AddOutput "Point5", oPoint3d

    Set oPoint3d = geomFactory.Points3d.CreateByPoint(Nothing, RectSolidLength / 2, 0, -RectSolidHeight / 2)
    m_outputColl.AddOutput "Point6", oPoint3d
    Set oPoint3d = Nothing
    
    Points(0) = 0
    Points(1) = -RectSolidWidth / 2
    Points(2) = -RectSolidHeight / 2
    Points(3) = 0
    Points(4) = -RectSolidWidth / 2
    Points(5) = RectSolidHeight / 2
    Points(6) = 0
    Points(7) = RectSolidWidth / 2
    Points(8) = RectSolidHeight / 2
    Points(9) = 0
    Points(10) = RectSolidWidth / 2
    Points(11) = -RectSolidHeight / 2
    Set plane = geomFactory.Planes3d.CreateByPoints(Nothing, 4, Points)
    m_outputColl.AddOutput "Plane1", plane
    Set plane = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, Points(0), Points(1), Points(2), Points(3), Points(4), Points(5))
    m_outputColl.AddOutput "Edge1", line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, Points(3), Points(4), Points(5), Points(6), Points(7), Points(8))
    m_outputColl.AddOutput "Edge2", line
    Set line = Nothing
    
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, Points(6), Points(7), Points(8), Points(9), Points(10), Points(11))
    m_outputColl.AddOutput "Edge3", line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, Points(9), Points(10), Points(11), Points(0), Points(1), Points(2))
    m_outputColl.AddOutput "Edge4", line
    Set line = Nothing
    
    Points(0) = RectSolidLength
    Points(1) = -RectSolidWidth / 2
    Points(2) = -RectSolidHeight / 2
    Points(3) = RectSolidLength
    Points(4) = RectSolidWidth / 2
    Points(5) = -RectSolidHeight / 2
    Points(6) = RectSolidLength
    Points(7) = RectSolidWidth / 2
    Points(8) = RectSolidHeight / 2
    Points(9) = RectSolidLength
    Points(10) = -RectSolidWidth / 2
    Points(11) = RectSolidHeight / 2
    Set plane = geomFactory.Planes3d.CreateByPoints(Nothing, 4, Points)
    m_outputColl.AddOutput "Plane2", plane
    Set plane = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, Points(0), Points(1), Points(2), Points(3), Points(4), Points(5))
    m_outputColl.AddOutput "Edge9", line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, Points(3), Points(4), Points(5), Points(6), Points(7), Points(8))
    m_outputColl.AddOutput "Edge10", line
    Set line = Nothing
    
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, Points(6), Points(7), Points(8), Points(9), Points(10), Points(11))
    m_outputColl.AddOutput "Edge11", line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, Points(9), Points(10), Points(11), Points(0), Points(1), Points(2))
    m_outputColl.AddOutput "Edge12", line
    Set line = Nothing
    
    Points(0) = 0
    Points(1) = RectSolidWidth / 2
    Points(2) = -RectSolidHeight / 2
    Points(3) = 0
    Points(4) = RectSolidWidth / 2
    Points(5) = RectSolidHeight / 2
    Points(6) = RectSolidLength
    Points(7) = RectSolidWidth / 2
    Points(8) = RectSolidHeight / 2
    Points(9) = RectSolidLength
    Points(10) = RectSolidWidth / 2
    Points(11) = -RectSolidHeight / 2
    Set plane = geomFactory.Planes3d.CreateByPoints(Nothing, 4, Points)
    m_outputColl.AddOutput "Plane3", plane
    Set plane = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, Points(0), Points(1), Points(2), Points(9), Points(10), Points(11))
    m_outputColl.AddOutput "Edge5", line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, Points(3), Points(4), Points(5), Points(6), Points(7), Points(8))
    m_outputColl.AddOutput "Edge6", line
    Set line = Nothing
    
    Points(0) = 0
    Points(1) = -RectSolidWidth / 2
    Points(2) = -RectSolidHeight / 2
    Points(3) = RectSolidLength
    Points(4) = -RectSolidWidth / 2
    Points(5) = -RectSolidHeight / 2
    Points(6) = RectSolidLength
    Points(7) = -RectSolidWidth / 2
    Points(8) = RectSolidHeight / 2
    Points(9) = 0
    Points(10) = -RectSolidWidth / 2
    Points(11) = RectSolidHeight / 2
    Set plane = geomFactory.Planes3d.CreateByPoints(Nothing, 4, Points)
    m_outputColl.AddOutput "Plane4", plane
    Set plane = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, Points(0), Points(1), Points(2), Points(3), Points(4), Points(5))
    m_outputColl.AddOutput "Edge7", line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, Points(9), Points(10), Points(11), Points(6), Points(7), Points(8))
    m_outputColl.AddOutput "Edge8", line
    Set line = Nothing
    
    Points(0) = 0
    Points(1) = RectSolidWidth / 2
    Points(2) = -RectSolidHeight / 2
    Points(3) = RectSolidLength
    Points(4) = RectSolidWidth / 2
    Points(5) = -RectSolidHeight / 2
    Points(6) = RectSolidLength
    Points(7) = -RectSolidWidth / 2
    Points(8) = -RectSolidHeight / 2
    Points(9) = 0
    Points(10) = -RectSolidWidth / 2
    Points(11) = -RectSolidHeight / 2
    Set plane = geomFactory.Planes3d.CreateByPoints(Nothing, 4, Points)
    m_outputColl.AddOutput "Plane5", plane
    Set plane = Nothing

    Points(0) = 0
    Points(1) = RectSolidWidth / 2
    Points(2) = RectSolidHeight / 2
    Points(3) = 0
    Points(4) = -RectSolidWidth / 2
    Points(5) = RectSolidHeight / 2
    Points(6) = RectSolidLength
    Points(7) = -RectSolidWidth / 2
    Points(8) = RectSolidHeight / 2
    Points(9) = RectSolidLength
    Points(10) = RectSolidWidth / 2
    Points(11) = RectSolidHeight / 2
    Set plane = geomFactory.Planes3d.CreateByPoints(Nothing, 4, Points)
    m_outputColl.AddOutput "Plane6", plane
    Set plane = Nothing
    
    Set geomFactory = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub


'Following method will be removed very soon. This has been
'included to improve the performance. These lines will be removed as soon as
'symbol subsytem allows this to be removed from Symbol Definition .- Raju 05/05/99.

' CMCache custom method to cache the input argument into a parameter contend and the reverse conversion
'It is up to you to find a way to convert your reference data object to a parameter content
Public Sub CMCacheForPart(pInputCM As Object, bArgToCache As Boolean, pToConvert As Object, ByRef pOutput As Object)

 If bArgToCache Then
       
        'Need to convert the graphic input pToConvert into a Parameter ( pOutput)
        Dim oPC As IJDParameterContent
        Set oPC = New DParameterContent
        
        Dim oPart As IJDPart
        Set oPart = pToConvert
        'MsgBox "Partnumber :" & oPart.PartNumber
        
        ' Create a PC whose value is an identifier of the input
        ' Raju,
        ' the property Part_Number must be retrieved form from the pToConvert argument.
        ' I am hardcoding it
        '
        oPC.Type = igString
        oPC.String = oPart.PartNumber

        Set oPart = Nothing

        ' Always return this PC
        Set pOutput = oPC
        Set oPC = Nothing
 Else
        'Need to convert the cached Parameter pToConvert into your reference data object pOutput
        Dim oPCout As IJDParameterContent
        Dim oPCCache As IJDParameterContent
        Set oPCout = New DParameterContent
        
        ' Here there is three options
        ' o Return a parameter contents, that containts the part_number stored by the pToConvert argument
        '   In this case the edit command will have to retrieve the Part object when needed.
        '   Note : It is better to return a copy of the cached object instead the cached object itself.
        '          This allow to avoid the edition of the cached object.
        ' o Retrieve the Part_number and retrieve your part object with it.
        '   With this solution you can have assoc assertion while when this method is called you are in
        '   in a compute process.
        ' o Get the symbol (or equipment) from the pInputCM, query for the IJDReferencesArg interface,
        '   then get the argument at index 1 it is the SiteProxy.
        '   With this last method, the part has to be passed by reference to the symbol,
        '   but that is what you doing with your design.
        
        ' returning NULL means that the cache method is unable to resolve the cache.
        ' as of now anyway the part is passed as an input argument. this issue of caching
        ' has to be resolved in cycle2.
        
        Set pOutput = Nothing
        
        ' Here is the implementation of the option 1
'        Set oPCCache = pToConvert
'        oPCout.Type = oPCCache.Type
'        oPCout.String = oPCCache.String
'        Set pOutput = oPCout
'        Set oPCout = Nothing
        
        ' Here is the implementation of the option 2
'        Dim oRefDB As GSCADRefDataServices.IJDRefDBGlobal
'        Dim oPart As IJDPart

'        Set oRefDB = New GSCADRefDataServices.RefDBGlobal
'        Set oPart = oRefDB.GetCatalogPart("Storage Tanks", oPCCache.String)
'        Set pOutput = oRefDB
'        Set oRefDB = Nothing
'        Set oPart = Nothing
End If

End Sub


