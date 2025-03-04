VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CircularTorusServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    '******************************************************************
' Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'File
'    CircularTorusServices.cls
'
'Author
'    Mule
'
'Description
'    Definition of CircularTorusServices Symbol
'History
'10 Apr 2002    MU           Creation
'30 May 2003    VS      Change the normal direction of the boundary imposed on plane so that
'                           the area inside the boundary is treated as material
'02 Jun 2003    VS     Change the creation of end planes to use the CreateByOuterBdry method,
'                           this allows to check if the boundary has been created properly
'09 Jul 2003    SymbolTeam  Copyright Information, Header  is added/Updated.
'               (India)
'24 Jul 2003    VS      Do not clear middle tier errors. This may clear out other errors too.
'16 Oct 2003    VS     Added edges at the faces for CR36485. SmartSketch will locate points on these edges
'09 Nov 2004    VS      TR64758 - Change the symbol so that the X axis is perpendicular to the face.
'
'Notes
'
'    <notes>
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  19.DEC.2007     PK            Undone the changes made in TR-132021
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Private Const E_FAIL = -2147467259
Private Const PI = 3.14159265
Private Const TOLERANCE = 0.000001
Private m_bAngle As Boolean 'Flag required to take care of 360 degrees
Implements IJDUserSymbolServices

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = "CircularTorus.CircularTorusServices"
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
   
    'Feed CircularTorusDefnition
    ' Inputs:
    '          1. "Part"  ( Catalog part )
    '          2. "TorusDiameter"
    '          3. "TorusDistance"
    '          4. "TorusSweepAngle"
    ' Representations :
    '           Physical
    
    On Error GoTo ErrorHandler
    
     ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
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
    
    'Define inputs for CircularTorus
    Dim Index As Integer
    For Index = 1 To 3
        Set Inputs(Index) = New IMSSymbolEntities.DInput
        Inputs(Index).Properties = igINPUT_IS_A_PARAMETER
        If (Index = 1) Then
            PC.UomValue = 0.2
            Inputs(Index).Name = "TorusDiameter"
            Inputs(Index).Description = "Diameter of the Torus"
        ElseIf (Index = 2) Then
            PC.UomValue = 1#
            Inputs(Index).Name = "TorusDistance"
        Else
            PC.UomValue = PI / 4
            Inputs(Index).Name = "TorusSweepAngle"
        End If
        Inputs(Index).DefaultParameterValue = PC
    Next
    
    For Index = 1 To 3
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next
    
    'Define the outputs
    Dim O1 As IMSSymbolEntities.DOutput
    Set O1 = New IMSSymbolEntities.DOutput
    
    O1.Name = "Torii"
    O1.Description = "CircularTorii"
    O1.Properties = 0
    
    Dim oLocatePlane As IMSSymbolEntities.DOutput
    Set oLocatePlane = New IMSSymbolEntities.DOutput
    
    oLocatePlane.Name = "Surface1"
    oLocatePlane.Description = "Default Surface used for creating relations"
    oLocatePlane.Properties = 0
    
    Dim oLocatePlane1 As IMSSymbolEntities.DOutput
    Set oLocatePlane1 = New IMSSymbolEntities.DOutput
    
    oLocatePlane1.Name = "Surface2"
    oLocatePlane1.Description = "Default Surface used for creating relations"
    oLocatePlane1.Properties = 0
    
     'Define the outputs
    Dim oPoint1 As IMSSymbolEntities.DOutput
    Set oPoint1 = New IMSSymbolEntities.DOutput
    
    oPoint1.Name = "Point1"
    oPoint1.Description = "CircularTorus"
    oPoint1.Properties = 0
    
    Dim oPoint2 As IMSSymbolEntities.DOutput
    Set oPoint2 = New IMSSymbolEntities.DOutput
    
    oPoint2.Name = "Point2"
    oPoint2.Description = "CircularTorus"
    oPoint2.Properties = 0
    
    Dim oPoint3 As IMSSymbolEntities.DOutput
    Set oPoint3 = New IMSSymbolEntities.DOutput
    
    oPoint3.Name = "Point3"
    oPoint3.Description = "CircularTorus"
    oPoint3.Properties = 0
    
    Dim E1 As IMSSymbolEntities.DOutput
    Set E1 = New IMSSymbolEntities.DOutput
    E1.Name = "Edge1"
    E1.Description = "Edge1 of RightCircularCylinder"
    E1.Properties = 0
    
    Dim E2 As IMSSymbolEntities.DOutput
    Set E2 = New IMSSymbolEntities.DOutput
    E2.Name = "Edge2"
    E2.Description = "Edge2 of RightCircularCylinder"
    E2.Properties = 0
    'Define the representation "Symbolic"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation

    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of CircularTorus"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim IJDOutputs As IMSSymbolEntities.IJDOutputs
    Set IJDOutputs = rep1

    ' Set the output
    IJDOutputs.SetOutput O1
    Set O1 = Nothing
    
    IJDOutputs.SetOutput oLocatePlane
    Set oLocatePlane = Nothing
    
    IJDOutputs.SetOutput oLocatePlane1
    Set oLocatePlane1 = Nothing
    
    IJDOutputs.SetOutput oPoint1
    Set oPoint1 = Nothing
    
    IJDOutputs.SetOutput oPoint2
    Set oPoint2 = Nothing
    
    IJDOutputs.SetOutput oPoint3
    Set oPoint3 = Nothing
    
    IJDOutputs.SetOutput E1
    Set E1 = Nothing
    
    IJDOutputs.SetOutput E2
    Set E2 = Nothing
    
    'Set the representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1

    Set rep1 = Nothing
    Set RepsIf = Nothing
    Set IJDOutputs = Nothing
    
    'Define evaluation for Physical representation
    Dim PhysicalRepEval As IJDRepresentationEvaluation
    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "Physical representation"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "CircularTorus.CircularTorusServices"
    
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
    oSymbolDefinition.ProgId = "CircularTorus.CircularTorusServices"
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
'CREATION OF PHYSICAL REPRESENTATION OF CircularTorus
'=========================================================================
Private Function Physical(ByRef arrayOfInputs())
    On Error GoTo ErrorHandler
    
    Dim dDiameter As Double, dDistance As Double
    Dim dSweepAngle As Double
    Dim oPart As IJDPart
    Set oPart = arrayOfInputs(1)
    
    dDiameter = arrayOfInputs(2)
    dDistance = arrayOfInputs(3)
    dSweepAngle = arrayOfInputs(4)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    If dDiameter <= 0 Or dDistance <= 0 Or dSweepAngle <= 0 Then
        oErrors.Add E_FAIL, "CircularTorus.CircularTorusServices", "Shape Dimensions should be greater than zero", "ZeroOrNegative"
        GoTo ErrorHandler
    ElseIf dDiameter > 2 * dDistance Then
        oErrors.Add E_FAIL, "CircularTorus.CircularTorusServices", "A should not be grater then twice B", "A>2*B"
        GoTo ErrorHandler
    End If
    
     'Case of 360 degrees.
    If (Abs(dSweepAngle - 6.283185) < TOLERANCE) Then
        m_bAngle = True
    Else
        While (dSweepAngle > 6.283185)
            dSweepAngle = dSweepAngle - 6.283185
            If (dSweepAngle < TOLERANCE) Then
                 m_bAngle = True
                 dSweepAngle = 6.283185
            End If
        Wend
    End If
    
    Dim dAxisCenterX As Double, dAxisCenterY As Double, dAxisCenterZ As Double
    dAxisCenterX = 0#
    dAxisCenterY = -dDistance
    dAxisCenterZ = 0#

    Dim dAxisVecX As Double, dAxisVecY As Double, dAxisVecZ As Double
    dAxisVecX = 0#
    dAxisVecY = 0#
    dAxisVecZ = -1#

    Dim dOriginDirX As Double, dOriginDirY As Double, dOriginDirZ As Double
    dOriginDirX = 0#
    dOriginDirY = 1#
    dOriginDirZ = 0#

    '=====================================
    ' CONSTRUCTION OF  CircularTorus
    '=====================================
    Dim oGeomFactory As New GeometryFactory
    Dim oTorus As Torus3d
    
    Set oTorus = oGeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(Nothing, _
                                dAxisCenterX, dAxisCenterY, dAxisCenterZ, _
                                dAxisVecX, dAxisVecY, dAxisVecZ, _
                                dOriginDirX, dOriginDirY, dOriginDirZ, _
                                  dDistance, dDiameter / 2, dSweepAngle, True)
    m_outputColl.AddOutput "Torii", oTorus
    Set oTorus = Nothing
    
    If (m_bAngle = False) Then
        Dim oPlane3d As Plane3d
        Dim oCircle3d As Circle3d
        Dim oPoint3d As Point3d
        
        Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, 0, 0, 0)
        m_outputColl.AddOutput "Point1", oPoint3d
        Set oPoint3d = Nothing
        
        Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, dDistance * Sin(dSweepAngle), -2 * dDistance * Sin(dSweepAngle / 2) * Sin(dSweepAngle / 2), 0)
        m_outputColl.AddOutput "Point2", oPoint3d
        Set oPoint3d = Nothing
        
        Set oCircle3d = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                     0, 0, 0, -1, 0, 0, dDiameter / 2)
        Set oPlane3d = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oCircle3d)
        m_outputColl.AddOutput "Surface1", oPlane3d
        m_outputColl.AddOutput "Edge1", oCircle3d
        Set oPlane3d = Nothing
        Set oCircle3d = Nothing
        
        Set oCircle3d = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                     dDistance * Sin(dSweepAngle), -2 * dDistance * Sin(dSweepAngle / 2) * Sin(dSweepAngle / 2), 0, Cos(dSweepAngle), -Sin(dSweepAngle), 0, dDiameter / 2)
        Set oPlane3d = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oCircle3d)
        m_outputColl.AddOutput "Surface2", oPlane3d
        m_outputColl.AddOutput "Edge2", oCircle3d
        Set oPlane3d = Nothing
        Set oCircle3d = Nothing
    End If
                
    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, dAxisCenterX, dAxisCenterY, dAxisCenterZ)
    
    m_outputColl.AddOutput "Point3", oPoint3d
    Set oPoint3d = Nothing
    
    Set oGeomFactory = Nothing
    Exit Function
ErrorHandler:
    Err.Raise E_FAIL
End Function

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


