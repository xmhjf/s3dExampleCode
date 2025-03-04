VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         kkk
'   Creation Date:  Monday, December 12 2005
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is based on the following Part data basis value that govern its geometry,
'   1)Conduit cap, type E, specified by length, height, width, cover width and cover length
'   Change History:
'   dd.mmm.yyyy     who     change description
'
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   22.MAY.2008     PK   CR-93287  Enhanced the cap symbol implementing partdatabasis 9031
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const COVERHEIGHT = 0.002    'Assumed height of cover
Private Sub Class_Initialize()

'''

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ConduitDiam     As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjNozzle1  As GSCADNozzleEntities.IJConduitPortOcc
    Dim ObjNozzle2  As GSCADNozzleEntities.IJConduitPortOcc
    Dim ObjNozzle3  As GSCADNozzleEntities.IJConduitPortOcc
    Dim ObjCapBody1 As Object
    
    Dim parFacetoEnd As Double
    
    Dim parLength         As Double
    Dim parHeight         As Double
    Dim parWidth          As Double
    Dim parCoverLength    As Double
    Dim parCoverWidth     As Double
    
    Dim lPartdatabasis         As Long
    Dim oPipeComponent         As IJDPipeComponent

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFacetoEnd = arrayOfInputs(2)
    'parLength = arrayOfInputs(3)
    'parHeight = arrayOfInputs(4)
    'parWidth = arrayOfInputs(5)
    'parCoverLength = arrayOfInputs(6)
    'parCoverWidth = arrayOfInputs(7)
    
    Set oPipeComponent = oPartFclt
    lPartdatabasis = oPipeComponent.PartDataBasis
    
    If lPartdatabasis <= 1 Then
        parFacetoEnd = arrayOfInputs(2)
    ElseIf lPartdatabasis = 9031 Then
        parLength = arrayOfInputs(3)
        parHeight = arrayOfInputs(4)
        parWidth = arrayOfInputs(5)
        parCoverLength = arrayOfInputs(6)
        parCoverWidth = arrayOfInputs(7)
    End If
   
    iOutput = 0
    
    If lPartdatabasis <= 1 Then
    
        'Insert your code for output 1(Cylinder Body )
        RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
        Dim stPoint1 As AutoMath.DPosition
        Dim enPoint1 As AutoMath.DPosition
        Dim ObjBody As Object
    
        Set stPoint1 = New AutoMath.DPosition
        Set enPoint1 = New AutoMath.DPosition
        Dim ConduitOD1 As Double
        Dim DomeHeight As Double
        DomeHeight = parFacetoEnd / 4
    
        ConduitOD1 = BodyOD(flangeDiam, ConduitDiam)
    
        stPoint1.Set -parFacetoEnd, 0, 0
        enPoint1.Set parFacetoEnd, 0, 0
    
        Set ObjCapBody1 = PlaceCylinder(m_OutputColl, stPoint1, enPoint1, ConduitOD1, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCapBody1
        Set ObjCapBody1 = Nothing
        Set stPoint1 = Nothing
        Set enPoint1 = Nothing
        
        'Construction of Elliptical Cap -  Output-2
        'Construct Elliptical Head
        'Construct elliptical arc (quarter of ellipce)
        'Center of ellipce
        Dim EllipCenterX As Double, EllipCenterY As Double, EllipCenterZ As Double
        EllipCenterX = parFacetoEnd
        EllipCenterY = 0
        EllipCenterZ = 0
        'Normal vector to plane of ellipce)
        Dim EllipNormX As Double, EllipNormY As Double, EllipNormZ As Double
        EllipNormX = 0
        EllipNormY = 0
        EllipNormZ = -1
        
        'Length of major axis (half of ellipce major axis)
        Dim MajorX As Double, MajorY As Double, MajorZ As Double
        
        MajorX = 0
        MajorY = ConduitOD1 / 2
        MajorZ = 0
        
        'The Ratio of minor axis over major axis
        Dim parRatio As Double
        Dim CenterPos       As New AutoMath.DPosition
        
        parRatio = DomeHeight / (ConduitDiam / 2)
        'Exact value of PI as angle is defined in radian the arc is drawn from 0 to 90 degree
        Dim PI
        PI = 4 * Atn(1)
        CenterPos.Set 0, 0, 0
    
        Dim geomFactory As IngrGeom3D.GeometryFactory
        Set geomFactory = New IngrGeom3D.GeometryFactory
        Dim CapEllip As IngrGeom3D.EllipticalArc3d
        Set CapEllip = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                                Nothing, EllipCenterX, _
                                                EllipCenterY, EllipCenterZ, EllipNormX, _
                                                EllipNormY, EllipNormZ, MajorX, MajorY, MajorZ, _
                                                parRatio, 0, PI / 2)
        'Rotate elliptical arc
        Dim axis        As New AutoMath.DVector
        Dim oPlacePoint As New AutoMath.DPosition
        Set oPlacePoint = New AutoMath.DPosition
        axis.Set 1, 0, 0
        Set ObjBody = PlaceRevolution(m_OutputColl, CapEllip, axis, CenterPos, 2 * PI, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
        Set ObjBody = Nothing
         
    ElseIf lPartdatabasis = 9031 Then
    
        'Insert your code for Body
        RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
      
        Dim oGeomFactory    As IngrGeom3D.GeometryFactory
        Dim oLine           As Object
        Dim oArc            As IngrGeom3D.Arc3d
        Dim lines           As Collection
        Dim oCString        As IngrGeom3D.ComplexString3d
        Dim oStPoint        As AutoMath.DPosition
        Dim iCount          As Integer
        Dim oBody           As Object
        Dim oAxis           As AutoMath.DVector
        
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oCString = New IngrGeom3D.ComplexString3d
        Set oArc = New IngrGeom3D.Arc3d
        Set lines = New Collection
        Set oStPoint = New AutoMath.DPosition
        Set oAxis = New AutoMath.DVector
        
        'Insert your code for Complex String
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
            0.2 * parLength + parWidth / 2, parWidth / 2, 0.5 * parHeight, _
            parLength - parWidth / 2, parWidth / 2, 0.5 * parHeight)
        lines.Add oLine
        
        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
            parLength - parWidth / 2, parWidth / 2, 0.5 * parHeight, _
            parLength, 0, 0.5 * parHeight, _
            parLength - parWidth / 2, -parWidth / 2, 0.5 * parHeight)
        lines.Add oArc
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
            parLength - parWidth / 2, -parWidth / 2, 0.5 * parHeight, _
            0.2 * parLength + parWidth / 2, -parWidth / 2, 0.5 * parHeight)
        lines.Add oLine
        
        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
            0.2 * parLength + parWidth / 2, -parWidth / 2, 0.5 * parHeight, _
            0.2 * parLength, 0, 0.5 * parHeight, _
            0.2 * parLength + parWidth / 2, parWidth / 2, 0.5 * parHeight)
        lines.Add oArc
        
        oStPoint.Set 0.2 * parLength + parWidth / 2, parWidth / 2, 0.5 * parHeight
        oAxis.Set 0, 0, -1
        Set oCString = PlaceTrCString(oStPoint, lines)
        Set oBody = PlaceProjection(m_OutputColl, oCString, oAxis, parHeight, True)
        
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
        
        For iCount = 1 To lines.Count
            lines.Remove 1
        Next iCount
        Set lines = Nothing
        Set oCString = Nothing
        Set oStPoint = Nothing
        Set oLine = Nothing
        Set oArc = Nothing
        Set oBody = Nothing
        Set oAxis = Nothing
        
        'Insert your code for Cover
        Dim oCoverString   As IngrGeom3D.ComplexString3d
        Dim oCover         As Object
        
        Set oAxis = New AutoMath.DVector
        Set lines = New Collection
        Set oArc = New IngrGeom3D.Arc3d
        Set oStPoint = New AutoMath.DPosition
        Set oCoverString = New IngrGeom3D.ComplexString3d
                      
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
            0.6 * parLength - parCoverLength / 2 + parCoverWidth / 2, _
            parCoverWidth / 2, 0.5 * parHeight, _
            0.6 * parLength + parCoverLength / 2 - parCoverWidth / 2, _
            parCoverWidth / 2, 0.5 * parHeight)
        lines.Add oLine
        
        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
            0.6 * parLength + parCoverLength / 2 - parCoverWidth / 2, _
            parCoverWidth / 2, 0.5 * parHeight, _
            0.6 * parLength + parCoverLength / 2, 0, 0.5 * parHeight, _
            0.6 * parLength + parCoverLength / 2 - parCoverWidth / 2, _
            -parCoverWidth / 2, 0.5 * parHeight)
        lines.Add oArc
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
            0.6 * parLength + parCoverLength / 2 - parCoverWidth / 2, _
            -parCoverWidth / 2, 0.5 * parHeight, _
            0.6 * parLength - parCoverLength / 2 + parCoverWidth / 2, _
            -parCoverWidth / 2, 0.5 * parHeight)
        lines.Add oLine
        
        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
            0.6 * parLength - parCoverLength / 2 + parCoverWidth / 2, _
            -parCoverWidth / 2, 0.5 * parHeight, _
            0.6 * parLength - parCoverLength / 2, 0, 0.5 * parHeight, _
            0.6 * parLength - parCoverLength / 2 + parCoverWidth / 2, _
            parCoverWidth / 2, 0.5 * parHeight)
        lines.Add oArc
        
        oStPoint.Set 0.6 * parLength - parCoverLength / 2 + parCoverWidth / 2, _
            parCoverWidth / 2, 0.5 * parHeight
            
        oAxis.Set 0, 0, 1
        Set oCoverString = PlaceTrCString(oStPoint, lines)
        Set oCover = PlaceProjection(m_OutputColl, oCoverString, oAxis, COVERHEIGHT, True)

        m_OutputColl.AddOutput "Cover", oCover
        
        Set oGeomFactory = Nothing
        Set oArc = Nothing
        Set oLine = Nothing
        For iCount = 1 To lines.Count
            lines.Remove 1
        Next iCount
        Set lines = Nothing
        Set oStPoint = Nothing
        Set oCoverString = Nothing
       
        'Insert your code for the Conduit body connection
        Dim oConnectObj As Object
        Dim oEnPoint    As AutoMath.DPosition
        
        Set oEnPoint = New AutoMath.DPosition
        Set oStPoint = New AutoMath.DPosition
           
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0.3 * parLength, 0, 0
        
        Set oConnectObj = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * ConduitDiam, True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oConnectObj
        Set oConnectObj = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        
    End If
        
    'OutPut - 3 Place Nozzle 1
    'Insert your code for output 3(Nozzle 1)
        
    Dim oDir        As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitDiam, flangeThick, flangeDiam, sptOffset, depth
    
    If lPartdatabasis <= 1 Then
        oPlacePoint.Set -sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
    ElseIf lPartdatabasis = 9031 Then
        oPlacePoint.Set -sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
    End If
    
    Set ObjNozzle1 = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 1)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1
    
    Set ObjNozzle1 = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
       
End Sub
