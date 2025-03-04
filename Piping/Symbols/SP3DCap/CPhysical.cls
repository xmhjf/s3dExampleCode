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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Wednesday, Nov 22 2000
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   09.Jul.2003     SymbolTeam(India)   Copyright Information, Header  is added.
'   01.Aug.2006     KKK                 Replaced 'm_OutputColl.ResourceManager' with "Nothing'.
'  08.SEP.2006      KKC                 DI-95670  Replace names with initials in all revision history sheets and symbols.
'  30.Jan.2007      dkl                 CR 112819 Implemented Part data basis to modify symbol for HDPE spec.
'  19.Feb.2008      RUK                 CR-33401  Create symbols, as required, for jacketed piping specs
'                                       Implemented the PDB:"Cap, jacket, specified by cap length, tangent length, and inside diameter (819)"
'   25.Sep.2008     MP     CR-150009    Implemented part data basis values 1087, 1088 and 1089.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
    Private Const MODULE = "Physical:"    'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim cptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double

    Dim iOutput As Integer
    Dim parFacetoEnd As Double
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    
   
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoEnd = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    'parFace1toCenter = arrayOfInputs(4)
    'parCapInsideDia = arrayOfInputs(4)
    'parCapLength = arrayOfInputs(5)
    'parTangentLength = arrayOfInputs(6)
    
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis

    If lPartDataBasis = 1087 Or lPartDataBasis = 1088 Or lPartDataBasis = 1089 Then
        parFacetoFace = arrayOfInputs(2)
    Else
        parFacetoEnd = arrayOfInputs(2)
    End If
    
    iOutput = 0


    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    Dim oPlacePoint As New AutoMath.DPosition
    Set oPlacePoint = New AutoMath.DPosition
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    Dim oCenPoint As AutoMath.DPosition
    
    Dim dNozzleLength As Double
    Dim objBody As Object

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                       cptOffset, depth

    ' The Cap symbol is based on the following three Part data basis values that govern its geometry,
    ' End Cap, welded' - 57, 'End Cap, machined' - 60, and 'End Cap, molded' - 65 .

    ' Code for both Butt Welded and Socket Welded or Female Threaded Cap
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
            
    Select Case lPartDataBasis

    Case Is <= 1, 57    'End Cap is Welded or the default state.
        'The value of Port Depth defines if the Cap is Male end:Bevled End or
        'Female end: Socket Welded or Female Threaded.
        If CmpDblEqual(depth, 0) Then   'Cap is Beveled End
            Dim dDomeHeight As Double
            dNozzleLength = parFacetoEnd * 3 / 4
            dDomeHeight = parFacetoEnd / 4

            'Construct Elliptical Head

            'Center of ellipse
            Dim dEllipseCenterX As Double, dEllipseCenterY As Double, dEllipseCenterZ As Double
            dEllipseCenterX = dNozzleLength
            dEllipseCenterY = 0
            dEllipseCenterZ = 0

            'Vector normal to plane of the ellipse
            Dim dEllipseNormX As Double, dEllipseNormY As Double, dEllipseNormZ As Double
            dEllipseNormX = 0
            dEllipseNormY = 0
            dEllipseNormZ = -1

            'Length of major axis (half of ellipse major axis)
            Dim dMajorX As Double, dMajorY As Double, dMajorZ As Double
            dMajorX = 0
            dMajorY = pipeDiam / 2
            dMajorZ = 0

            'The Ratio of minor axis to major axis
            Dim parRatio As Double
            parRatio = dDomeHeight / (pipeDiam / 2)

            'The value of PI as angle is defined in radians.
            'Construct elliptical arc (quarter of ellipse)

            Dim CapEllip As IngrGeom3D.EllipticalArc3d

            Set CapEllip = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                           Nothing, dEllipseCenterX, dEllipseCenterY, dEllipseCenterZ, _
                           dEllipseNormX, dEllipseNormY, dEllipseNormZ, _
                           dMajorX, dMajorY, dMajorZ, parRatio, 0, PI / 2)

            'Rotate the elliptical arc
            Dim oAxisVect As AutoMath.DVector
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 1, 0, 0

            Set objBody = PlaceRevolution(m_OutputColl, CapEllip, oAxisVect, _
                                          oPlacePoint, 2 * PI, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
            Set CapEllip = Nothing
            Set oAxisVect = Nothing

            oDir.Set -1, 0, 0
            oPlacePoint.Set 0, 0, 0

            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                                   oPlacePoint, dNozzleLength)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle

        Else   'Cap is Socket Welded or Female Threaded
            Set oStPoint = New AutoMath.DPosition
            Set oEndPoint = New AutoMath.DPosition
            oStPoint.Set -(depth - cptOffset), 0, 0
            oEndPoint.Set parFacetoEnd - (depth - cptOffset), 0, 0

            Set objBody = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, flangeDiam, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
            Set oStPoint = Nothing
            Set oEndPoint = Nothing

            oDir.Set -1, 0, 0
            oPlacePoint.Set 0, 0, 0

            'The Nozzle is just a circle at the port CenterPos ==> Boolean is TRUE
            Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_OutputColl, _
                                                    oDir, oPlacePoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        End If

    Case 60    'End Cap is IPS machined.
        Set oStPoint = New AutoMath.DPosition
        Set oEndPoint = New AutoMath.DPosition

        oStPoint.Set -(depth - cptOffset), 0, 0
        oEndPoint.Set parFacetoEnd - (depth - cptOffset), 0, 0

        Set objBody = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, pipeDiam, True)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set oStPoint = Nothing
        Set oEndPoint = Nothing

        'Create Nozzle
        oDir.Set -1, 0, 0
        oPlacePoint.Set 0, 0, 0

        'The Nozzle is just a circle at the port CenterPos ==> Boolean is TRUE
        Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle

    Case 65    'End Cap is IPS molded.
        Dim parFace1toCenter As Double
        Dim LineStrPoints(0 To 20) As Double
        Dim oLineString As IngrGeom3D.LineString3d
        Set oAxisVect = New AutoMath.DVector
        Set oCenPoint = New AutoMath.DPosition
        parFace1toCenter = arrayOfInputs(4)

        ' Assumptions:
        ' 1) The smaller diameter of the cylinder at the other end (the end other than
        '    the port) equals 0.75 x Pipe diameter.
        ' 2) The length of the tapered transition equals 0.25 x Face1 to Center.

        ' Creating the tapered end portion of the Cap.

        'Point 1
        LineStrPoints(0) = -(depth - cptOffset)
        LineStrPoints(1) = pipeDiam / 2
        LineStrPoints(2) = 0

        'Point 2
        LineStrPoints(3) = LineStrPoints(0) + parFace1toCenter
        LineStrPoints(4) = pipeDiam / 2
        LineStrPoints(5) = LineStrPoints(2)

        'Point 3, forming the tapered transition. it is calculated Based on the
        'assumtion 2)mentioned above.
        LineStrPoints(6) = LineStrPoints(3) + 0.25 * parFace1toCenter
        LineStrPoints(7) = (0.75 * pipeDiam) / 2    'Based on the assumtion 1) mentioned above.
        LineStrPoints(8) = LineStrPoints(2)

        'Point 4, reaching the end of the Cap.
        LineStrPoints(9) = parFacetoEnd - (depth - cptOffset)
        LineStrPoints(10) = LineStrPoints(7)
        LineStrPoints(11) = LineStrPoints(2)

        'Point 5
        LineStrPoints(12) = LineStrPoints(9)
        LineStrPoints(13) = 0
        LineStrPoints(14) = LineStrPoints(2)

        'Point 6
        LineStrPoints(15) = LineStrPoints(0)
        LineStrPoints(16) = 0
        LineStrPoints(17) = LineStrPoints(2)

        'Point 7 coinciding with Point 1
        LineStrPoints(18) = LineStrPoints(0)
        LineStrPoints(19) = LineStrPoints(1)
        LineStrPoints(20) = LineStrPoints(2)

        Set geomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                   7, LineStrPoints)
        oAxisVect.Set 1, 0, 0
        oCenPoint.Set 0, -0.0001, 0
        PI = 4 * Atn(1)

        Set objBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenPoint, _
                                    2 * PI, True)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        Set oAxisVect = Nothing
        Set oCenPoint = Nothing
        Set oLineString = Nothing

        'Create Nozzle with length
        oDir.Set -1, 0, 0
        dNozzleLength = parFace1toCenter
        oPlacePoint.Set 0, 0, 0

        Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_OutputColl, _
                                                oDir, oPlacePoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    
    Case 819    'Cap, jacket, specified by cap length, tangent length, and inside diameter
    'CR-33401  Create symbols, as required, for jacketed piping specs
        Dim parCapInsideDia As Double
        Dim parCapLength As Double
        Dim parTangentLength As Double
        
        'Inputs
        parCapInsideDia = arrayOfInputs(5)
        parCapLength = arrayOfInputs(6)
        parTangentLength = arrayOfInputs(7)
        
        Dim oCompStr As IngrGeom3D.ComplexString3d
        Dim oCollection As Collection
        Dim oLine As IngrGeom3D.Line3d
        Dim oEllArc As IngrGeom3D.EllipticalArc3d
        Dim oMajor As AutoMath.DPosition
        Dim oNormal As AutoMath.DVector
        Dim dMMRatio As Double
        Dim iCount As Integer
        
        Set oCompStr = New ComplexString3d
        Set oCollection = New Collection
        Set oLine = New Line3d
        Set oEllArc = New EllipticalArc3d
        Set oStPoint = New DPosition
        Set oEndPoint = New DPosition
        Set oCenPoint = New DPosition
        Set oMajor = New DPosition
        Set oNormal = New DVector
        
        'Assume Tangent Length = 3/4 of Cap Length, if its value = 0
        Dim dTangentLength As Double
        dTangentLength = parTangentLength
        If CmpDblEqual(parTangentLength, 0) Then
            dTangentLength = (3 / 4) * parCapLength
        End If
        
        'Line 1
        oStPoint.Set -(depth - cptOffset), pipeDiam / 2, 0
        oEndPoint.Set -((depth - cptOffset) - dTangentLength), oStPoint.y, oStPoint.z
        Set oLine = PlaceTrLine(oStPoint, oEndPoint)
        oCollection.Add oLine
        Set oLine = Nothing
        
        'Elliptical Arc
        oCenPoint.Set -(depth - cptOffset) + dTangentLength, parCapInsideDia / 2, 0
        oNormal.Set 0, 0, -1
        oMajor.Set 0, (pipeDiam - parCapInsideDia) / 2, 0
        dMMRatio = (parCapLength - dTangentLength) / ((pipeDiam - parCapInsideDia) / 2)
        Set oEllArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                        oCenPoint.x, oCenPoint.y, oCenPoint.z, _
                                        oNormal.x, oNormal.y, oNormal.z, _
                                        oMajor.x, oMajor.y, oMajor.z, dMMRatio, 0, PI / 2)
        oCollection.Add oEllArc
        Set oEllArc = Nothing
        Set oMajor = Nothing
        
        'Create the Complex String
        oStPoint.Set -(depth - cptOffset), pipeDiam / 2, 0
        Set oCompStr = PlaceTrCString(oStPoint, oCollection)
        
        'Revolve the complex string about the center
        oNormal.Set -1, 0, 0
        oCenPoint.Set 0, 0, 0
        Set objBody = PlaceRevolution(m_OutputColl, oCompStr, oNormal, oCenPoint, 2 * PI, False)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
        
        'Remove the References
        Set oCompStr = Nothing
        For iCount = 1 To oCollection.Count
            oCollection.Remove 1
        Next iCount
        Set oCollection = Nothing
        Set oNormal = Nothing
        
        'Create Nozzle
        oDir.Set -1, 0, 0
        oPlacePoint.Set 0, 0, 0

        Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        
    Case 1087, 1088, 1089
    
        Dim lNumSides As Long
        Dim oTransMat As AutoMath.DT4x4
        Dim oVec As AutoMath.DVector
        Set oVec = New AutoMath.DVector
        Set oTransMat = New AutoMath.DT4x4
    
        If lPartDataBasis = 1087 Then
            lNumSides = 6
        ElseIf lPartDataBasis = 1088 Then
            lNumSides = 8
        ElseIf lPartDataBasis = 1089 Then
            lNumSides = 10
        End If

        Dim dInculdedAngle As Double
        Dim dSideLength As Double
        Set oEndPoint = New AutoMath.DPosition
        
        dInculdedAngle = 360 / lNumSides
        dInculdedAngle = (dInculdedAngle * PI) / 180
        oEndPoint.Set parFacetoFace, 0, 0
        
        dSideLength = flangeDiam * Tan(dInculdedAngle / 2)
        
        Set objBody = PlaceNnagon(m_OutputColl, lNumSides, dSideLength, parFacetoFace, True)
        
        oVec.Set 0, 0, 0
        oTransMat.LoadIdentity
        oTransMat.Translate oVec
        objBody.Transform oTransMat
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
                
        'Create Nozzle
        oDir.Set -1, 0, 0
        oPlacePoint.Set depth - cptOffset, 0, 0
        
        Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        
        
    Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.
    End Select
    Set geomFactory = Nothing
    Set objBody = Nothing
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
    Set oCenPoint = Nothing
    Set oPipeComponent = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
