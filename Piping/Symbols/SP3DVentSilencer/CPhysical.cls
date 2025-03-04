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
'   Author:          svsmylav
'   Creation Date:  Tuesday, Jun 4 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   The symbol supports the following part data basis values:
'   1.)Default
'   2.)Silencer, vent or blowoff(958)
'   3.)Silencer, gas diffuser(959)

'   For the default case, Symbol Model No.: S4AC or S4AZ, Page No. D-100 PDS Piping Component Data Reference Guide.
'   SP3DVentSilencer Symbol is created with Eight Outputs - Five are Insulation aspect outputs.
'   Other three are physical aspect outputs created as follows:
'   ObjBody - Using 'PlaceCylinder' function and
'   two ObjNozzle objects  by using 'CreateNozzleWithLength' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  09.Jul.2003     SymbolTeam(India)    Copyright Information, Header  is added.
'  08.SEP.2006     KKC                  DI-95670  Replace names with initials in all revision history sheets and symbols
'  23.Apr.2008     RRK                  CR-136581 Enhanced the symbol to support new part data basis values-
'                                                  'Silencer, vent or blowoff(958)' and 'Silencer, gas diffuser(959)'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
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
    Dim sptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double

    Dim iOutput As Double
    Dim ObjVentSilencerBody As Object

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset1 As Double
    Dim parOffset2 As Double
    Dim parVentSilencerBodyDiameter As Double
    Dim parVentSilencerBodyLength1 As Double
    Dim parVentSilencerBodyLength2 As Double
    Dim parInsulationThickness As Double
    Dim parHeight As Double
    Dim parLength As Double
    Dim parDiameter As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFace1toCenter = arrayOfInputs(2)
    'parFace2toCenter = arrayOfInputs(3)
    'parOffset1 = arrayOfInputs(4)
    'parOffset2 = arrayOfInputs(5)
    'parVentSilencerBodyDiameter = arrayOfInputs(6)
    'parVentSilencerBodyLength1 = arrayOfInputs(7)
    'parVentSilencerBodyLength2 = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    'parHeight = arrayOfInputs(10)
    'parLength = arrayOfInputs(11)
    'parDiameter = arrayOfInputs(12)

    'Retrieving part data basis value
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Dim dPort1LocationX As Double
    Dim dPort1LocationZ As Double
    Dim NozzleLength As Double

    Select Case lPartDataBasis
    Case Is <= 1    'Default
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
        parOffset1 = arrayOfInputs(4)
        parOffset2 = arrayOfInputs(5)
        parVentSilencerBodyDiameter = arrayOfInputs(6)
        parVentSilencerBodyLength1 = arrayOfInputs(7)
        parVentSilencerBodyLength2 = arrayOfInputs(8)
        
        dPort1LocationX = -parFace1toCenter - sptOffset + depth
        dPort1LocationZ = -parOffset1
        NozzleLength = parFace1toCenter - parVentSilencerBodyLength1
        
    Case 958    'Silencer, vent or blowoff
        parHeight = arrayOfInputs(10)
        parDiameter = arrayOfInputs(12)
        
        dPort1LocationX = -sptOffset + depth
        dPort1LocationZ = 0
        NozzleLength = 0
    Case 959    'Silencer, gas diffuser
        parLength = arrayOfInputs(11)
        parDiameter = arrayOfInputs(12)
        
        dPort1LocationX = -sptOffset + depth
        dPort1LocationZ = 0
        NozzleLength = 0
    Case Else
        GoTo ErrorLabel:    ' Invalid Part data basis.
    End Select
    iOutput = 0
    
    ' Insert your code for output 4(Body of VentSilencer)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Select Case lPartDataBasis
    Case Is <= 1
        stPoint.Set -parVentSilencerBodyLength1, 0, 0
        enPoint.Set parVentSilencerBodyLength2, 0, 0
        Set ObjVentSilencerBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVentSilencerBodyDiameter, True)
    Case 958
        'Setting the points for the line string
        Dim dPoints() As Double    'representing points in the X-Y plane
        ReDim dPoints(0 To 14)

        Dim dCyl2Length As Double
        Dim dEllipsoidLength As Double
        Dim dCyl2Diam As Double
        
        'Setting the diameter of the cylinder 2(of diffuser)
        dCyl2Diam = pipeDiam
        'Check to see that cylinder 2 diameter is greater than silencer diameter
        If CmpDblGreaterthan(dCyl2Diam, parDiameter) Then dCyl2Diam = 0.5 * pipeDiam
        
        'Assumptions:
        '1)The length of the cylinder is assumed to be 10% of Height excluding flange thickness
        '2)The length of the ellipsoid portion is assumed to be 8% of height excluding flange thickness
        dCyl2Length = 0.1 * (parHeight - flangeThick)
        dEllipsoidLength = 0.08 * (parHeight - flangeThick)
        
        'Creating points for the complex string
        'Point 1
        dPoints(0) = flangeThick
        dPoints(1) = dCyl2Diam / 2
        dPoints(2) = 0

        'Point 2
        dPoints(3) = dPoints(0) + dCyl2Length
        dPoints(4) = dPoints(1)
        dPoints(5) = 0

        'Point 3
        dPoints(6) = dPoints(3) + dEllipsoidLength
        dPoints(7) = parDiameter / 2
        dPoints(8) = 0

        'Point 4
        dPoints(9) = parHeight
        dPoints(10) = dPoints(7)
        dPoints(11) = 0
        
        'Point 5
        dPoints(12) = parHeight
        dPoints(13) = 0
        dPoints(14) = 0

        Dim objCollection As Collection
        Set objCollection = New Collection

        Dim oLine As IngrGeom3D.Line3d
        
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        'Creating line 1
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(0), dPoints(1), dPoints(2), _
                                                         dPoints(3), dPoints(4), dPoints(5))
        objCollection.Add oLine


        'Creating ellptical arc
        Dim dMajorRadius As Double
        Dim dMinorRadius As Double

        dMajorRadius = parDiameter / 2 - dCyl2Diam / 2
        dMinorRadius = dEllipsoidLength

        Dim oEllipse As Object
        Set oEllipse = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                       Nothing, _
                       dPoints(6), dCyl2Diam / 2, 0, _
                       0, 0, -1, _
                       0, dMajorRadius, 0, _
                       dMinorRadius / dMajorRadius, _
                       -PI / 2, PI / 2)

        objCollection.Add oEllipse
        Set oEllipse = Nothing

        'Creating line 2
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(6), dPoints(7), dPoints(8), _
                                                         dPoints(9), dPoints(10), dPoints(11))

        objCollection.Add oLine

        stPoint.Set dPoints(0), dPoints(1), dPoints(2)
        
        'Creating line 3
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(9), dPoints(10), dPoints(11), _
                                                         dPoints(12), dPoints(13), dPoints(14))

        objCollection.Add oLine

        stPoint.Set dPoints(0), dPoints(1), dPoints(2)

        Dim objBodyOutline As IngrGeom3D.ComplexString3d
        Set objBodyOutline = PlaceTrCString(stPoint, objCollection)

        Set ObjVentSilencerBody = oGeomFactory.Revolutions3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                                           objBodyOutline, 1, 0, 0, 0, 0, 0, 2 * PI, True)
        Set objCollection = Nothing
        Set oLine = Nothing
        Set objBodyOutline = Nothing

    Case 959
        'Setting the points for the line string
        ReDim dPoints(0 To 8)
        
        'Assumption:The length of the cylinder is assumed to be 85% of length of the silencer excluding flange thickness
        Dim dCylLength As Double
        dCylLength = 0.85 * (parLength - flangeThick)
        
        'Creating points for the complex string
        'Point 1
        dPoints(0) = flangeThick
        dPoints(1) = parDiameter / 2
        dPoints(2) = 0

        'Point 2
        dPoints(3) = dPoints(0) + dCylLength
        dPoints(4) = dPoints(1)
        dPoints(5) = 0

        'Point 3
        dPoints(6) = parLength
        dPoints(7) = 0
        dPoints(8) = 0

        Set objCollection = New Collection

        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        
        'Creating line 1
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(0), dPoints(1), dPoints(2), _
                                                         dPoints(3), dPoints(4), dPoints(5))
        objCollection.Add oLine
        
        'Creating circular arc
        Dim oArc As IngrGeom3D.Arc3d
        Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                               dPoints(3), 0, 0, _
                                                               0, 0, -1, _
                                                               dPoints(3), dPoints(4), dPoints(5), _
                                                               dPoints(6), dPoints(7), dPoints(8))
        objCollection.Add oArc

        stPoint.Set dPoints(0), dPoints(1), dPoints(2)

        Set objBodyOutline = PlaceTrCString(stPoint, objCollection)

        Set ObjVentSilencerBody = oGeomFactory.Revolutions3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                                           objBodyOutline, 1, 0, 0, 0, 0, 0, 2 * PI, True)
        Set objCollection = Nothing

        Set oLine = Nothing
        Set oArc = Nothing
        Set objBodyOutline = Nothing
    End Select

    '   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVentSilencerBody
    Set ObjVentSilencerBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    ' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set dPort1LocationX, 0, dPort1LocationZ
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
    
    '   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    ' Place Nozzle 2( for default case)
    If lPartDataBasis <= 1 Then
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

        oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, parOffset2
        oDir.Set 1, 0, 0
        NozzleLength = parFace2toCenter - parVentSilencerBodyLength2
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

        '   Set the output
        m_OutputColl.AddOutput "PNoz2", objNozzle
        Set objNozzle = Nothing
    End If

    Set oDir = Nothing
    Set oPlacePoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next

End Sub

