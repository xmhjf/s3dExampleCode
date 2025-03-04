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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Friday, Nov 24 2000
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy        who                  change description
'   -----------        -----                ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   20.Nov.2003     SymbolTeam(India)       TR-52620--Introduced check for output 'Body' so that it does not fail for Butt-welded End Preparations.
'   01.Sep.2006      KKC                    CR-100189  Added Part Data Basis property to specify coupling either in face-to-face
'                                                       or face-to-center or face1-to-center and face2 to center.
'   14-SEP-06        KKC                    DI-95670  Replace names with initials in all revision history sheets and symbols")
'   21-Dec-06       svsmylav                TR-112718: FacetoCenter and Face1toCenter are optional inputs added in V7 -
'                                           these are moved to the end of first two V6.1 existing inputs (otherwise mirror operation failed).
'   06-Nov-07        VRG                    CR-127647 Enhanced the symbol to support Flexible Coupling using PartDataBasis value.
'                                           (Source: http://www.mdcvacuum.com/urd/uniface.urd/ecf0070w.display?1.6.4.1)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:"    'Used for error messages
Private PI As Double

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
    Dim ObjBody As Object

    Dim parInsulationThickness As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim lPartDataBasis As Long
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInsulationThickness = arrayOfInputs(3)

    iOutput = 0

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    If (lPartDataBasis = 5) Or (lPartDataBasis <= 1) Then   'Applicable for 5 "Face-to-face dimension basis"
        Dim parFacetoFace As Double
        parFacetoFace = arrayOfInputs(2)
        parFacetoCenter = parFacetoFace / 2     'parFace1toCenter and parFace2toCenter: These dimensions are calculated and used in computation.
        parFace1toCenter = parFacetoFace / 2
    ElseIf (lPartDataBasis = 420) Then          'Partdatabasis value 420 is used by flexible coupling for Kwik flange/Large flange
        parFacetoFace = arrayOfInputs(2)
        parFacetoCenter = parFacetoFace / 2
        parFace1toCenter = parFacetoFace / 2
    ElseIf (lPartDataBasis = 10) Then           'Applicable for  1 "Default" & 10 -"Face-to-center dimension basis"
        parFacetoCenter = arrayOfInputs(4)
        parFace1toCenter = arrayOfInputs(5)
        If CmpDblEqual(parFace1toCenter, 0) Then
            parFace1toCenter = parFacetoCenter
        End If
    End If
    'Insert your code for output 1(Body)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set -parFacetoCenter, 0, 0
    enPoint.Set parFace1toCenter, 0, 0

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim parFlangeDiameter As Double

    parFlangeDiameter = flangeDiam
    If CmpDblLessThanOrEqualTo(flangeDiam, 0) Then parFlangeDiameter = pipeDiam

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim parFlangeDiameter2 As Double

    parFlangeDiameter2 = flangeDiam
    If CmpDblLessThanOrEqualTo(flangeDiam, 0) Then parFlangeDiameter2 = pipeDiam

    Dim parBodyOD As Double
    parBodyOD = BodyOD(parFlangeDiameter, parFlangeDiameter2)
'   We provided implementation for flexible coupling (lPartDataBasis = 420) case below, and for other
'   part data basis cases 'Else' block is used
    If (lPartDataBasis = 420) Then

        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        'Insert your code for output 1(Body)
        Dim oAxisVec As AutoMath.DVector
        Dim oCenter As AutoMath.DPosition
        Set oAxisVec = New AutoMath.DVector
        Set oCenter = New AutoMath.DPosition
        Dim dLineStrPoints(0 To 23) As Double
        Dim oLineString As IngrGeom3D.LineString3d

        dLineStrPoints(0) = -parFacetoFace / 2
        dLineStrPoints(1) = 0
        dLineStrPoints(2) = 0
        dLineStrPoints(3) = -parFacetoFace / 2
        dLineStrPoints(4) = pipeDiam / 2
        dLineStrPoints(5) = 0
        dLineStrPoints(6) = -0.3 * parFacetoFace
        dLineStrPoints(7) = pipeDiam / 2
        dLineStrPoints(8) = 0
        dLineStrPoints(9) = -0.3 * parFacetoFace
        dLineStrPoints(10) = 0.6 * pipeDiam
        dLineStrPoints(11) = 0
        dLineStrPoints(12) = 0.3 * parFacetoFace
        dLineStrPoints(13) = 0.6 * pipeDiam
        dLineStrPoints(14) = 0
        dLineStrPoints(15) = 0.3 * parFacetoFace
        dLineStrPoints(16) = pipeDiam / 2
        dLineStrPoints(17) = 0
        dLineStrPoints(18) = parFacetoFace / 2
        dLineStrPoints(19) = pipeDiam / 2
        dLineStrPoints(20) = 0
        dLineStrPoints(21) = parFacetoFace / 2
        dLineStrPoints(22) = 0
        dLineStrPoints(23) = 0

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 8, dLineStrPoints)
        oAxisVec.Set 1, 0, 0
        oCenter.Set 0, 0, 0
        Set ObjBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVec, oCenter, 2 * PI, True)
        Set oLineString = Nothing
    Else

        Set ObjBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parBodyOD, False)

    End If
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

    'Place Nozzle 1

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoCenter - cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    'Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    oPlacePoint.Set parFace1toCenter + cptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
