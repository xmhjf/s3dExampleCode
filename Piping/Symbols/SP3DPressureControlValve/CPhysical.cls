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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RRK
'   Creation Date:  Tuesday 24, Jul 2008
'   Description:

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   30.Jul.2008     RRK     CR-141774  Created the new symbol to include two options of pressure control valve
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub
Private Sub Class_Terminate()
    ''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parBodyOutsideDiameter As Double
    Dim parPort3XOffset As Double
    Dim parPort3YOffset As Double
    Dim parPort4XOffset As Double
    Dim parPort4YOffset As Double
    Dim parPort5XOffset As Double
    Dim parPort5ZOffset As Double
    Dim parInsulationThickness As Double
    
'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parBodyOutsideDiameter = arrayOfInputs(3)
    parPort3XOffset = arrayOfInputs(4)
    parPort3YOffset = arrayOfInputs(5)
    parPort4XOffset = arrayOfInputs(6)
    parPort4YOffset = arrayOfInputs(7)
    parPort5XOffset = arrayOfInputs(8)
    parPort5ZOffset = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)

    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    'Setting the X and Y offsets of port 4 equal to that of port 3 in the case of
    'part data basis value 'Pressure control valve, Type 2'(1039)
    If lPartDataBasis = 1039 Then
        parPort4XOffset = parPort3XOffset
        parPort4YOffset = parPort3YOffset
    End If
        
'Creating Curve Body (Output 1)
    Dim objBody As Object
    
    'Creating the elliptical arc
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim dHalfMajorAxis As Double
    Dim dHalfMinorAxis As Double
    Dim dNozzleLength As Double
    Dim dCurveBodyLength As Double
    Dim mMRatio As Double
    
    'Assumption:
    'Length of the curve body is assumed to be 80% of the face to face excluding flange thickness of two ports
    dCurveBodyLength = 0.8 * (parFacetoFace - 2 * flangeThick)
    
    'Setting the parameters for the elliptical curve
    dHalfMajorAxis = dCurveBodyLength / 2
    dHalfMinorAxis = 0.5 * (parBodyOutsideDiameter - pipeDiam)
    mMRatio = dHalfMinorAxis / dHalfMajorAxis
    
    'Creating the elliptical arc
    Dim oEllipArc As IngrGeom3D.EllipticalArc3d
    Set oEllipArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                    0, pipeDiam / 2, 0, 0, 0, -1, _
                    -dHalfMajorAxis, 0, 0, mMRatio, 0, PI)
    
    'Creating the revolution of the elliptical arc
    Dim oAxisVec As AutoMath.DVector
    Dim oCentPos As AutoMath.DPosition
    Set oAxisVec = New DVector
    Set oCentPos = New DPosition
    oAxisVec.Set 1, 0, 0
    oCentPos.Set 0, 0, 0
    
    Set objBody = PlaceRevolution(m_OutputColl, oEllipArc, oAxisVec, oCentPos, 2 * PI, True)
    Set oEllipArc = Nothing
    Set oCentPos = Nothing
    Set oAxisVec = Nothing
    Set geomFactory = Nothing
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
    Set objBody = Nothing
    
'Creating Nozzle 1 (Output 2)
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
        
    dNozzleLength = 0.5 * (parFacetoFace - dCurveBodyLength)
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

'Creating Nozzle 2 (Output 3)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

'Creating Nozzle 3 (Output 4)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set -parFacetoFace / 2 + parPort3XOffset, parPort3YOffset + sptOffset - depth, 0
    oDir.Set 0, 1, 0
    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, parPort3YOffset)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
'Creating Nozzle 4 (Output 5)
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set parFacetoFace / 2 - parPort4XOffset, parPort4YOffset + sptOffset - depth, 0
    oDir.Set 0, 1, 0
    Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, oPlacePoint, parPort4YOffset)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
'Creating Nozzle 5 (Output 6)
    RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set -parFacetoFace / 2 + parPort5XOffset, 0, parPort5ZOffset + sptOffset - depth
    oDir.Set 0, 0, 1
    
    Set objNozzle = CreateNozzleWithLength(5, oPartFclt, m_OutputColl, oDir, oPlacePoint, parPort5ZOffset)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

