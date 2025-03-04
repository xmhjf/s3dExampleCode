VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RRK
'   Creation Date:  Wednesday 30, Jul 2008
'   Description:

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   30.Jul.2008     RRK     CR-141774  Created the new symbol to include two options of pressure control valve
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
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
    
' Inputs
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
    
    'Setting the X and Y offsets of port 4 equal to that of port 3in the case of
    'part data basis value 'Pressure control valve, Type 2'(1039)
    If lPartDataBasis = 1039 Then
        parPort4XOffset = parPort3XOffset
        parPort4YOffset = parPort3YOffset
    End If
        
'Creating insulation for the curved body (Output 7)
    Dim objInsBody As Object
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    'Assumption:
    'Length of the curve body is assumed to be 80% of the face to face excluding flange thickness of two ports
    Dim dCurveBodyLength As Double
    dCurveBodyLength = 0.8 * (parFacetoFace - 2 * flangeThick)
    
    Dim dInsCurveBodyLength As Double
    dInsCurveBodyLength = dCurveBodyLength + 2 * parInsulationThickness
    
    'Setting the parameters for the elliptical arc
    Dim dHalfMajorAxis As Double
    Dim dHalfMinorAxis As Double
    Dim mMRatio As Double
    dHalfMajorAxis = dInsCurveBodyLength / 2
    dHalfMinorAxis = 0.5 * (parBodyOutsideDiameter - pipeDiam) + parInsulationThickness
    mMRatio = dHalfMinorAxis / dHalfMajorAxis
    
    'Creating the elliptical arc
    Dim oEllipArc As IngrGeom3D.EllipticalArc3d
    Set oEllipArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                    0, pipeDiam / 2, 0, 0, 0, -1, _
                    -dHalfMajorAxis, 0, 0, mMRatio, 0, PI)
    
    Dim oAxisVec As AutoMath.DVector
    Dim oCentPos As AutoMath.DPosition
    Set oAxisVec = New DVector
    Set oCentPos = New DPosition
    oAxisVec.Set 1, 0, 0
    oCentPos.Set 0, 0, 0
    Set objInsBody = PlaceRevolution(m_OutputColl, oEllipArc, oAxisVec, oCentPos, 2 * PI, True)
    Set oEllipArc = Nothing
    Set oAxisVec = Nothing
    Set oCentPos = Nothing
    Set geomFactory = Nothing
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
    Set objInsBody = Nothing
        
'Creating Insulation for Cylinder 1 (Output 8)
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim objInsCylinder1 As Object
    Dim dInsulationDia As Double
    Dim dNozzleLength As Double
    dNozzleLength = 0.5 * (parFacetoFace - dCurveBodyLength)
    
    oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
    oEnPoint.Set -parFacetoFace / 2 + dNozzleLength, 0, 0
    
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    Set objInsCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCylinder1
    Set objInsCylinder1 = Nothing
            
'Creating insulation for port 1(Variable Output)
    'Creating the output only when the end has flange or hub(when the flange thickness is greater than zero)
    If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
        Dim ObjInsulatedPort1 As Object
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
        
        'Check to see that insulation doesn't exceed face to face dimension
        If CmpDblGreaterthan(flangeThick + parInsulationThickness, parFacetoFace) Then
            oEnPoint.Set parFacetoFace / 2, 0, 0
        End If
        dInsulationDia = flangeDiam + 2 * parInsulationThickness
        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
        ' Set the output
        m_OutputColl.AddOutput "InsPort1", ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing
    End If
    
'Creating insulation Cylinder 2 (Output 9)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim objInsCylinder2 As Object

    oStPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
    oEnPoint.Set parFacetoFace / 2 - dNozzleLength, 0, 0
    
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    Set objInsCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCylinder2
    Set objInsCylinder2 = Nothing
    
'Creating insulation for port 2(Variable Output)
    'Creating the output only when the end has flange or hub(when the flange thickness is greater than zero)
    If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
        Dim ObjInsulatedPort2 As Object
        oStPoint.Set parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0
        
        'Check to see that insulation doesn't exceed face to face dimension
        If CmpDblGreaterthan(flangeThick + parInsulationThickness, parFacetoFace) Then
            oEnPoint.Set -parFacetoFace / 2, 0, 0
        End If
        dInsulationDia = flangeDiam + 2 * parInsulationThickness
        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
        ' Set the output
        m_OutputColl.AddOutput "InsPort2", ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing
    End If
    
'Creating insulation for port 3 (Output 10)
    Dim ObjInsulatedPort3  As Object
    
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oStPoint.Set -parFacetoFace / 2 + parPort3XOffset, parPort3YOffset, 0
    oEnPoint.Set -parFacetoFace / 2 + parPort3XOffset, 0, 0
    
    dInsulationDia = GreaterDim(pipeDiam, flangeDiam) + 2 * parInsulationThickness
    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing

'Creating insulation for port 4 (Output 11)
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim ObjInsulatedPort4  As Object
    
    oStPoint.Set parFacetoFace / 2 - parPort4XOffset, parPort4YOffset, 0
    oEnPoint.Set parFacetoFace / 2 - parPort4XOffset, 0, 0
    
    dInsulationDia = GreaterDim(pipeDiam, flangeDiam) + 2 * parInsulationThickness
    Set ObjInsulatedPort4 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort4
    Set ObjInsulatedPort4 = Nothing
    
'Creating insulation for port 5 (Output 12)
    RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim ObjInsulatedPort5  As Object
    
    oStPoint.Set -parFacetoFace / 2 + parPort5XOffset, 0, parPort5ZOffset
    oEnPoint.Set -parFacetoFace / 2 + parPort5XOffset, 0, 0
    
    dInsulationDia = GreaterDim(pipeDiam, flangeDiam) + 2 * parInsulationThickness
    Set ObjInsulatedPort5 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort5
    Set ObjInsulatedPort5 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
Private Function GreaterDim(ByVal Dim1 As Double, ByVal Dim2 As Double) As Double
    If CmpDblGreaterthanOrEqualTo(Dim1, Dim2) Then
        GreaterDim = Dim1
    Else
        GreaterDim = Dim2
    End If
End Function

