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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         dkl
'   Creation Date:  Tuesday, 10, Jul 2007
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The symbol is modelled per Model No VC 100 of Hayward Marine Catalog.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ----   ---------------------------------------
'   10.Jul.2007     dkl  CR-123054  Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
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

    Dim parValveHeight As Double
    Dim parDiameter As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parValveHeight = arrayOfInputs(2)
    parDiameter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0
'Assumption : Origin is at Port 1.

' Insert your code for Insualted Cover at the base (output 5).
    'Assumption : The height of the cylinder is 10% of the Valve Height.
    Dim oStPoint As AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    oStPoint.Set parValveHeight, 0, 0
    Set oEndPoint = New AutoMath.DPosition
    oEndPoint.Set 0.9 * parValveHeight, 0, 0
    
    Dim objInsBottomCover As Object
    Set objInsBottomCover = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, _
                                        parDiameter + 2 * parInsulationThickness, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBottomCover
    Set objInsBottomCover = Nothing

' Insert your code for middle portion of the body of the check valve (output 6).
    'Assumption : 1) The height of the cylinderical portion is 55% of the Valve Height.
    '             2) The diameter of the cylinderical portion is 80% of the Valve Diameter.
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                                flangeDiam, sptOffset, depth
    
    Dim dRadiusMiddlePortion As Double
    dRadiusMiddlePortion = 0.4 * parDiameter + parInsulationThickness
    Dim dInsualtionDia As Double
    dInsualtionDia = pipeDiam + 2 * parInsulationThickness

'Points specifying the B-Spline curve are approximated to produce closer geometry.
    Dim oPoints(1 To 5) As New AutoMath.DPosition
    oPoints(1).Set 0.9 * parValveHeight, dRadiusMiddlePortion, 0
    oPoints(2).Set 0.55 * parValveHeight, dRadiusMiddlePortion, 0
    oPoints(3).Set 0.35 * parValveHeight, dRadiusMiddlePortion, 0
    oPoints(4).Set 0.21 * parValveHeight, (Abs(dRadiusMiddlePortion - dInsualtionDia / 2) / 2.1 + dInsualtionDia / 2), 0
    oPoints(5).Set 0.2 * parValveHeight, dInsualtionDia / 2, 0
    
    Dim objBsplineCurve As IngrGeom3D.BSplineCurve3d
    Set objBsplineCurve = PlaceTrBspline(4, oPoints)
    
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    oAxisVec.Set 1, 0, 0
    Dim oCenPoint As AutoMath.DPosition
    Set oCenPoint = New AutoMath.DPosition
    oCenPoint.Set 0, 0, 0
    
    Dim objInsMiddlePortion As Object
    Set objInsMiddlePortion = PlaceRevolution(m_OutputColl, objBsplineCurve, oAxisVec, _
                                                oCenPoint, 8 * Atn(1), True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsMiddlePortion
    Set objInsMiddlePortion = Nothing
    Set oCenPoint = Nothing
    Set oAxisVec = Nothing
    Set objBsplineCurve = Nothing

' Insert your code for Top Cylinder and the port (output 7).
    'Assumption : The height of this cylinder is 20% of the Valve Height.
    oStPoint.Set 0, 0, 0
    oEndPoint.Set 0.2 * parValveHeight, 0, 0
    
    If CmpDblGreaterthanOrEqualTo(flangeDiam, pipeDiam) Then _
                            dInsualtionDia = flangeDiam + 2 * parInsulationThickness
        
    Dim objInsTopCylinder As Object
    Set objInsTopCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, _
                                            dInsualtionDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTopCylinder
    Set objInsTopCylinder = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
