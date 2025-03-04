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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: dkl
'   Creation Date:  Tuesday, 10, Jul 2007
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The symbol is modelled per Model No VC 100 of Hayward Marine Catalog.
'
'   Change History:
'   1)Second port is added for venting purpose.
'   dd.mmm.yyyy     who     change description
'   -----------     ----   ---------------------------------------
'   10.Jul.2007     dkl  CR-123054  Created the symbol.
'   06.Aug.2008     MP     TR-141556   Inverted vent check valve data has incorrect Piping Commodity Value
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'
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

' Insert your code for Cover at the base (output 1).
    'Assumption : The height of the cylinder is 10% of the Valve Height.
    Dim oStPoint As AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    oStPoint.Set parValveHeight, 0, 0
    Set oEndPoint = New AutoMath.DPosition
    oEndPoint.Set 0.9 * parValveHeight, 0, 0
    
    Dim objBottomCover As Object
    Set objBottomCover = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, parDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBottomCover
    Set objBottomCover = Nothing

' Insert your code for middle portion of the body of the check valve (output 2).
    'Assumption : 1) The height of the cylinderical portion is 55% of the Valve Height.
    '             2) The diameter of the cylinderical portion is 80% of the Valve Diameter.
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                                flangeDiam, sptOffset, depth
    
    Dim dRadiusMiddlePortion As Double
    dRadiusMiddlePortion = 0.4 * parDiameter
'Points specifying the B-Spline curve are approximated to produce closer geometry.
    Dim oPoints(1 To 5) As New AutoMath.DPosition
    oPoints(1).Set 0.9 * parValveHeight, dRadiusMiddlePortion, 0
    oPoints(2).Set 0.55 * parValveHeight, dRadiusMiddlePortion, 0
    oPoints(3).Set 0.35 * parValveHeight, dRadiusMiddlePortion, 0
    oPoints(4).Set 0.21 * parValveHeight, (Abs(dRadiusMiddlePortion - pipeDiam / 2) / 2.1 + pipeDiam / 2), 0
    oPoints(5).Set 0.2 * parValveHeight, pipeDiam / 2, 0
    
    Dim objBsplineCurve As IngrGeom3D.BSplineCurve3d
    Set objBsplineCurve = PlaceTrBspline(4, oPoints)
    
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    oAxisVec.Set 1, 0, 0
    Dim oCenPoint As AutoMath.DPosition
    Set oCenPoint = New AutoMath.DPosition
    oCenPoint.Set 0, 0, 0
    
    Dim objMiddlePortion As Object
    Set objMiddlePortion = PlaceRevolution(m_OutputColl, objBsplineCurve, oAxisVec, _
                                                oCenPoint, 8 * Atn(1), True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMiddlePortion
    Set objMiddlePortion = Nothing
    Set oCenPoint = Nothing
    Set oAxisVec = Nothing
    Set objBsplineCurve = Nothing

' Insert your code for Top Cylinder adjoining the port (output 3).
    'Assumption : The height of this cylinder is 20% of the Valve Height.
    oStPoint.Set flangeThick, 0, 0
    oEndPoint.Set 0.2 * parValveHeight, 0, 0
    
    Dim objTopCylinder As Object
    Set objTopCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, pipeDiam, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTopCylinder
    Set objTopCylinder = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    oPlacePoint.Set depth - sptOffset, 0, 0
    Set oDir = New AutoMath.DVector
    oDir.Set -1, 0, 0
  
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

' Place Nozzle 2
  
    Dim objNozzle2   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    oPlacePoint.Set parValveHeight, 0, 0
    Set oDir = New AutoMath.DVector
    oDir.Set 1, 0, 0
    Set objNozzle2 = CreateNozzleJustaCircle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
'   Set the output
'   iOutput = iOutput + 1
    m_OutputColl.AddOutput "Nozzle2", objNozzle2
    Set objNozzle2 = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

