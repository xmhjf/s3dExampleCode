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
'   Author:         MA
'   Creation Date:  Monday, Oct 15 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is developed for CR-127627
'   Source: Saunders Aseptic Diaphragm Valves, ZDT Zero Deadleg ‘T’ Pattern
'
'   Change History:
'   dd.mmm.yyyy         who         change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Sub Class_Initialize()
'''
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
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parEndLength As Double
    Dim parBranchLength As Double
    Dim parPortVOffset As Double
    Dim parOpCentertoEnd As Double
    Dim parValCentertoOpCenter As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double
    Dim parPositionerOffset As Double
    Dim parWidth As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double

    Dim objInsCyl As Object
    Dim objInsBox As Object
    Dim dInsCyldia As Double
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim objInsulatedNozzle1 As Object
    Dim objInsulatedNozzle2 As Object
    Dim objInsulatedNozzle3 As Object
    Dim parInsulationDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parFacetoCenter = arrayOfInputs(3)
'    parFace1toCenter = arrayOfInputs(4)
'    parFace2toCenter = arrayOfInputs(5)
    parEndLength = arrayOfInputs(6)
    parBranchLength = arrayOfInputs(7)
    parPortVOffset = arrayOfInputs(8)
    parOpCentertoEnd = arrayOfInputs(9)
    parValCentertoOpCenter = arrayOfInputs(10)
    parInsulationThickness = arrayOfInputs(11)
    parHandwheelAngle = arrayOfInputs(12)
    parPositionerOffset = arrayOfInputs(13)
    parWidth = arrayOfInputs(14)

    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        sptOffset, depth

 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -5): FacetoFace
'   (Part data Basis value -10): FacetoCenter
'   (Part data Basis value -13): Face1toCenter and Face2toCenter
    
    Select Case lPartDataBasis
    
        Case Is <= 1, 5
            parFacetoFace = arrayOfInputs(2)
            dFace1toCenter = parFacetoFace / 2
            dFace2toCenter = parFacetoFace / 2
    
        Case 10
            parFacetoCenter = arrayOfInputs(3)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
        
        Case 13
            parFace1toCenter = arrayOfInputs(4)
            parFace2toCenter = arrayOfInputs(5)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
    
        Case Else
            GoTo ErrorLabel:
        
    End Select
    
    'Insert your code for output 9(Cylinder 1 Insulation)
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set -dFace1toCenter, 0, 0
    oEnPoint.Set dFace2toCenter, 0, 0
    dInsCyldia = pipeDiam + 2 * parInsulationThickness

    Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCyldia, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
    Set objInsCyl = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
            
    'Insert your code for output 10(Box 1 Insulation)
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set -(dFace1toCenter - parEndLength) - parInsulationThickness, _
                                2.5 * parPortVOffset + parInsulationThickness, _
                                    -parValCentertoOpCenter - parInsulationThickness
    oEnPoint.Set (dFace2toCenter - parEndLength) + parInsulationThickness, _
                            -parValCentertoOpCenter - parInsulationThickness, _
                                0.8 * parValCentertoOpCenter + parInsulationThickness
    
    Set objInsBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBox
    Set objInsBox = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    'Insert your code for output 11(Box 2 Insulation)
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set -(dFace1toCenter - parEndLength) - parInsulationThickness, _
                            2.5 * parPortVOffset + parInsulationThickness, _
                                0.8 * parValCentertoOpCenter + parInsulationThickness
    oEnPoint.Set (dFace2toCenter - parEndLength) + parInsulationThickness, _
                            -2 * parPortVOffset - parInsulationThickness, _
                    (parValCentertoOpCenter + parOpCentertoEnd - parBranchLength) _
                                    + parInsulationThickness
   
    Set objInsBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBox
    Set objInsBox = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
            
    'Insert your code for output 12(Cylinder 2 Insulation)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        sptOffset, depth
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set 0, parPortVOffset, _
                        (parValCentertoOpCenter + parOpCentertoEnd - parBranchLength)
    oEnPoint.Set 0, parPortVOffset, (parValCentertoOpCenter + parOpCentertoEnd)
    dInsCyldia = pipeDiam + 2 * parInsulationThickness

    Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCyldia, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
    Set objInsCyl = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    'Place Insulation for Port 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set -dFace1toCenter, 0, 0
    oEnPoint.Set -dFace1toCenter + flangeThick + parInsulationThickness, 0, 0

    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If

    Set objInsulatedNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                parInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle1
    Set objInsulatedNozzle1 = Nothing

    'Place Insulation for Port 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
    oStPoint.Set dFace2toCenter, 0, 0
    oEnPoint.Set dFace2toCenter - flangeThick - parInsulationThickness, 0, 0

    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If

    Set objInsulatedNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                        parInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle2
    Set objInsulatedNozzle2 = Nothing

    'Place Insulation for Port 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
    oStPoint.Set 0, parPortVOffset, (parValCentertoOpCenter + parOpCentertoEnd)
    oEnPoint.Set 0, parPortVOffset, (parValCentertoOpCenter + parOpCentertoEnd) _
                                    - flangeThick - parInsulationThickness
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If

    Set objInsulatedNozzle3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                       parInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle3
    Set objInsulatedNozzle3 = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub


