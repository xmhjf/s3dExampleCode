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
'   Creation Date:  Monday, Sep 24 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is developed for CR-127634
'   Source:
'   For Part data Basis value -416: Saunders Aseptic Diaphragm Valves, Close-Coupled
'                                                           Branch Valves.
'   For Part data Basis value -417: Saunders Aseptic Diaphragm Valves,Point-of-Use
'                                      Options (Close-Coupled Branch Valve ‘U’ Bend).
'
'   Change History:
'   dd.mmm.yyyy         who         change description
'   -----------         -----        ------------------
'  07.Nov.2007            MA        TR-128456: Provided a check on end points of non flanged insulation port cylinders
'                                                in which case small cylinders of negligible thickness will be created
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
    Dim parOpCentertoEnd As Double
    Dim parValCentertoOpCenter As Double
    Dim parPortVOffset As Double
    Dim parPorttoPortLenOff As Double
    Dim parPorttoCenterHeiOff As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double
    Dim parPositionerOffset As Double
    Dim parWidth As Double
    
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
'    parOpCentertoEnd = arrayOfInputs(3)
'    parValCentertoOpCenter = arrayOfInputs(4)
'    parPortVOffset = arrayOfInputs(5)
'    parPorttoPortLenOff = arrayOfInputs(6)
'    parPorttoCenterHeiOff = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    parHandwheelAngle = arrayOfInputs(9)
    parPositionerOffset = arrayOfInputs(10)
    parWidth = arrayOfInputs(11)
    
    iOutput = 0

 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    Select Case lPartDataBasis

        Case Is <= 1, 416
        
            parFacetoFace = arrayOfInputs(2)
            parOpCentertoEnd = arrayOfInputs(3)
            parValCentertoOpCenter = arrayOfInputs(4)
            parPortVOffset = arrayOfInputs(5)
            
            'Insert your code for output 9(Cylinder 1 Insulation)
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
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
            oStPoint.Set -0.3 * parFacetoFace - parInsulationThickness, _
                                    2.5 * parPortVOffset + parInsulationThickness, _
                                    -parValCentertoOpCenter - parInsulationThickness
            oEnPoint.Set 0.3 * parFacetoFace + parInsulationThickness, _
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
            oStPoint.Set -0.3 * parFacetoFace - parInsulationThickness, _
                                2.5 * parPortVOffset + parInsulationThickness, _
                                0.8 * parValCentertoOpCenter + parInsulationThickness
            oEnPoint.Set 0.3 * parFacetoFace + parInsulationThickness, _
                                    -0.4 * parPortVOffset - parInsulationThickness, _
                            (parValCentertoOpCenter + 0.6 * parOpCentertoEnd) + parInsulationThickness
            
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
            oStPoint.Set 0, parPortVOffset, (parValCentertoOpCenter + 0.6 * parOpCentertoEnd)
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
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0

            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If
            
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set -parFacetoFace / 2 + NEGLIGIBLE_THICKNESS, 0, 0
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
            oStPoint.Set parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0

            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If
            
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set parFacetoFace / 2 - NEGLIGIBLE_THICKNESS, 0, 0
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
            oStPoint.Set 0, parPortVOffset, (parValCentertoOpCenter + parOpCentertoEnd)
            oEnPoint.Set 0, parPortVOffset, _
                            (parValCentertoOpCenter + parOpCentertoEnd) - flangeThick - _
                                        parInsulationThickness

            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If

            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set 0, parPortVOffset, _
                                (parValCentertoOpCenter + parOpCentertoEnd) - NEGLIGIBLE_THICKNESS
            End If
            
            Set objInsulatedNozzle3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                parInsulationDiameter, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedNozzle3
            Set objInsulatedNozzle3 = Nothing
           
        Case 417
        
            parOpCentertoEnd = arrayOfInputs(3)
            parValCentertoOpCenter = arrayOfInputs(4)
            parPortVOffset = arrayOfInputs(5)
            parPorttoPortLenOff = arrayOfInputs(6)
            parPorttoCenterHeiOff = arrayOfInputs(7)
            
            'Insert your code for output 9(Box 1 Insulation)
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.15 * parPorttoPortLenOff - parInsulationThickness, _
                                    2.5 * parPortVOffset + parInsulationThickness, _
                                    -parValCentertoOpCenter - parInsulationThickness
            oEnPoint.Set 0.15 * parPorttoPortLenOff + parInsulationThickness, _
                                    -parValCentertoOpCenter - parInsulationThickness, _
                                    0.8 * parValCentertoOpCenter + parInsulationThickness
            
            Set objInsBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBox
            Set objInsBox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Insert your code for output 10(Box 2 Insulation)
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.15 * parPorttoPortLenOff - parInsulationThickness, _
                                2.5 * parPortVOffset + parInsulationThickness, _
                                0.8 * parValCentertoOpCenter + parInsulationThickness
            oEnPoint.Set 0.15 * parPorttoPortLenOff + parInsulationThickness, _
                                    -0.4 * parPortVOffset - parInsulationThickness, _
                                (parValCentertoOpCenter + 0.6 * parOpCentertoEnd) + parInsulationThickness
            
            Set objInsBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBox
            Set objInsBox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Insert your code for output 11(Cylinder 1 Insulation)
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, parPortVOffset, (parValCentertoOpCenter + 0.6 * parOpCentertoEnd)
            oEnPoint.Set 0, parPortVOffset, (parValCentertoOpCenter + parOpCentertoEnd)
            dInsCyldia = pipeDiam + 2 * parInsulationThickness

            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCyldia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Insert your code for output 12
            Dim oHorLine   As IngrGeom3D.Line3d
            Dim oArc  As IngrGeom3D.Arc3d
            Dim oVerLine   As IngrGeom3D.Line3d
    
            'Horizontal line
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.15 * parPorttoPortLenOff, 0, 0
            oEnPoint.Set -0.25 * parPorttoPortLenOff, 0, 0
         
            Set oHorLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
            'Arc
            Dim oCenPoint As AutoMath.DPosition
            Set oCenPoint = New AutoMath.DPosition
            Dim oArcPoint As AutoMath.DPosition
            oStPoint.Set -0.25 * parPorttoPortLenOff, 0, 0
            oEnPoint.Set -0.5 * parPorttoPortLenOff, 0, -0.25 * parPorttoPortLenOff
            oCenPoint.Set -0.25 * parPorttoPortLenOff, 0, -0.25 * parPorttoPortLenOff
            
            Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
            
            'Vertical line
            oStPoint.Set -0.5 * parPorttoPortLenOff, 0, -0.25 * parPorttoPortLenOff
            oEnPoint.Set -0.5 * parPorttoPortLenOff, 0, -parPorttoCenterHeiOff
            
            Set oVerLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
            Dim oSweepCurve         As IngrGeom3D.ComplexString3d
            Dim oSweepCurveCol      As Collection
            Set oSweepCurveCol = New Collection
            oSweepCurveCol.Add oHorLine
            oSweepCurveCol.Add oArc
            oSweepCurveCol.Add oVerLine
  
            Dim oStartBC   As AutoMath.DPosition
            Set oStartBC = New AutoMath.DPosition
            oStartBC.Set -0.15 * parPorttoPortLenOff, 0, 0
    
            Set oSweepCurve = PlaceTrCString(oStartBC, oSweepCurveCol)
            Dim oDirProj As AutoMath.DVector
            Set oDirProj = New AutoMath.DVector
    
            Dim oCircle As IngrGeom3D.Circle3d
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            -0.15 * parPorttoPortLenOff, 0, 0, _
                                -1, 0, 0, pipeDiam / 2 + parInsulationThickness)
    
            Dim Surfset   As IngrGeom3D.IJElements
            Dim stnorm() As Double
            Dim ednorm() As Double

            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                              oSweepCurve, oCircle, CircularCorner, 0, stnorm, ednorm, False)

            Dim ObjSurface As Object
            For Each ObjSurface In Surfset
                iOutput = iOutput + 1
                m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface
            Next ObjSurface
    
            'Insert your code for output 13
            'Horizontal line
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.15 * parPorttoPortLenOff, 0, 0
            oEnPoint.Set 0.25 * parPorttoPortLenOff, 0, 0
         
            Set oHorLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
            'Arc
            Set oCenPoint = New AutoMath.DPosition
            oStPoint.Set 0.25 * parPorttoPortLenOff, 0, 0
            oEnPoint.Set 0.5 * parPorttoPortLenOff, 0, -0.25 * parPorttoPortLenOff
            oCenPoint.Set 0.25 * parPorttoPortLenOff, 0, -0.25 * parPorttoPortLenOff
            
            Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
            
            'Vertical line
            oStPoint.Set 0.5 * parPorttoPortLenOff, 0, -0.25 * parPorttoPortLenOff
            oEnPoint.Set 0.5 * parPorttoPortLenOff, 0, -parPorttoCenterHeiOff
            
            Set oVerLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
            Set oSweepCurveCol = New Collection
            oSweepCurveCol.Add oHorLine
            oSweepCurveCol.Add oArc
            oSweepCurveCol.Add oVerLine
    
            Set oStartBC = New AutoMath.DPosition
            oStartBC.Set 0.15 * parPorttoPortLenOff, 0, 0
    
            Set oSweepCurve = PlaceTrCString(oStartBC, oSweepCurveCol)
            Set oDirProj = New AutoMath.DVector
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0.15 * parPorttoPortLenOff, 0, 0, _
                                1, 0, 0, pipeDiam / 2 + parInsulationThickness)
    
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                             oSweepCurve, oCircle, CircularCorner, 0, stnorm, ednorm, False)
            
            For Each ObjSurface In Surfset
                iOutput = iOutput + 1
                m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface
            Next ObjSurface
    
            Set oHorLine = Nothing
            Set oArc = Nothing
            Set oVerLine = Nothing
            Set oSweepCurve = Nothing
            Set oSweepCurveCol = Nothing
            Set oStartBC = Nothing
            Set oDirProj = Nothing
            Set oCircle = Nothing
            Set Surfset = Nothing
            Set oGeomFactory = Nothing
            
            'Place Insulation for Port 1
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.5 * parPorttoPortLenOff, 0, -parPorttoCenterHeiOff
            oEnPoint.Set -0.5 * parPorttoPortLenOff, _
                                0, -parPorttoCenterHeiOff + flangeThick + parInsulationThickness

            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If

            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set -0.5 * parPorttoPortLenOff, _
                                    0, -parPorttoCenterHeiOff + NEGLIGIBLE_THICKNESS
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
            oStPoint.Set 0.5 * parPorttoPortLenOff, 0, -parPorttoCenterHeiOff
            oEnPoint.Set 0.5 * parPorttoPortLenOff, _
                                0, -parPorttoCenterHeiOff + flangeThick + parInsulationThickness

            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If

            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set 0.5 * parPorttoPortLenOff, _
                                    0, -parPorttoCenterHeiOff + NEGLIGIBLE_THICKNESS
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
            oStPoint.Set 0, parPortVOffset, parValCentertoOpCenter + parOpCentertoEnd
            oEnPoint.Set 0, parPortVOffset, _
                            parValCentertoOpCenter + parOpCentertoEnd - flangeThick - parInsulationThickness

            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If

            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set 0, parPortVOffset, _
                                parValCentertoOpCenter + parOpCentertoEnd - NEGLIGIBLE_THICKNESS
            End If
            
            Set objInsulatedNozzle3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                parInsulationDiameter, True)

            'Set the output
            m_OutputColl.AddOutput "InsNoz3", objInsulatedNozzle3
            Set objInsulatedNozzle3 = Nothing
      
        Case Else
            GoTo ErrorLabel:
    End Select
            
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
       
End Sub


