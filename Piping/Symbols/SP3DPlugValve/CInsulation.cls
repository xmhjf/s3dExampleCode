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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Thursday, Jan 24 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   09.Jul.2003     SymbolTeam(India)  Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav           Defined new input parameter HandwheelAngle.
'   02.Jan.2006     kkk                CR-91106 Updated the symbol to be based on either face-to-face
'                                               or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)  Modified to accept default value of lPartDataBasis to be less than or equal to 1.
'   25.sep.06       KKC                TR-102869 Modified symbol code to use Face to Center or Face to Face to Dimension.
'                                                Modified symbol code to use Face to Face as a input parameter.
'   18.Jul.2007     MA                 CR-123497 Implemented Part data basis to use Face 1 to Center and Face 2 to Center Dimensions.
'   14.Feb.2008     RRK                CR-33401  Provided implmentation for part data basis values-Plug valve, 2-way, partial jacket,
'                                                 specified by face-to-face dimension(815),Plug valve, 2-way, full jacket, oversized flanges, specified by face-to-face dimension(820),
'                                                 and Plug valve, 2-way, full jacket, standard flanges, specified by face-to-face dimension(821)
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

    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double

    Dim pipeDiam3        As Double
    Dim flangeThick3     As Double
    Dim sptOffset3       As Double
    Dim flangeDiam3      As Double
    Dim depth3           As Double

    Dim pipeDiam4        As Double
    Dim flangeThick4     As Double
    Dim sptOffset4       As Double
    Dim flangeDiam4      As Double
    Dim depth4           As Double

    Dim iOutput     As Double
    Dim ObjInsulCylinder As Object
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle   As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInletYOffset As Double
    Dim parInletZOffset As Double
    Dim parOutletXOffset As Double
    Dim parOutletYOffset As Double
    Dim parInletXOffset As Double
    Dim parInletOutletYOffset As Double
    Dim parInletOutletZOffset As Double
    Dim parLength As Double
    Dim parWidth As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parHandwheelAngle = arrayOfInputs(4)
'    parFacetoCenter = arrayOfInputs(5)
'    parFace1toCenter = arrayOfInputs(6)
'    parFace2toCenter = arrayOfInputs(7)
'    parInletYOffset = arrayOfInputs(8)
'    parInletZOffset = arrayOfInputs(9)
'    parOutletXOffset = arrayOfInputs(10)
'    parOutletYOffset = arrayOfInputs(11)
'    parInletOutletYOffset = arrayOfInputs(13)
'    parInletOutletZOffset = arrayOfInputs(14)
'    parLength = arrayOfInputs(15)
'    parWidth = arrayOfInputs(16)

    iOutput = 0
 ' Insert your code for output 5(Insul Cylinder)
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim InsulationDiameter As Double
    Dim InsulationDiameter1 As Double
    Dim InsulationDiameter2 As Double

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2

 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -5): Face-to-face dimension basis
'   (Part data Basis value -10): Face-to-center dimension basis
'   (Part data Basis value -13): Asymmetrical Face-to-Center dimension basis
'   (Part data Basis value -815): Plug valve, 2-way, partial jacket, specified by face-to-face dimension
'   (Part data Basis value -820): Plug valve, 2-way, full jacket, oversized flanges, specified by face-to-face dimension
'   (Part data Basis value -821): Plug valve, 2-way, full jacket, standard flanges, specified by face-to-face dimension

    Select Case lPartDataBasis

    Case Is <= 1, 5
        parFacetoFace = arrayOfInputs(2)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2

    Case 10
        parFacetoCenter = arrayOfInputs(5)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter

    Case 13
        parFace1toCenter = arrayOfInputs(6)
        parFace2toCenter = arrayOfInputs(7)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter

    Case 815
        parFacetoFace = arrayOfInputs(2)
        parInletYOffset = arrayOfInputs(8)
        parInletZOffset = arrayOfInputs(9)
        parOutletXOffset = arrayOfInputs(10)
        parOutletYOffset = arrayOfInputs(11)

        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2

        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, sptOffset3, depth3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam4, sptOffset4, depth4
        
    Case 820
        parFacetoFace = arrayOfInputs(2)
        parInletYOffset = arrayOfInputs(8)
        parInletZOffset = arrayOfInputs(9)
        parOutletYOffset = arrayOfInputs(11)
        parInletXOffset = arrayOfInputs(12)
        
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
        
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, sptOffset3, depth3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam4, sptOffset4, depth4
        
    Case 821
        parFacetoFace = arrayOfInputs(2)
        parInletOutletYOffset = arrayOfInputs(13)
        parInletOutletZOffset = arrayOfInputs(14)
        parLength = arrayOfInputs(15)
        parWidth = arrayOfInputs(16)

        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2

        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, sptOffset3, depth3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam4, sptOffset4, depth4

    Case Else
        GoTo ErrorLabel:
    End Select

'Place Insulation Body

    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        InsulationDiameter1 = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDiameter1 = pipeDiam + parInsulationThickness * 2
    End If

    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
        InsulationDiameter2 = flangeDiam2 + parInsulationThickness * 2
    Else
        InsulationDiameter2 = pipeDiam2 + parInsulationThickness * 2
    End If
    If CmpDblGreaterthan(InsulationDiameter1, InsulationDiameter2) Then
        InsulationDiameter = InsulationDiameter1
    Else
        InsulationDiameter = InsulationDiameter2
    End If

    Select Case lPartDataBasis
    Case Is <= 1, 5, 10, 13 'Setting end point for part data basis values 5, 10 and 13
        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set dFace2toCenter, 0, 0
        Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, InsulationDiameter, True)

    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
        Set ObjInsulCylinder = Nothing

    Case 815 'Creating outputs for part data basis value-Plug valve, 2-way, partial jacket,specified by face-to-face dimension
        Dim dJacketBoxLength As Double
        'Assumptions
        Dim dMaxJacketBoxLength As Double
        dMaxJacketBoxLength = (parFacetoFace - flangeThick - flangeThick2)

        'Calculating maximum diameter
        Dim dMaxDia As Double
         If CmpDblGreaterthan(flangeDiam4, pipeDiam4) Then
            dMaxDia = flangeDiam4
        Else
            dMaxDia = pipeDiam4
        End If

        Dim dMinJacketBoxLength As Double
        dMinJacketBoxLength = parFacetoFace - 2 * parOutletXOffset + 1.01 * dMaxDia

        dJacketBoxLength = (dMaxJacketBoxLength + dMinJacketBoxLength) / 2

        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set -dJacketBoxLength / 2, 0, 0
        Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, InsulationDiameter1, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
        Set ObjInsulCylinder = Nothing

        Dim ObjInsulCylinder2 As Object
        oStPoint.Set dJacketBoxLength / 2, 0, 0
        oEnPoint.Set dFace2toCenter, 0, 0
        Set ObjInsulCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            InsulationDiameter2, True)
        ' Set the output
        m_OutputColl.AddOutput "InsulationCyl2", ObjInsulCylinder2
        Set ObjInsulCylinder2 = Nothing

    'Place Jacket Body
        Dim objInsJacketBox As Object

        Dim dJacketBoxWidth As Double
        Dim dJacketBoxTopPos As Double
        Dim dJacketBoxBottomPos As Double

        'Setting jacket box width equal to twice of the Z-offset of inlet minus flange thickness of jacket inlet
        dJacketBoxWidth = 2 * (parInletZOffset - flangeThick3)

        'Check for case where flange thick ness of jacket inlet  is zero
        If CmpDblEqual(flangeThick3, 0) Then dJacketBoxWidth = 1.8 * parInletZOffset

        'Setting the jacket box top position
        If CmpDblGreaterthan(pipeDiam3, flangeDiam3) Then
            dJacketBoxTopPos = 1.01 * (parInletYOffset + (pipeDiam3 / 2))
        Else
            dJacketBoxTopPos = 1.01 * (parInletYOffset + (flangeDiam3 / 2))
        End If

        'Setting the jacket box bottom position
        dJacketBoxBottomPos = parOutletYOffset - flangeThick4
        If CmpDblEqual(flangeThick4, 0) Then dJacketBoxBottomPos = 0.9 * parOutletYOffset

        'Assumptions for jacket top box dimensions
        Dim dBoxTopPos As Double
        Dim dBoxBottomPos As Double
        dBoxBottomPos = dJacketBoxTopPos + 0.01 * pipeDiam
        dBoxTopPos = dBoxBottomPos + 0.2 * pipeDiam

        If CmpDblLessThan(dBoxTopPos, 0.5 * pipeDiam) Then
            dBoxTopPos = 0.55 * pipeDiam
        End If

        Dim dInsJacketBoxLength As Double
        dInsJacketBoxLength = dJacketBoxLength + 2 * parInsulationThickness

        If CmpDblGreaterthan(dInsJacketBoxLength, parFacetoFace) Then
            dInsJacketBoxLength = parFacetoFace
        End If

        oStPoint.Set -dInsJacketBoxLength / 2, -dJacketBoxBottomPos - parInsulationThickness, _
                                    -dJacketBoxWidth / 2 - parInsulationThickness
        oEnPoint.Set dInsJacketBoxLength / 2, dBoxTopPos + parInsulationThickness, _
                            dJacketBoxWidth / 2 + parInsulationThickness

        Set objInsJacketBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    ' Set the output
        m_OutputColl.AddOutput "JacketBoxInsulation", objInsJacketBox
        Set objInsJacketBox = Nothing
    
    Case 820 'Creating outputs for part data basis value-Plug valve, 2-way, full jacket, standard flanges, specified by face-to-face dimension
        Dim geomFactory As IngrGeom3D.GeometryFactory
        Set geomFactory = New IngrGeom3D.GeometryFactory
        Dim dLineStrPoints(0 To 14) As Double
        Dim oLine As IngrGeom3D.Line3d
        Dim oEllipArc As IngrGeom3D.EllipticalArc3d
        Dim oCircArc As IngrGeom3D.Arc3d
        Dim ObjCurveBody As Object
        Dim oComplexStr As IngrGeom3D.ComplexString3d
        Dim oComplexStr2 As IngrGeom3D.ComplexString3d
        Dim objCollection As Collection
        Set objCollection = New Collection
        
        Dim objCollection2 As Collection
        Set objCollection2 = New Collection
        
        'Computing the maximum of flange diameter and pipe diameter for inlet(port 1)
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            dMaxDia = flangeDiam
        Else
            dMaxDia = pipeDiam
        End If
        
        Dim dMaxDia3 As Double
        'Computing the maximum of flange diameter and pipe diameter of jacket inlet (port 3)
        If CmpDblGreaterthan(flangeDiam3, pipeDiam3) Then
            dMaxDia3 = flangeDiam3
        Else
            dMaxDia3 = pipeDiam3
        End If
        
        'Computing the curve width of complex string 1
        Dim dCurveWidth As Double
        If CmpDblEqual(flangeThick3, 0) Then
            dCurveWidth = 1.9 * parInletZOffset
        Else
            dCurveWidth = parInletZOffset - flangeThick3
        End If
        
        'Computing the curve bottom position of complex string 2
        Dim dCurveBottomPos As Double
        If CmpDblEqual(flangeThick4, 0) Then
            dCurveBottomPos = 1.9 * parOutletYOffset
        Else
            dCurveBottomPos = parOutletYOffset - flangeThick4
        End If
        
        Dim dMaxWidth1 As Double
        Dim dMaxWidth2 As Double
        
        'Computing limiting width of complex string 1
        dMaxWidth1 = dMaxDia / 2
        dMaxWidth2 = dCurveBottomPos
        
        Dim dLimitingWidth As Double
        dLimitingWidth = SmallerDim(dMaxWidth1, dMaxWidth2)
        
        'Check to see that width of complex string 1 doesn't exceed limiting width
        If CmpDblGreaterthan(dCurveWidth, dLimitingWidth) Then
            dCurveWidth = 0.85 * dLimitingWidth
        End If
        
        'Creating first complex string at flanged end(port 1) which consists of four elements- a line, an elliptical arc, a circular arc and an elliptical arc in order
        'Assumptions
        '1.The top most point is 1.05 times the Y-offset for inlet plus the maximum diameter for jacket inlet(port 3)
        '2.Width of the top portion is 70% of curve width of complex string 1
        '3.Radius of the circular arc is curve width of complex string 1
        
        dLineStrPoints(0) = -dFace1toCenter + flangeThick
        dLineStrPoints(1) = 1.05 * (parInletYOffset + dMaxDia3 / 2)
        dLineStrPoints(2) = -0.35 * dCurveWidth - parInsulationThickness

        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = dLineStrPoints(1)
        dLineStrPoints(5) = -dLineStrPoints(2)
        
        dLineStrPoints(6) = dLineStrPoints(0)
        dLineStrPoints(7) = 0
        dLineStrPoints(8) = dCurveWidth + parInsulationThickness
        
        
        dLineStrPoints(9) = dLineStrPoints(0)
        dLineStrPoints(10) = 0
        dLineStrPoints(11) = -dLineStrPoints(8)
        
        'Creating the element 1(line) of first complex string
        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                        dLineStrPoints(0), dLineStrPoints(1), dLineStrPoints(2), _
                        dLineStrPoints(3), dLineStrPoints(4), dLineStrPoints(5))

        objCollection.Add oLine
        Set oLine = Nothing
        
        'Creating the element 2(elliptical arc) of first complex string
        Dim mMRatio As Double
        mMRatio = (dLineStrPoints(8) - dLineStrPoints(5)) / (dLineStrPoints(1))
        
        Set oEllipArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                        dLineStrPoints(0), 0, dLineStrPoints(2), 1, 0, 0, _
                        0, dLineStrPoints(1), 0, mMRatio, 0, PI / 2)

        objCollection.Add oEllipArc
        Set oEllipArc = Nothing
        
        'Creating the element 3(circular arc) of first complex string
        Set oCircArc = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, dLineStrPoints(0), 0, 0, 1, 0, 0, _
                                            dLineStrPoints(6), dLineStrPoints(7), dLineStrPoints(8), _
                                            dLineStrPoints(9), dLineStrPoints(10), dLineStrPoints(11))
        
        objCollection.Add oCircArc
        Set oCircArc = Nothing
        
        'Creating the element 4(elliptical arc which is mirror of the first elliptical arc about X-Y plane)
        Set oEllipArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                         dLineStrPoints(0), 0, -dLineStrPoints(2), 1, 0, 0, _
                        0, dLineStrPoints(1), 0, mMRatio, -PI / 2, PI / 2)
        
        objCollection.Add oEllipArc
        Set oEllipArc = Nothing
        
        'Setting the first complex string
        oStPoint.Set dLineStrPoints(0), dLineStrPoints(1), dLineStrPoints(2)
        Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
        
        'Creating second complex string at a distance(to the left of the center) of half of the maximum diameter of jacket outlet.
        'It consists of four elements- line, first elliptical arc, second elliptical arc and third elliptical arc which is the mirror of first one about XY plane
        'Assumptions
        '1.The top most point is same as the top most point of the first complex string
        '2.Width of the top portion is same as the width of the first complex string
        '3.Maximum width of the compex string is taken as 1.1 times the maximum width of the first complex string
        Dim dLineStrPoints2(0 To 11) As Double
        
        'Computing the maximum of pipe and flange diameters for port 4
        Dim dMaxDia4 As Double
        If CmpDblGreaterthan(flangeDiam4, pipeDiam4) Then
            dMaxDia4 = flangeDiam4
        Else
            dMaxDia4 = pipeDiam4
        End If
            
        dLineStrPoints2(0) = -dMaxDia4 / 2
        dLineStrPoints2(1) = dLineStrPoints(1)
        dLineStrPoints2(2) = dLineStrPoints(2)

        dLineStrPoints2(3) = dLineStrPoints2(0)
        dLineStrPoints2(4) = dLineStrPoints2(1)
        dLineStrPoints2(5) = -dLineStrPoints2(2)
        
        dLineStrPoints2(6) = dLineStrPoints2(0)
        dLineStrPoints2(7) = 0
        dLineStrPoints2(8) = 1.1 * dCurveWidth + parInsulationThickness
        
        dLineStrPoints2(9) = dLineStrPoints2(0)
        dLineStrPoints2(10) = 0
        dLineStrPoints2(11) = -dLineStrPoints2(8)
        
        'Creating element 1(Line) of the second complex string
        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                        dLineStrPoints2(0), dLineStrPoints2(1), dLineStrPoints2(2), _
                        dLineStrPoints2(3), dLineStrPoints2(4), dLineStrPoints2(5))
        
        objCollection2.Add oLine
        Set oLine = Nothing
        
        'Creating element 2(elliptical arc) of the second complex string
        mMRatio = (dLineStrPoints2(8) - dLineStrPoints(5)) / dLineStrPoints(1)
        Set oEllipArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                        0, 0, dLineStrPoints(5), 1, 0, 0, _
                        0, dLineStrPoints(1), 0, mMRatio, 0, PI / 2)
        
        objCollection2.Add oEllipArc
        Set oEllipArc = Nothing
        
        'Creating element 3(elliptical arc) of the second complex string
        mMRatio = (dLineStrPoints2(8)) / (dCurveBottomPos + parInsulationThickness)
        Set oEllipArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                        0, 0, 0, 1, 0, 0, _
                        0, -dCurveBottomPos - parInsulationThickness, 0, mMRatio, -PI / 2, PI)
        
        objCollection2.Add oEllipArc
        Set oEllipArc = Nothing
        
        'Creating element 4(elliptical arc which is the mirror of element 2 about XY plane) of the second complex string
        mMRatio = (dLineStrPoints2(8) - dLineStrPoints(5)) / dLineStrPoints(1)
        Set oEllipArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                        0, 0, -dLineStrPoints(5), 1, 0, 0, _
                        0, dLineStrPoints(1), 0, mMRatio, -PI / 2, PI / 2)
        
        objCollection2.Add oEllipArc
        Set oEllipArc = Nothing
        
        'Setting the second complex string
        oStPoint.Set dLineStrPoints2(0), dLineStrPoints2(1), dLineStrPoints2(2)
        Set oComplexStr2 = PlaceTrCString(oStPoint, objCollection2)
        
        'Creating the ruled surface between the first and second complex strings
        Set ObjCurveBody = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oComplexStr, oComplexStr2, True)
        
    '     Set the output
        m_OutputColl.AddOutput "ObjCurveBody", ObjCurveBody
        Set ObjCurveBody = Nothing
        
        'Creating the projection of second complex string
        Dim oAxisVec As AutoMath.DVector
        Set oAxisVec = New DVector
        oAxisVec.Set 1, 0, 0
        
        Dim ObjProjectionBody As Object
        Set ObjProjectionBody = PlaceProjection(m_OutputColl, oComplexStr2, oAxisVec, dMaxDia4, True)
        
    '     Set the output
        m_OutputColl.AddOutput "ProjectionBody", ObjProjectionBody
        Set ObjProjectionBody = Nothing
    
        'Creating the other half of the body which is the mirror of the ruled surface about YZ plane
        Dim oTransformationMat  As AutoMath.DT4x4
        Set oTransformationMat = New AutoMath.DT4x4
        
        'Creating the mirror of two complex strings about YZ plane
        oTransformationMat.LoadIdentity
        oTransformationMat.IndexValue(0) = -1
        oComplexStr.Transform oTransformationMat
        oComplexStr2.Transform oTransformationMat
        
        Dim ObjCurveBody2 As Object
        Set ObjCurveBody2 = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                            oComplexStr, oComplexStr2, True)
    '   Set the output
        m_OutputColl.AddOutput "ObjCurveBody2", ObjCurveBody2
        Set ObjCurveBody2 = Nothing
        Set oTransformationMat = Nothing
        Set oComplexStr = Nothing
        Set oComplexStr2 = Nothing
        Set oAxisVec = Nothing
        Set objCollection = Nothing
        Set objCollection2 = Nothing
              
        'Creating the top box
        Dim objInsBox As Object
        Dim dBoxLength As Double
        Dim dBoxWidth  As Double
        
        dBoxLength = 0.6 * (parFacetoFace - flangeThick - flangeThick2)
        
        oStPoint.Set -dBoxLength / 2, dLineStrPoints(1), -dLineStrPoints(2)
        oEnPoint.Set dBoxLength / 2, dLineStrPoints(1) + 0.1 * dMaxDia + parInsulationThickness, dLineStrPoints(2)
        
        Set objInsBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        m_OutputColl.AddOutput "InsTopBox", objInsBox
        Set objInsBox = Nothing
        
        'Creating Insulation for port 1
        If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
            Dim ObjInsulatedPort1 As Object
            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -(dFace1toCenter - flangeThick - parInsulationThickness), 0, 0
            
            'Check to see that insulation doesn't exceed face to face dimension
            If CmpDblGreaterthan(flangeThick + parInsulationThickness, parFacetoFace) Then
                oEnPoint.Set dFace2toCenter, 0, 0
            End If
            Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, InsulationDiameter1, True)
        
            ' Set the output
            m_OutputColl.AddOutput "InsPort1", ObjInsulatedPort1
            Set ObjInsulatedPort1 = Nothing
        End If
        
        'Creating Insulation for port 2
        If CmpDblGreaterthan(flangeThick2, LINEAR_TOLERANCE) Then
            Dim ObjInsulatedPort2 As Object
            
            oStPoint.Set dFace2toCenter, 0, 0
            oEnPoint.Set (dFace2toCenter - flangeThick2 - parInsulationThickness), 0, 0
            
            'Check to see that header insulation doesn't exceed face to face dimension
            If CmpDblGreaterthan(flangeThick2 + parInsulationThickness, parFacetoFace) Then
                oEnPoint.Set -dFace1toCenter, 0, 0
            End If
            
            Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, InsulationDiameter2, True)
            ' Set the output
            m_OutputColl.AddOutput "InsPort2", ObjInsulatedPort2
            Set ObjInsulatedPort2 = Nothing
        End If
        
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set geomFactory = Nothing
        
    Case 821  'Creating outputs for part data basis value- Plug valve, 2-way, full jacket, standard flanges, specified by face-to-face dimension

        Set geomFactory = New IngrGeom3D.GeometryFactory

        Dim oArc As IngrGeom3D.Arc3d
        Dim ObjInsCurveBody As Object
        Set objCollection = New Collection
        
        'Computing the maximum of flange diameter and pipe diameter
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            dMaxDia = flangeDiam
        Else
            dMaxDia = pipeDiam
        End If
        
        'Creating the outline for jacket body in YZ plane which is projected along X-axis. It consists of three lines and a semi circular arc.
        'Assuming the top most point of jacket portion is 55% of maximum diameter for port 1
        dLineStrPoints(0) = -parLength / 2
        dLineStrPoints(1) = 0.55 * dMaxDia + parInsulationThickness
        dLineStrPoints(2) = -parWidth / 2 - parInsulationThickness

        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = dLineStrPoints(1)
        dLineStrPoints(5) = parWidth / 2 + parInsulationThickness

        dLineStrPoints(6) = dLineStrPoints(0)
        dLineStrPoints(7) = 0
        dLineStrPoints(8) = dLineStrPoints(5)

        dLineStrPoints(9) = dLineStrPoints(0)
        dLineStrPoints(10) = 0
        dLineStrPoints(11) = -parWidth / 2 - parInsulationThickness

        'Center
        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = 0
        dLineStrPoints(14) = 0

        'Creating Line 1
        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                        dLineStrPoints(0), dLineStrPoints(1), dLineStrPoints(2), _
                        dLineStrPoints(3), dLineStrPoints(4), dLineStrPoints(5))
        objCollection.Add oLine
        Set oLine = Nothing

        'Creating Line 2
        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                        dLineStrPoints(3), dLineStrPoints(4), dLineStrPoints(5), _
                        dLineStrPoints(6), dLineStrPoints(7), dLineStrPoints(8))
        objCollection.Add oLine
        Set oLine = Nothing

        'Creating Arc
        Set oArc = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                        dLineStrPoints(12), dLineStrPoints(13), dLineStrPoints(14), 1, 0, 0, _
                        dLineStrPoints(6), dLineStrPoints(7), dLineStrPoints(8), _
                        dLineStrPoints(9), dLineStrPoints(10), dLineStrPoints(11))

        objCollection.Add oArc
        Set oArc = Nothing

        'Creating Line 3
        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                        dLineStrPoints(9), dLineStrPoints(10), dLineStrPoints(11), _
                        dLineStrPoints(0), dLineStrPoints(1), dLineStrPoints(2))
        objCollection.Add oLine
        Set oLine = Nothing
        
        'Setting the complex string
        oStPoint.Set dLineStrPoints(0), dLineStrPoints(1), dLineStrPoints(2)
        Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
        
        'Creating the projection using the complex string
        Set oAxisVec = New DVector
        oAxisVec.Set 1, 0, 0
        Set ObjInsCurveBody = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, parLength, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCurveBody", ObjInsCurveBody
        Set ObjInsCurveBody = Nothing
        Set oAxisVec = Nothing
        Set oComplexStr = Nothing
        Set objCollection = Nothing

    'Place Box

        'Assumptions for jacket top box
        dBoxLength = 0.75 * parLength
        dBoxWidth = parWidth + 2 * parInsulationThickness
        dBoxBottomPos = parInletOutletYOffset - flangeThick3 + parInsulationThickness

        If CmpDblEqual(flangeThick3, 0) Then dBoxBottomPos = 1.9 * parInletOutletYOffset

        oStPoint.Set dBoxLength / 2, 0, dBoxWidth / 2
        oEnPoint.Set -dBoxLength / 2, -dBoxBottomPos, -dBoxWidth / 2

        Set objInsBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    ' Set the output
        m_OutputColl.AddOutput "InsBox", objInsBox
        Set objInsBox = Nothing
        Set geomFactory = Nothing
    End Select

    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
