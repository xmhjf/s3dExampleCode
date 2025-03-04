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
'   Creation Date:  Wednesday 13, Aug 2008

'   Description:
'   The symbol supports the following part data basis values:
'   1)Double basket strainer, Type 1 (994)
'   2)Double basket strainer, Type 2 (995)
'   3)Double basket strainer, Type 3a (996)
'   4)Double basket strainer, Type 3b (997)
'   5)Double basket strainer, Type 4 (998)
'   6)Double basket strainer, Type 5 (999)
'   7)Double basket strainer, Type 6 (1000)

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   13-Aug-2008     RRK     CR-143531  Created the new Symbol to implement sven part data basis values from 994 to 1000
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
    iOutput = 0

    Dim ObjCylinder As Object
    Dim objCover As Object

    Dim parA As Double
    Dim parB As Double
    Dim parC As Double
    Dim parD As Double
    Dim parE As Double
    Dim parF As Double
    Dim parG As Double
    Dim parH As Double
    Dim parI As Double
    Dim parXOffset As Double
    Dim parYOffset As Double
    Dim parZOffset As Double
    Dim parInsulationThickness As Double

    Dim dFacetoFace As Double
    Dim dWidth As Double
    Dim dCentertoBottom As Double
    Dim dCentertoTop As Double
    Dim dLeverLength As Double
    Dim dCentertoCenter As Double
    Dim dCylinderDiam As Double
    Dim dFacetoCenter As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dBasketChamberWidth As Double
    Dim dOverallWidth As Double
    Dim dCenter2toBottom As Double
    Dim dHeight As Double
    Dim dLength As Double
    Dim dHandleLength As Double
    Dim dCentertoEnd1 As Double
    Dim dCentertoEnd2 As Double
    Dim dCenter1toBottom As Double
    Dim dDrainXOffset As Double
    Dim dPort3XLocation As Double
    Dim dPort3ZLocation As Double
    Dim dPort4XLocation As Double
    Dim dPort4ZLocation As Double
    Dim dFace1toEnd As Double
    Dim dCentertoCylEnd As Double
    Dim dOperatorDia As Double
    Dim dHandleCentertoBottom As Double

    Set oPartFclt = arrayOfInputs(1)
' Inputs
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, sptOffset3, depth3
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam4, sptOffset4, depth4

    Select Case lPartDataBasis

    Case Is <= 1, 994
        dFacetoFace = arrayOfInputs(2) 'A
        dWidth = arrayOfInputs(3) 'B
        dCentertoBottom = arrayOfInputs(4) 'C
        dCentertoTop = arrayOfInputs(5) 'D
        dLeverLength = arrayOfInputs(6) 'E
        dCentertoCenter = arrayOfInputs(7) 'F
        parYOffset = arrayOfInputs(12)
        parInsulationThickness = arrayOfInputs(14)

        If CmpDblEqual(dCentertoTop, 0) Then dCentertoTop = 2 * pipeDiam
        If CmpDblEqual(dLeverLength, 0) Then dLeverLength = 0.5 * dWidth
        If CmpDblEqual(dCentertoCenter, 0) Then dCentertoCenter = 0.75 * dWidth

    Case 995
        dFacetoCenter = arrayOfInputs(2) 'A
        dBasketChamberWidth = arrayOfInputs(3) 'B
        dOverallWidth = arrayOfInputs(4) 'C
        dCenter2toBottom = arrayOfInputs(5) 'D
        dHeight = arrayOfInputs(6) 'E
        dCenter1toBottom = arrayOfInputs(7) 'F
        dWidth = arrayOfInputs(8) 'G
        parYOffset = arrayOfInputs(12)
        parInsulationThickness = arrayOfInputs(14)

    Case 996
        dFacetoFace = arrayOfInputs(2) 'A
        dFace2toCenter = arrayOfInputs(3) 'B
        dHeight = arrayOfInputs(4) 'C
        dCentertoBottom = arrayOfInputs(5) 'D
        dLeverLength = arrayOfInputs(6) 'E
        dCylinderDiam = arrayOfInputs(7) 'F
        parYOffset = arrayOfInputs(12)
        parZOffset = arrayOfInputs(13)
        parInsulationThickness = arrayOfInputs(14)

        If CmpDblEqual(dCylinderDiam, 0) Then dCylinderDiam = SmallerDim(dFacetoFace - dFace2toCenter, dFace2toCenter)

        dDrainXOffset = Sqr((dCylinderDiam / 2) ^ 2 - (parZOffset) ^ 2)
        dPort3XLocation = -dDrainXOffset - sptOffset3 + depth3

        dOverallWidth = 1.8 * dLeverLength
        dPort3ZLocation = (dOverallWidth / 2 - dCylinderDiam / 2 + parZOffset)

        dPort4XLocation = -dDrainXOffset - sptOffset4 + depth4
        dPort4ZLocation = -(dOverallWidth / 2 - dCylinderDiam / 2 + parZOffset)

    Case 997
        dFacetoFace = arrayOfInputs(2) 'A
        dFace2toCenter = arrayOfInputs(3) 'B
        dHeight = arrayOfInputs(4) 'C
        dCentertoBottom = arrayOfInputs(5) 'D
        dLeverLength = arrayOfInputs(6) 'E
        parYOffset = arrayOfInputs(12)
        parZOffset = arrayOfInputs(13)
        parInsulationThickness = arrayOfInputs(14)

        dOverallWidth = 2.6 * parZOffset
        dCylinderDiam = 2 * (dOverallWidth / 2 - parZOffset)

        dPort3XLocation = -dCylinderDiam / 2 - sptOffset3 + depth3
        dPort3ZLocation = parZOffset

        dPort4XLocation = -dCylinderDiam / 2 - sptOffset4 + depth4
        dPort4ZLocation = -parZOffset

    Case 998
        dFacetoFace = arrayOfInputs(2) 'A
        dCentertoBottom = arrayOfInputs(3)  'B
        dHeight = arrayOfInputs(4) 'C
        dCentertoEnd1 = arrayOfInputs(5) 'D
        dLength = arrayOfInputs(6) 'E
        dCentertoEnd2 = arrayOfInputs(7)  'F
        parYOffset = arrayOfInputs(12)
        parZOffset = arrayOfInputs(13)
        parInsulationThickness = arrayOfInputs(14)

    Case 999
        dFace1toCenter = arrayOfInputs(2) 'A
        dFace1toEnd = arrayOfInputs(3)  'B
        dCentertoCylEnd = arrayOfInputs(4) 'C
        dHeight = arrayOfInputs(5) 'D
        dCentertoBottom = arrayOfInputs(6)  'E
        dWidth = arrayOfInputs(7)   'F
        dFacetoFace = arrayOfInputs(8)    'G
        dOperatorDia = arrayOfInputs(9)    'H
        dHandleCentertoBottom = arrayOfInputs(10)    'I
        parXOffset = arrayOfInputs(11)
        parYOffset = arrayOfInputs(12)
        parZOffset = arrayOfInputs(13)
        parInsulationThickness = arrayOfInputs(14)

    Case 1000
        dWidth = arrayOfInputs(2) 'A
        dHandleLength = arrayOfInputs(3) 'B
        dCentertoBottom = arrayOfInputs(4) 'C
        dFace1toCenter = arrayOfInputs(5) 'D
        dFace2toCenter = arrayOfInputs(6) 'E
        parYOffset = arrayOfInputs(12)
        parZOffset = arrayOfInputs(13)
        parInsulationThickness = arrayOfInputs(14)
    Case Else
        GoTo ErrorLabel:
    End Select

    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

    Select Case lPartDataBasis
    Case Is <= 1, 994
    'Creating the Output 1(Cylinder 1)
        oStPoint.Set 0, -dCentertoBottom - parInsulationThickness, -dCentertoCenter / 2
        oEnPoint.Set 0, 0.5 * dCentertoTop + parInsulationThickness, -dCentertoCenter / 2
        
        dCylinderDiam = (dWidth - dCentertoCenter)
        
        Dim dInsCylinderDiam As Double
        dInsCylinderDiam = (dWidth - dCentertoCenter) + 2 * parInsulationThickness
        
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylinderDiam, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder 1", ObjCylinder
        Set ObjCylinder = Nothing

    'Creating the Output 2(Cylinder 2)
        Dim ObjCylinder2 As Object
        oStPoint.Set 0, -dCentertoBottom - parInsulationThickness, dCentertoCenter / 2
        oEnPoint.Set 0, 0.5 * dCentertoTop + parInsulationThickness, dCentertoCenter / 2

        Set ObjCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylinderDiam, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder 2", ObjCylinder2
        Set ObjCylinder2 = Nothing

    'Creating the Output 3(Box 1)
        Dim ObjBox As Object
        Dim dBoxWidth As Double
        Dim dBoxLength As Double

        dBoxWidth = 0.8 * (dCentertoCenter - dCylinderDiam)
        dBoxLength = 0.5 * dFacetoFace
        Dim dInsBoxLength As Double
        dInsBoxLength = dBoxLength + 2 * parInsulationThickness

        oStPoint.Set -dInsBoxLength / 2, -0.9 * dCentertoBottom - parInsulationThickness, -dBoxWidth / 2
        oEnPoint.Set dInsBoxLength / 2, 0.5 * dCentertoTop + parInsulationThickness, dBoxWidth / 2

        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    ' Set the output
        m_OutputColl.AddOutput "InsBox 1", ObjBox
        Set ObjBox = Nothing

    'Creating the Output 4(Box 2)
        Dim ObjBox2 As Object
        Dim dBox2Length As Double
        Dim dBox2Width As Double

        dBox2Length = 0.5 * dBoxLength
                
        If CmpDblGreaterthan(dBox2Length, dCylinderDiam) Then dBox2Length = 0.5 * dCylinderDiam
        Dim dInsBox2Length As Double
        dInsBox2Length = dBox2Length + 2 * parInsulationThickness
    
        dBox2Width = dCentertoCenter - 2 * Sqr((dCylinderDiam / 2) ^ 2 - (dBox2Length / 2) ^ 2)

        oStPoint.Set -dInsBox2Length / 2, -0.9 * dCentertoBottom - parInsulationThickness, -dBox2Width / 2
        oEnPoint.Set dInsBox2Length / 2, 0.4 * dCentertoTop + parInsulationThickness, dBox2Width / 2

        Set ObjBox2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    ' Set the output
        m_OutputColl.AddOutput "InsBox 2", ObjBox2
        Set ObjBox2 = Nothing

    'Creating the Output 5(Insulation for body Cylinder)
        Dim ObjInsBodyCylinder As Object

        oStPoint.Set -dFacetoFace / 2, 0, 0
        oEnPoint.Set dFacetoFace / 2, 0, 0

        Set ObjInsBodyCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBodyCylinder
        Set ObjInsBodyCylinder = Nothing

        'Creating the Output 9(Curve Body)
        ReDim dLineStrPoints(0 To 11) As Double
        Dim objCollection As Collection
        Set objCollection = New Collection
        Dim oComplexStr As IngrGeom3D.ComplexString3d
        Dim oLine As IngrGeom3D.Line3d
        Dim oArc As IngrGeom3D.Arc3d
        Dim ObjCurveBody As Object

        dLineStrPoints(0) = 0.6 * dBoxLength + parInsulationThickness
        dLineStrPoints(1) = 0
        dLineStrPoints(2) = -0.6 * pipeDiam

        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = -0.4 * dCentertoBottom
        dLineStrPoints(5) = dLineStrPoints(2)

        dLineStrPoints(6) = dLineStrPoints(0)
        dLineStrPoints(7) = dLineStrPoints(4)
        dLineStrPoints(8) = -dLineStrPoints(5)

        dLineStrPoints(9) = dLineStrPoints(0)
        dLineStrPoints(10) = 0
        dLineStrPoints(11) = dLineStrPoints(8)
        
        Dim geomFactory As IngrGeom3D.GeometryFactory
        Set geomFactory = New IngrGeom3D.GeometryFactory

        'Creating Line 1
        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                        dLineStrPoints(0), dLineStrPoints(1), dLineStrPoints(2), _
                        dLineStrPoints(3), dLineStrPoints(4), dLineStrPoints(5))
        objCollection.Add oLine
        Set oLine = Nothing

        'Creating Arc
        Set oArc = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                        dLineStrPoints(0), -0.4 * dCentertoBottom, 0, -1, 0, 0, _
                        dLineStrPoints(3), dLineStrPoints(4), dLineStrPoints(5), _
                        dLineStrPoints(6), dLineStrPoints(7), dLineStrPoints(8))

        objCollection.Add oArc
        Set oArc = Nothing

        'Creating Line 2
        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                        dLineStrPoints(6), dLineStrPoints(7), dLineStrPoints(8), _
                        dLineStrPoints(9), dLineStrPoints(10), dLineStrPoints(11))
        objCollection.Add oLine
        Set oLine = Nothing


        'Creating Arc
        Set oArc = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                        dLineStrPoints(0), 0, 0, -1, 0, 0, _
                        dLineStrPoints(9), dLineStrPoints(10), dLineStrPoints(11), _
                        dLineStrPoints(0), dLineStrPoints(1), dLineStrPoints(2))

        objCollection.Add oArc
        Set oArc = Nothing

        'Setting the complex string
        oStPoint.Set dLineStrPoints(0), dLineStrPoints(1), dLineStrPoints(2)
        Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
        Set objCollection = Nothing

        'Creating the projection using the complex string
        Dim oAxisVec As AutoMath.DVector
        Set oAxisVec = New DVector
        oAxisVec.Set -1, 0, 0
        Set ObjCurveBody = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, 0.1 * dBoxLength + parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsObjCurveBody", ObjCurveBody
        Set ObjCurveBody = Nothing
        Set objCollection = Nothing
        Set oAxisVec = Nothing
        Set oComplexStr = Nothing
    
    'Creating Insulation for port 1
        If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
            Dim ObjInsulatedPort1 As Object
            oStPoint.Set -dFacetoFace / 2, 0, 0
            oEnPoint.Set -(dFacetoFace / 2 - flangeThick - parInsulationThickness), 0, 0
            
        'Check to see that insulation doesn't exceed face to face dimension
        If CmpDblGreaterthan(flangeThick + parInsulationThickness, dFacetoFace) Then
                oEnPoint.Set dFacetoFace / 2, 0, 0
            End If
            Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam + 2 * parInsulationThickness, True)
        
            ' Set the output
            m_OutputColl.AddOutput "InsPort1", ObjInsulatedPort1
            Set ObjInsulatedPort1 = Nothing
        End If
        
        'Creating Insulation for port 2
        If CmpDblGreaterthan(flangeThick2, LINEAR_TOLERANCE) Then
            Dim ObjInsulatedPort2 As Object
            
            oStPoint.Set dFacetoFace / 2, 0, 0
            oEnPoint.Set (dFacetoFace / 2 - flangeThick2 - parInsulationThickness), 0, 0
            
            'Check to see that header insulation doesn't exceed face to face dimension
            If CmpDblGreaterthan(flangeThick2 + parInsulationThickness, dFacetoFace) Then
                oEnPoint.Set -dFacetoFace / 2, 0, 0
            End If
            
            Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam2 + 2 * parInsulationThickness, True)
            ' Set the output
            m_OutputColl.AddOutput "InsPort2", ObjInsulatedPort2
            Set ObjInsulatedPort2 = Nothing
        End If

        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set geomFactory = Nothing

    Case 995
    'Creating the Output 1(Cylinder 1)
        dCylinderDiam = dBasketChamberWidth - (dOverallWidth - dWidth)
        dInsCylinderDiam = dCylinderDiam + 2 * parInsulationThickness

        oStPoint.Set 0, -dCenter1toBottom - parInsulationThickness, -(dWidth - dCylinderDiam) / 2
        oEnPoint.Set 0, 0.5 * (dHeight - dCenter1toBottom), -(dWidth - dCylinderDiam) / 2

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylinderDiam, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder 1", ObjCylinder
        Set ObjCylinder = Nothing

    'Creating the Output 2(Cylinder 2)
        oStPoint.Set 0, -dCenter1toBottom - parInsulationThickness, (dWidth - dCylinderDiam) / 2
        oEnPoint.Set 0, 0.5 * (dHeight - dCenter1toBottom), (dWidth - dCylinderDiam) / 2

        Set ObjCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylinderDiam, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder 2", ObjCylinder2
        Set ObjCylinder2 = Nothing

    'Creating the Output 3(Cylinder 3)
        Dim ObjCylinder3 As Object
        oStPoint.Set 0, 0.5 * (dHeight - dCenter1toBottom), -(dWidth - dCylinderDiam) / 2
        oEnPoint.Set 0, 0.75 * (dHeight - dCenter1toBottom) + parInsulationThickness, -(dWidth - dCylinderDiam) / 2
        
        Dim dCylinderDiam3 As Double
        dCylinderDiam3 = dCylinderDiam + (0.2 * (dOverallWidth - dWidth))
        Dim dInsCylinderDiam3 As Double
        dInsCylinderDiam3 = dCylinderDiam3 + 2 * parInsulationThickness
        
        Set ObjCylinder3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylinderDiam3, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder 3", ObjCylinder3
        Set ObjCylinder3 = Nothing

    'Creating the Output 4(Cylinder 4)
        Dim ObjCylinder4 As Object
        oStPoint.Set 0, 0.5 * (dHeight - dCenter1toBottom), (dWidth - dCylinderDiam) / 2
        oEnPoint.Set 0, 0.75 * (dHeight - dCenter1toBottom) + parInsulationThickness, (dWidth - dCylinderDiam) / 2

        Set ObjCylinder4 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylinderDiam3, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder 4", ObjCylinder4
        Set ObjCylinder4 = Nothing

    'Creating the Output 5(Curve Body)
        Set geomFactory = New IngrGeom3D.GeometryFactory

        Dim ObjBody As Object
        Dim dWidthExclCylDiam As Double
        dWidthExclCylDiam = dWidth - (2 * dCylinderDiam)

        ReDim dLineStrPoints(0 To 23) As Double

        dLineStrPoints(0) = 0
        dLineStrPoints(1) = 0.55 * pipeDiam + parInsulationThickness
        dLineStrPoints(2) = -0.5 * dWidthExclCylDiam - dCylinderDiam / 2

        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = dLineStrPoints(1)
        dLineStrPoints(5) = -0.45 * dWidthExclCylDiam

        dLineStrPoints(6) = dLineStrPoints(0)
        dLineStrPoints(7) = 0.6 * pipeDiam + parInsulationThickness
        dLineStrPoints(8) = dLineStrPoints(5)

        dLineStrPoints(9) = dLineStrPoints(0)
        dLineStrPoints(10) = dLineStrPoints(7)
        dLineStrPoints(11) = -0.35 * dWidthExclCylDiam

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = 0.5 * pipeDiam + parInsulationThickness
        dLineStrPoints(14) = dLineStrPoints(11)

        dLineStrPoints(15) = dLineStrPoints(0)
        dLineStrPoints(16) = dLineStrPoints(13)
        dLineStrPoints(17) = -0.15 * dWidthExclCylDiam

        dLineStrPoints(18) = dLineStrPoints(0)
        dLineStrPoints(19) = 0.6 * pipeDiam + parInsulationThickness
        dLineStrPoints(20) = dLineStrPoints(17)

        dLineStrPoints(21) = dLineStrPoints(0)
        dLineStrPoints(22) = dLineStrPoints(19)
        dLineStrPoints(23) = 0
        
        Dim oLineString As IngrGeom3D.LineString3d
        Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 8, dLineStrPoints)

        Set oAxisVec = New DVector
        oAxisVec.Set 0, 0, 1
        oStPoint.Set 0, 0, 0
        Set ObjBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVec, oStPoint, 2 * PI, True)

    ' Set the output
        m_OutputColl.AddOutput "InsBody", ObjBody
        Set ObjBody = Nothing

    'Creating the Output 6(Body 2)
        Dim objBody2 As Object
        Dim oTransformationMat  As AutoMath.DT4x4
        Set oTransformationMat = New AutoMath.DT4x4

        oTransformationMat.LoadIdentity
        oTransformationMat.IndexValue(10) = -1
        oLineString.Transform oTransformationMat

        oAxisVec.Set 0, 0, 1
        oStPoint.Set 0, 0, 0

        Set objBody2 = PlaceRevolution(m_OutputColl, oLineString, oAxisVec, oStPoint, 2 * PI, True)
    ' Set the output
        m_OutputColl.AddOutput "InsBody2", objBody2
        Set objBody2 = Nothing

    'Creating the Output 7(Body 3)
        Dim objBody3 As Object
        Dim oDirVector As AutoMath.DVector
        Set oDirVector = New AutoMath.DVector
        oTransformationMat.LoadIdentity
        oDirVector.Set 0, -(dCenter1toBottom - dCenter2toBottom), 0
        oTransformationMat.Translate oDirVector
        oLineString.Transform oTransformationMat

        oAxisVec.Set 0, 0, 1
        oStPoint.Set 0, -(dCenter1toBottom - dCenter2toBottom), 0

        Set objBody3 = PlaceRevolution(m_OutputColl, oLineString, oAxisVec, oStPoint, 2 * PI, True)

    ' Set the output
        m_OutputColl.AddOutput "InsBody3", objBody3
        Set objBody3 = Nothing

    'Creating the Output 8(Body 4)
        Dim objBody4 As Object

        oTransformationMat.LoadIdentity
        oTransformationMat.IndexValue(10) = -1
        oLineString.Transform oTransformationMat

        oAxisVec.Set 0, 0, 1
        oStPoint.Set 0, -(dCenter1toBottom - dCenter2toBottom), 0

        Set objBody4 = PlaceRevolution(m_OutputColl, oLineString, oAxisVec, oStPoint, 2 * PI, True)
    ' Set the output
        m_OutputColl.AddOutput "InsBody4", objBody4
        Set objBody4 = Nothing

    'Creating the Output 9 (End Cylinder 1)
        Dim ObjEndCylinder1 As Object
        oStPoint.Set 0, 0, -0.5 * dWidthExclCylDiam - dCylinderDiam / 2
        oEnPoint.Set 0, 0, -0.55 * dWidthExclCylDiam - dCylinderDiam

        Set ObjEndCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsEnd Cylinder 1", ObjEndCylinder1
        Set ObjEndCylinder1 = Nothing

    'Creating the Output 10 (End Cylinder 2)
        Dim ObjEndCylinder2 As Object
        oStPoint.Set 0, 0, -0.55 * dWidthExclCylDiam - dCylinderDiam
        oEnPoint.Set 0, 0, -0.65 * dWidthExclCylDiam - dCylinderDiam - parInsulationThickness

        Set ObjEndCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.2 * pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsEnd Cylinder 2", ObjEndCylinder2
        Set ObjEndCylinder2 = Nothing

    'Creating the Output 11 (End Cylinder 3)
        Dim ObjEndCylinder3 As Object
        oStPoint.Set 0, 0, 0.5 * dWidthExclCylDiam + dCylinderDiam / 2
        oEnPoint.Set 0, 0, 0.55 * dWidthExclCylDiam + dCylinderDiam

        Set ObjEndCylinder3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsEnd Cylinder 3", ObjEndCylinder3
        Set ObjEndCylinder3 = Nothing

    'Creating the Output 12 (End Cylinder 4)
        Dim ObjEndCylinder4 As Object
        oStPoint.Set 0, 0, 0.55 * dWidthExclCylDiam + dCylinderDiam
        oEnPoint.Set 0, 0, 0.65 * dWidthExclCylDiam + dCylinderDiam + parInsulationThickness

        Set ObjEndCylinder4 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.2 * pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsEnd Cylinder 4", ObjEndCylinder4
        Set ObjEndCylinder4 = Nothing

    'Creating the Output 13 (Port 1 Cylinder)
        Dim ObjPort1Cylinder As Object
        oStPoint.Set 0, 0, 0
        oEnPoint.Set -dFacetoCenter, 0, 0

        Set ObjPort1Cylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsPort 1 Cylinder", ObjPort1Cylinder
        Set ObjPort1Cylinder = Nothing

    'Creating the Output 14 (Port 2 Cylinder)
        Dim ObjPort2Cylinder As Object
        oStPoint.Set 0, -(dCenter1toBottom - dCenter2toBottom), 0
        oEnPoint.Set -dFacetoCenter, -(dCenter1toBottom - dCenter2toBottom), 0

        Set ObjPort2Cylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsPort 2 Cylinder", ObjPort2Cylinder
        Set ObjPort2Cylinder = Nothing

    'Creating the Output 14 (Cylinder 2 near port 1)
        Dim ObjPort1Cylinder2 As Object
        oStPoint.Set -0.4 * dFacetoCenter, 0, 0
        oEnPoint.Set -0.6 * dFacetoCenter, 0, 0

        Set ObjPort1Cylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.2 * pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder 2 at Port 1", ObjPort1Cylinder2
        Set ObjPort1Cylinder = Nothing

    'Creating the Output 15 (Cylinder 2 near port 2)
        Dim ObjPort2Cylinder2 As Object
        oStPoint.Set -0.4 * dFacetoCenter, -(dCenter1toBottom - dCenter2toBottom), 0
        oEnPoint.Set -0.6 * dFacetoCenter, -(dCenter1toBottom - dCenter2toBottom), 0

        Set ObjPort2Cylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.2 * pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder 2 at Port 2", ObjPort2Cylinder2
        Set ObjPort2Cylinder2 = Nothing

    'Creating the Output 16 (Cylinder for Handle)
        Dim ObjHandleCylinder As Object
        oStPoint.Set 0, -(dCenter1toBottom - dCenter2toBottom), 0
        oEnPoint.Set 0, 0.7 * pipeDiam, 0

        Set ObjHandleCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * dWidthExclCylDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder for Handle", ObjHandleCylinder
        Set ObjHandleCylinder = Nothing

    'Creating the Output 17 (Box for Handle)
        Dim ObjHandleBox As Object
        oStPoint.Set 0.1 * pipeDiam + parInsulationThickness, 0.7 * pipeDiam, -0.25 * dWidthExclCylDiam - parInsulationThickness
        oEnPoint.Set -0.1 * pipeDiam - parInsulationThickness, 0.8 * pipeDiam + parInsulationThickness, 0.25 * dWidthExclCylDiam + parInsulationThickness

        Set ObjHandleBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    ' Set the output
        m_OutputColl.AddOutput "InsBox for Handle", ObjHandleBox
        Set ObjHandleBox = Nothing
        Set oTransformationMat = Nothing
        Set oLineString = Nothing
        Set oAxisVec = Nothing
        Set geomFactory = Nothing

    'Creating Insulation for port 1
        If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
            oStPoint.Set -dFacetoCenter, 0, 0
            oEnPoint.Set -(dFacetoCenter - flangeThick - parInsulationThickness), 0, 0
    
            Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam + 2 * parInsulationThickness, True)
        
            ' Set the output
            m_OutputColl.AddOutput "InsPort1", ObjInsulatedPort1
            Set ObjInsulatedPort1 = Nothing
        End If
        
        'Creating Insulation for port 2
        If CmpDblGreaterthan(flangeThick2, LINEAR_TOLERANCE) Then
            oStPoint.Set -dFacetoCenter, -(dCenter1toBottom - dCenter2toBottom), 0
            oEnPoint.Set -(dFacetoCenter - flangeThick2 - parInsulationThickness), -(dCenter1toBottom - dCenter2toBottom), 0
            
            Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam2 + 2 * parInsulationThickness, True)
            ' Set the output
            m_OutputColl.AddOutput "InsPort2", ObjInsulatedPort2
            Set ObjInsulatedPort2 = Nothing
        End If
        Set oStPoint = Nothing
        Set oEnPoint = Nothing

    Case 996, 997
    'Creating the Output 1(Cylinder 1)
        oStPoint.Set 0, -dCentertoBottom - parInsulationThickness, -(dOverallWidth - dCylinderDiam) / 2
        oEnPoint.Set 0, 0.2 * (dHeight - dCentertoBottom) + parInsulationThickness, -(dOverallWidth - dCylinderDiam) / 2

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylinderDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder 1", ObjCylinder
        Set ObjCylinder = Nothing

    'Creating the Output 2(Cylinder 2)
        oStPoint.Set 0, -dCentertoBottom - parInsulationThickness, (dOverallWidth - dCylinderDiam) / 2
        oEnPoint.Set 0, 0.2 * (dHeight - dCentertoBottom) + parInsulationThickness, (dOverallWidth - dCylinderDiam) / 2

        Set ObjCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylinderDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder 2", ObjCylinder2
        Set ObjCylinder2 = Nothing

    'Creating the Output 3(Box at Center)
        Dim ObjCenterBox1 As Object
        dBoxWidth = 0.6 * (dOverallWidth - 2 * dCylinderDiam)
        dBoxLength = 1.1 * SmallerDim(dFacetoFace - dFace2toCenter, dFace2toCenter)
        dInsBoxLength = dBoxLength + 2 * parInsulationThickness

        oStPoint.Set -dInsBoxLength / 2, -0.9 * dCentertoBottom - parInsulationThickness, -dBoxWidth / 2
        oEnPoint.Set dInsBoxLength / 2, 0.3 * (dHeight - dCentertoBottom) + parInsulationThickness, dBoxWidth / 2

        Set ObjCenterBox1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    ' Set the output
        m_OutputColl.AddOutput "InsCenter Box 1", ObjCenterBox1
        Set ObjCenterBox1 = Nothing

    'Creating the Output 5(Box 2 connecting cylinders)
        Dim ObjCenterBox2 As Object
        dBox2Length = 0.6 * dCylinderDiam
        dInsBox2Length = dBox2Length + 2 * parInsulationThickness
        dBox2Width = dOverallWidth - dCylinderDiam - 2 * Sqr((dCylinderDiam / 2) ^ 2 - (dBox2Length / 2) ^ 2)

        oStPoint.Set -dInsBox2Length / 2, -0.8 * dCentertoBottom - parInsulationThickness, -dBox2Width / 2
        oEnPoint.Set dInsBox2Length / 2, 0.1 * (dHeight - dCentertoBottom) + parInsulationThickness, dBox2Width / 2

        Set ObjCenterBox2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    ' Set the output
        m_OutputColl.AddOutput "InsCenter Box 2", ObjCenterBox2
        Set ObjCenterBox2 = Nothing
    
    'Creating the Output 11(Cylinder connecting port 1 and port 2)
        Dim ObjBodyCylinder As Object

        oStPoint.Set -(dFacetoFace - dFace2toCenter), 0, 0
        oEnPoint.Set dFace2toCenter, 0, 0

         Set ObjBodyCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCylinder
        Set ObjBodyCylinder = Nothing
        
    'Creating Insulation for port 1
        If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
            oStPoint.Set -(dFacetoFace - dFace2toCenter), 0, 0
            oEnPoint.Set -(dFacetoFace - dFace2toCenter) + flangeThick + parInsulationThickness, 0, 0
            
        'Check to see that insulation doesn't exceed face to face dimension
        If CmpDblGreaterthan(flangeThick + parInsulationThickness, dFacetoFace) Then
                oEnPoint.Set dFace2toCenter, 0, 0
            End If
            Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam + 2 * parInsulationThickness, True)
        
            ' Set the output
            m_OutputColl.AddOutput "InsPort1", ObjInsulatedPort1
            Set ObjInsulatedPort1 = Nothing
        End If
        
    'Creating Insulation for port 2
        If CmpDblGreaterthan(flangeThick2, LINEAR_TOLERANCE) Then
            oStPoint.Set dFace2toCenter, 0, 0
            oEnPoint.Set (dFace2toCenter - flangeThick2 - parInsulationThickness), 0, 0
            
            'Check to see that header insulation doesn't exceed face to face dimension
            If CmpDblGreaterthan(flangeThick2 + parInsulationThickness, dFacetoFace) Then
                oEnPoint.Set -(dFacetoFace - dFace2toCenter), 0, 0
            End If
            
            Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam2 + 2 * parInsulationThickness, True)
            ' Set the output
            m_OutputColl.AddOutput "InsPort2", ObjInsulatedPort2
            Set ObjInsulatedPort2 = Nothing
        End If
        Set oStPoint = Nothing
        Set oEnPoint = Nothing

    Case 998

    'Creating the Output 1(Box for body)
        oStPoint.Set -dLength / 2 - parInsulationThickness, -dCentertoBottom - parInsulationThickness, -dCentertoEnd2 - parInsulationThickness
        oEnPoint.Set dLength / 2 + parInsulationThickness, (dHeight - dCentertoBottom) + parInsulationThickness, 0.6 * dCentertoEnd1 + parInsulationThickness

        Set ObjBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    ' Set the output
        m_OutputColl.AddOutput "InsBody", ObjBody
        Set ObjBody = Nothing

    'Creating the Output 2(Cylinder connecting port 1 and port 2)

        oStPoint.Set -dFacetoFace / 2, 0, 0
        oEnPoint.Set dFacetoFace / 2, 0, 0

         Set ObjBodyCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCylinder
        Set ObjBodyCylinder = Nothing

    'Creating the Output 3(Cylinder for handwheel 1)
        Dim dHandWheelCylDiam As Double
        dHandWheelCylDiam = 0.1 * dLength
        Dim objHandWheelCyl1 As Object

        oStPoint.Set 0.45 * dLength, 0, 0.6 * dCentertoEnd1
        oEnPoint.Set 0.45 * dLength, 0, 0.9 * dCentertoEnd1

        Set objHandWheelCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dHandWheelCylDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder for HandWheel 1", objHandWheelCyl1
        Set objHandWheelCyl1 = Nothing

    'Creating the Output 4(Cylinder for handwheel 2)
        Dim objHandWheelCyl2 As Object

        oStPoint.Set -0.45 * dLength, 0, 0.6 * dCentertoEnd1
        oEnPoint.Set -0.45 * dLength, 0, 0.9 * dCentertoEnd1

        Set objHandWheelCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dHandWheelCylDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder for HandWheel 2", objHandWheelCyl2
        Set objHandWheelCyl2 = Nothing

  'Creating Insulation for port 1
        If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
            oStPoint.Set -dFacetoFace / 2, 0, 0
            oEnPoint.Set -(dFacetoFace / 2 - flangeThick - parInsulationThickness), 0, 0
            
        'Check to see that insulation doesn't exceed face to face dimension
            If CmpDblGreaterthan(flangeThick + parInsulationThickness, dFacetoFace) Then
                oEnPoint.Set dFacetoFace / 2, 0, 0
            End If
            Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam + 2 * parInsulationThickness, True)
        
            ' Set the output
            m_OutputColl.AddOutput "InsPort1", ObjInsulatedPort1
            Set ObjInsulatedPort1 = Nothing
        End If
        
        'Creating Insulation for port 2
        If CmpDblGreaterthan(flangeThick2, LINEAR_TOLERANCE) Then
    
            oStPoint.Set dFacetoFace / 2, 0, 0
            oEnPoint.Set (dFacetoFace / 2 - flangeThick2 - parInsulationThickness), 0, 0
            
            'Check to see that header insulation doesn't exceed face to face dimension
            If CmpDblGreaterthan(flangeThick2 + parInsulationThickness, dFacetoFace) Then
                oEnPoint.Set -dFacetoFace / 2, 0, 0
            End If
            
            Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam2 + 2 * parInsulationThickness, True)
            ' Set the output
            m_OutputColl.AddOutput "InsPort2", ObjInsulatedPort2
            Set ObjInsulatedPort2 = Nothing
        End If
        
        Set oStPoint = Nothing
        Set oEnPoint = Nothing

    Case 999
    'Creating the Output 1(Cylinder for Basket)
        Dim objBasketCylinder As Object

        oStPoint.Set 0, -dCentertoBottom - parInsulationThickness, 0.8 * dCentertoCylEnd
        oEnPoint.Set 0, dHeight - dCentertoBottom + parInsulationThickness, 0.8 * dCentertoCylEnd

        dCylinderDiam = 0.4 * dCentertoCylEnd
        dInsCylinderDiam = 0.4 * dCentertoCylEnd + 2 * parInsulationThickness
        Set objBasketCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylinderDiam, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder for Basket", objBasketCylinder
        Set objBasketCylinder = Nothing

    'Creating the Output 2(Cylinder 2 for Basket)
        Dim objBasketCylinder2 As Object
        Dim dCylinderDiam2  As Double
        Dim dInsCylinderDiam2 As Double
        dCylinderDiam2 = 2 * (dWidth - 1.8 * dCentertoCylEnd)
        dInsCylinderDiam2 = dCylinderDiam2 + 2 * parInsulationThickness

        oStPoint.Set 0, -dCentertoBottom - parInsulationThickness, -0.8 * (dCentertoCylEnd)
        oEnPoint.Set 0, dHeight - dCentertoBottom + parInsulationThickness, -0.8 * (dCentertoCylEnd)

        Set objBasketCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylinderDiam2, True)

    ' Set the output
        m_OutputColl.AddOutput "InsCylinder for Basket 2", objBasketCylinder2
        Set objBasketCylinder2 = Nothing

    'Creating the Output 3(Connecting Cylinder 1)
        Dim objConnCylinder1 As Object
        Dim dConnCylinderDiam As Double
        Dim dCylRadialPos As Double
        Dim dAngle As Double
        dConnCylinderDiam = 0.4 * dCylinderDiam
        Dim dInsConnCylinderDiam As Double
        dInsConnCylinderDiam = dConnCylinderDiam + 2 * parInsulationThickness
        
        dCylRadialPos = 0.5 * Sqr(dCylinderDiam ^ 2 - dConnCylinderDiam ^ 2)
        dAngle = PI / 4

        oStPoint.Set 0, 0, 0.8 * dCentertoCylEnd
        oEnPoint.Set -1.5 * dCylRadialPos * Sin(dAngle), 0, 0.8 * dCentertoCylEnd - 1.5 * dCylRadialPos * Cos(dAngle)

        Set objConnCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsConnCylinderDiam, True)

    ' Set the output
        m_OutputColl.AddOutput "InsConnecting Cylinder 1", objConnCylinder1
        Set objConnCylinder1 = Nothing

    'Creating the Output 4(Connecting Cylinder 2)
        Dim objConnCylinder2 As Object

        oStPoint.Set 0, 0, 0.8 * dCentertoCylEnd
        oEnPoint.Set 1.5 * dCylRadialPos * Sin(dAngle), 0, 0.8 * dCentertoCylEnd - 1.5 * dCylRadialPos * Cos(dAngle)

        Set objConnCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsConnCylinderDiam, True)

    ' Set the output
        m_OutputColl.AddOutput "InsConnecting Cylinder 2", objConnCylinder2
        Set objConnCylinder2 = Nothing

    'Creating the Output 5(Connecting Cylinder 3)
        Dim objConnCylinder3 As Object
        Dim dConnCylinderDiam2 As Double
        Dim dCylRadialPos2 As Double
        Dim dInsConnCylinderDiam2 As Double
        dConnCylinderDiam2 = 0.4 * dCylinderDiam2
        dInsConnCylinderDiam2 = dConnCylinderDiam2 + 2 * parInsulationThickness
        dCylRadialPos2 = dCylRadialPos

        oStPoint.Set 0, 0, -0.8 * (dCentertoCylEnd)
        oEnPoint.Set -1.5 * dCylRadialPos2 * Sin(dAngle), 0, -0.8 * (dCentertoCylEnd) + 1.5 * dCylRadialPos2 * Cos(dAngle)

        Set objConnCylinder3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsConnCylinderDiam2, True)

    ' Set the output
        m_OutputColl.AddOutput "InsConnecting Cylinder 3", objConnCylinder3
        Set objConnCylinder3 = Nothing

    'Creating the Output 6(Connecting Cylinder 4)
        Dim objConnCylinder4 As Object

        oStPoint.Set 0, 0, -0.8 * (dCentertoCylEnd)
        oEnPoint.Set 1.5 * dCylRadialPos2 * Sin(dAngle), 0, -0.8 * (dCentertoCylEnd) + 1.5 * dCylRadialPos2 * Cos(dAngle)

        Set objConnCylinder4 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsConnCylinderDiam2, True)

    ' Set the output
        m_OutputColl.AddOutput "InsConnecting Cylinder 4", objConnCylinder4
        Set objConnCylinder4 = Nothing

    'Creating the Output 7(Flange Cylinder 1)
        Dim objFlangeCylinder1 As Object
        Dim dFlangeCylinderDiam As Double
        Dim dInsFlangeCylinderDiam As Double
        Dim dFlangeCylRadialPos As Double

        dFlangeCylinderDiam = 0.8 * flangeDiam
        dInsFlangeCylinderDiam = dFlangeCylinderDiam + 2 * parInsulationThickness
        dFlangeCylRadialPos = 1.5 * dCylRadialPos

        oStPoint.Set -dFlangeCylRadialPos * Sin(dAngle), 0, 0.8 * dCentertoCylEnd - dFlangeCylRadialPos * Cos(dAngle)
        oEnPoint.Set -1.5 * dFlangeCylRadialPos * Sin(dAngle), 0, 0.8 * dCentertoCylEnd - 1.5 * dFlangeCylRadialPos * Cos(dAngle)

        Set objFlangeCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsFlangeCylinderDiam, True)

    ' Set the output
        m_OutputColl.AddOutput "InsFlange Cylinder 1", objFlangeCylinder1
        Set objFlangeCylinder1 = Nothing

    'Creating the Output 8(Flange Cylinder 2)
        Dim objFlangeCylinder2 As Object

        oStPoint.Set dFlangeCylRadialPos * Sin(dAngle), 0, 0.8 * dCentertoCylEnd - dFlangeCylRadialPos * Cos(dAngle)
        oEnPoint.Set 1.5 * dFlangeCylRadialPos * Sin(dAngle), 0, 0.8 * dCentertoCylEnd - 1.5 * dFlangeCylRadialPos * Cos(dAngle)

        Set objFlangeCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsFlangeCylinderDiam, True)

    ' Set the output
        m_OutputColl.AddOutput "InsFlange Cylinder 2", objFlangeCylinder2
        Set objFlangeCylinder2 = Nothing

    'Creating the Output 9(Flange Cylinder 3)
        Dim objFlangeCylinder3 As Object
        Dim dFlangeCylinderDiam2  As Double
        Dim dInsFlangeCylinderDiam2  As Double
        Dim dFlangeCylRadialPos2 As Double

        dFlangeCylinderDiam2 = 0.8 * flangeDiam
        dInsFlangeCylinderDiam2 = dFlangeCylinderDiam2 + 2 * parInsulationThickness
        dFlangeCylRadialPos2 = 1.5 * dCylRadialPos2
        oStPoint.Set -dFlangeCylRadialPos2 * Sin(dAngle), 0, -0.8 * (dCentertoCylEnd) + dFlangeCylRadialPos2 * Cos(dAngle)
        oEnPoint.Set -1.5 * dFlangeCylRadialPos2 * Sin(dAngle), 0, -0.8 * (dCentertoCylEnd) + 1.5 * dFlangeCylRadialPos2 * Cos(dAngle)

        Set objFlangeCylinder3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsFlangeCylinderDiam2, True)

    ' Set the output
        m_OutputColl.AddOutput "InsFlange Cylinder 3", objFlangeCylinder3
        Set objFlangeCylinder3 = Nothing

    'Creating the Output 10(Flange Cylinder 4)
        Dim objFlangeCylinder4 As Object

        oStPoint.Set dFlangeCylRadialPos2 * Sin(dAngle), 0, -0.8 * (dCentertoCylEnd) + dFlangeCylRadialPos2 * Cos(dAngle)
        oEnPoint.Set 1.5 * dFlangeCylRadialPos2 * Sin(dAngle), 0, -0.8 * (dCentertoCylEnd) + 1.5 * dFlangeCylRadialPos2 * Cos(dAngle)

        Set objFlangeCylinder4 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsFlangeCylinderDiam2, True)

    ' Set the output
        m_OutputColl.AddOutput "InsFlange Cylinder 4", objFlangeCylinder4
        Set objFlangeCylinder4 = Nothing

    'Creating the Output 11(Curved portion of the strainer)
        Dim objCircle1   As IngrGeom3D.Circle3d
        Set geomFactory = New IngrGeom3D.GeometryFactory
        
        'Setting the start and end points of the centers of elbow curve
        oStPoint.Set -1.5 * dFlangeCylRadialPos * Sin(dAngle), _
                      0, 0.8 * dCentertoCylEnd - 1.5 * dFlangeCylRadialPos * Cos(dAngle)
        oEnPoint.Set -1.5 * dFlangeCylRadialPos2 * Sin(dAngle), _
                      0, -0.8 * (dCentertoCylEnd) + 1.5 * dFlangeCylRadialPos2 * Cos(dAngle)
                      
        Dim dInsCircleRadius As Double
        dInsCircleRadius = 0.45 * dFlangeCylinderDiam + parInsulationThickness
        
        If CmpDblGreaterthan(dInsCircleRadius, oStPoint.z / Cos(dAngle)) Then
            dInsCircleRadius = oStPoint.z / Cos(dAngle)
        End If
        
        'Creating circle 1
        Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                             -1.5 * dFlangeCylRadialPos * Sin(dAngle), 0, 0.8 * dCentertoCylEnd - 1.5 * dFlangeCylRadialPos * Cos(dAngle), _
                            Sin(dAngle), 0, Cos(dAngle), _
                            dInsCircleRadius)

        Dim CP As New AutoMath.DPosition 'revolution center point
        Dim CV As New AutoMath.DVector 'rotation vector for rotation
        CP.Set oStPoint.x + (oStPoint.z / Tan(dAngle)), 0, 0
        CV.Set 0, -1, 0

        'Creating the revolution
        Dim ObjElbow As Object
        Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle1, _
                            CV, CP, PI / 2, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
        Set ObjElbow = Nothing

    'Creating the Output 12(Curved portion 2 of the strainer)
        'Creating the copy of the above elbow surface to manipulate it using transformations in order to get the other elbow surface
        Dim ObjElbow2 As Object
        Set ObjElbow2 = PlaceRevolution(m_OutputColl, objCircle1, _
                            CV, CP, PI / 2, True)

        Set objCircle1 = Nothing

        Set oTransformationMat = New AutoMath.DT4x4

        'Creating the mirror of the first elbow surface
        oTransformationMat.LoadIdentity
        oTransformationMat.IndexValue(0) = -1
        ObjElbow2.Transform oTransformationMat

        'Translating the mirrored surface to get the correct location
        Set oDirVector = New AutoMath.DVector
        oTransformationMat.LoadIdentity
        oDirVector.Set 1.5 * (dFlangeCylRadialPos2 - dFlangeCylRadialPos) * Sin(dAngle), 0, 0
        oTransformationMat.Translate oDirVector
        ObjElbow2.Transform oTransformationMat

        'Set the output
        m_OutputColl.AddOutput "InsCurved Elbow 2", ObjElbow2
        Set ObjElbow2 = Nothing
        Set oTransformationMat = Nothing
        Set oDirVector = Nothing
    
        'Creating the Output 10(Insulated cylinder for port 1)
        Dim objInsPort1Cylinder As Object
        Dim dNozzleLength As Double
        dNozzleLength = dFace1toCenter + CP.x - ((0.8 * dCentertoCylEnd - 1.5 * dFlangeCylRadialPos * Cos(dAngle)) / Sin(dAngle))

        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set -dFace1toCenter + dNozzleLength, 0, 0

        Set objInsPort1Cylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "Insulated cylinder for port 1", objInsPort1Cylinder
        Set objInsPort1Cylinder = Nothing
    
    'Creating the Output 11(Insulated cylinder for port 2)
        Dim objInsPort2Cylinder As Object
        dNozzleLength = dFacetoFace - dFace1toCenter + CP.x - ((0.8 * dCentertoCylEnd - 1.5 * dFlangeCylRadialPos * Cos(dAngle)) / Sin(dAngle))

        oStPoint.Set dFacetoFace - dFace1toCenter, 0, 0
        oEnPoint.Set dFacetoFace - dFace1toCenter - dNozzleLength, 0, 0

        Set objInsPort2Cylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam2 + 2 * parInsulationThickness, True)

    ' Set the output
        m_OutputColl.AddOutput "Insulation for port 2", objInsPort2Cylinder
        Set objInsPort2Cylinder = Nothing
        
    'Creating Insulation for port 1
        If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + flangeThick + parInsulationThickness, 0, 0
            
        'Check to see that insulation doesn't exceed face to face dimension
            If CmpDblGreaterthan(flangeThick + parInsulationThickness, dFacetoFace) Then
                oEnPoint.Set dFacetoFace - dFace1toCenter, 0, 0
            End If
            Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam + 2 * parInsulationThickness, True)
        
            ' Set the output
            m_OutputColl.AddOutput "InsPort1", ObjInsulatedPort1
            Set ObjInsulatedPort1 = Nothing
        End If
        
    'Creating Insulation for port 2
        If CmpDblGreaterthan(flangeThick2, LINEAR_TOLERANCE) Then
            oStPoint.Set dFacetoFace - dFace1toCenter, 0, 0
            oEnPoint.Set (dFacetoFace - dFace1toCenter - flangeThick2 - parInsulationThickness), 0, 0
            
            'Check to see that header insulation doesn't exceed face to face dimension
            If CmpDblGreaterthan(flangeThick2 + parInsulationThickness, dFacetoFace) Then
                oEnPoint.Set -dFace1toCenter, 0, 0
            End If
            
            Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam2 + 2 * parInsulationThickness, True)
            ' Set the output
            m_OutputColl.AddOutput "InsPort2", ObjInsulatedPort2
            Set ObjInsulatedPort2 = Nothing
        End If
        
        Set CP = Nothing
        Set CV = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set geomFactory = Nothing
    Case 1000
    'Creating the Output 1(Box 1 for body)
        'Assumption for the overall width
        dOverallWidth = 1.1 * dWidth

        'Check on the overall width
        If CmpDblLessThan(dOverallWidth, parZOffset + flangeDiam3) Then
            dOverallWidth = 1.2 * (parZOffset + flangeDiam3)
        End If

        'Assumption for the center to top dimension
        dCentertoTop = 2.5 * pipeDiam

        Dim dBoxEnd1X As Double
        Dim dBoxEnd2X As Double
        dBoxEnd1X = 0.5 * (dFace1toCenter - flangeThick)
        dBoxEnd2X = 0.5 * (dFace2toCenter - flangeThick)

        oStPoint.Set -dBoxEnd1X - parInsulationThickness, -dCentertoBottom - parInsulationThickness, _
                                                    dOverallWidth / 2 + parInsulationThickness
        oEnPoint.Set dBoxEnd2X + parInsulationThickness, 0.5 * dCentertoTop + parInsulationThickness, _
                                                    dOverallWidth / 4

        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    ' Set the output
        m_OutputColl.AddOutput "InsBox 1", ObjBox
        Set ObjBox = Nothing

    'Creating the Output 2(Box 2 for body)

        oStPoint.Set -dBoxEnd1X - parInsulationThickness, -dCentertoBottom - parInsulationThickness, _
                                                -dOverallWidth / 2 - parInsulationThickness
        oEnPoint.Set dBoxEnd2X + parInsulationThickness, 0.5 * dCentertoTop + parInsulationThickness, -dOverallWidth / 4

        Set ObjBox2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    ' Set the output
        m_OutputColl.AddOutput "InsBox 2", ObjBox2
        Set ObjBox2 = Nothing

    'Creating the Output 3(Box 3 for body)
        Dim ObjBox3 As Object
        oStPoint.Set -dBoxEnd1X - parInsulationThickness, -0.75 * dCentertoBottom - parInsulationThickness, dOverallWidth / 4
        oEnPoint.Set dBoxEnd2X + parInsulationThickness, 0.25 * dCentertoTop + parInsulationThickness, -dOverallWidth / 4

        Set ObjBox3 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    ' Set the output
        m_OutputColl.AddOutput "InsBox 3", ObjBox3
        Set ObjBox3 = Nothing

    'Creating the Output 4(Cylinder connecting port 1 and port 2)
        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set dFace2toCenter, 0, 0

         Set ObjBodyCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * parInsulationThickness, True)

    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCylinder
        Set ObjBodyCylinder = Nothing
    
    'Creating Insulation for port 1
        If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + flangeThick + parInsulationThickness, 0, 0
            
        'Check to see that insulation doesn't exceed face to face dimension
            If CmpDblGreaterthan(flangeThick + parInsulationThickness, dFace1toCenter + dFace2toCenter) Then
                oEnPoint.Set dFace2toCenter, 0, 0
            End If
            Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam + 2 * parInsulationThickness, True)
        
            ' Set the output
            m_OutputColl.AddOutput "InsPort1", ObjInsulatedPort1
            Set ObjInsulatedPort1 = Nothing
        End If
        
    'Creating Insulation for port 2
        If CmpDblGreaterthan(flangeThick2, LINEAR_TOLERANCE) Then
            oStPoint.Set dFace2toCenter, 0, 0
            oEnPoint.Set (dFace2toCenter - flangeThick2 - parInsulationThickness), 0, 0
            
            'Check to see that header insulation doesn't exceed face to face dimension
            If CmpDblGreaterthan(flangeThick2 + parInsulationThickness, dFace1toCenter + dFace2toCenter) Then
                oEnPoint.Set -dFace1toCenter, 0, 0
            End If
            
            Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam2 + 2 * parInsulationThickness, True)
            ' Set the output
            m_OutputColl.AddOutput "InsPort2", ObjInsulatedPort2
            Set ObjInsulatedPort2 = Nothing
        End If
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
    End Select
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub



