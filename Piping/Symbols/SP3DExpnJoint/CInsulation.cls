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
'   Author:         KKC
'   Creation Date:  Tuesday, Feb 14 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the Style 443-B and 443-BS in the 440-Series-6page.pdf
'    Proco's catalog available at URL http://www.bigriverrubber.com/proco/expjointlocation.htm
'    For the part data basis 935- 944 - Source - www.ejsus.com
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

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim sptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double
    Dim sptOffset1 As Double
    Dim depth1 As Double
    Dim flangeThick1 As Double
    Dim flangeThick2 As Double
    Dim flangeDiam1 As Double
    Dim flangeDiam2 As Double

    Dim iOutput As Double

    Dim parFacetoFace As Double
    Dim parPipingSpecialtyDiameter As Double
    Dim parCenterSpoollength As Double
    Dim parAngleFlangeWidth As Double
    Dim parOuterCasingLength As Double
    Dim parPurgePortLenoffset As Double
    Dim parPurgePortHeioffset As Double
    Dim parDrainPortLenoffset As Double
    Dim parDrainPortHeioffset As Double
    Dim parHingeWidth As Double
    Dim parGimbalWidth As Double
    Dim parFlangeWidth As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double

    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oCP As AutoMath.DPosition
    Set oCP = New AutoMath.DPosition
    Dim ocv As AutoMath.DVector
    Set ocv = New AutoMath.DVector
    Dim oTransMatrix As IJDT4x4
    Set oTransMatrix = New AutoMath.DT4x4
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector


    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Select Case lPartDataBasis
    Case Is <= 1
        parFacetoFace = arrayOfInputs(2)    'L
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parInsulationThickness = arrayOfInputs(16)
    Case 935
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parInsulationThickness = arrayOfInputs(16)
    Case 936
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parInsulationThickness = arrayOfInputs(16)
    Case 936
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parCenterSpoollength = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(16)
    Case 937
        parFacetoFace = arrayOfInputs(2)
        parOuterCasingLength = arrayOfInputs(5)
        parPurgePortLenoffset = arrayOfInputs(6)
        parPurgePortHeioffset = arrayOfInputs(7)
        parDrainPortLenoffset = arrayOfInputs(8)
        parDrainPortHeioffset = arrayOfInputs(9)
        parInsulationThickness = arrayOfInputs(16)
    Case 938
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parAngleFlangeWidth = arrayOfInputs(10)
        parInsulationThickness = arrayOfInputs(16)

    Case 939
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parCenterSpoollength = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(16)

    Case 940
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parHingeWidth = arrayOfInputs(11)
        parInsulationThickness = arrayOfInputs(16)

    Case 941
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parGimbalWidth = arrayOfInputs(12)
        parInsulationThickness = arrayOfInputs(16)

    Case 942
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parFace1toCenter = arrayOfInputs(13)
        parFace2toCenter = arrayOfInputs(14)
        parFlangeWidth = arrayOfInputs(15)
        parInsulationThickness = arrayOfInputs(16)

    Case 943
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parInsulationThickness = arrayOfInputs(16)

    Case 944
        parFacetoFace = arrayOfInputs(2)
        parOuterCasingLength = arrayOfInputs(5)
        parPurgePortLenoffset = arrayOfInputs(6)
        parPurgePortHeioffset = arrayOfInputs(7)
        parDrainPortLenoffset = arrayOfInputs(8)
        parDrainPortHeioffset = arrayOfInputs(9)
        parInsulationThickness = arrayOfInputs(16)
    End Select
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick1, flangeDiam1, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam2, sptOffset1, depth1
  
   If lPartDataBasis <= 1 Then
        iOutput = 0
        '   Assumption: Origin of the symbol is taken at mid-point of parFacetoFace

        '   Insert your code for output 1(Left Hand Side Flange)
        '   Assumptions:
        '    1. The Body of 'Limit bolt Outside Ear' is taken to be cylinder, and
        '    2. Its thickness is taken to be equal to that of the flange thickness
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1

        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set oStPoint.x + flangeThick + parInsulationThickness, 0, 0
        Dim ObjInsFlCylinder As Object
        Set ObjInsFlCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parPipingSpecialtyDiameter + parInsulationThickness, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjInsFlCylinder
        Set ObjInsFlCylinder = Nothing

        '   Insert your code for output 2(Body of Expansion Joint)
        '    Assumption: The Bellows outside diameter is taken to be equal to 1.2 time the pipe diameter
        Dim dBellowOD As Double

        dBellowOD = pipeDiam * 1.2
        oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oEnPoint.Set parFacetoFace / 2 - flangeThick, 0, 0

        Dim ObjBodyCylinder As Object
        Set ObjBodyCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBellowOD + 2 * parInsulationThickness, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjIns_", ObjBodyCylinder
        Set ObjBodyCylinder = Nothing

        '   Insert your code for output 3(Right Hand Side Flange)
        oStPoint.Set parFacetoFace / 2, 0, 0
        oEnPoint.Set oStPoint.x - flangeThick, 0, 0

        Set ObjInsFlCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parPipingSpecialtyDiameter + 2 * parInsulationThickness, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjInsFlCylinder
        Set ObjInsFlCylinder = Nothing

   ElseIf lPartDataBasis = 937 Or lPartDataBasis = 944 Then
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0

        Dim ObjCylinder As Object
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 2 * pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing
        If Not CmpDblEqual(flangeThick1, 0) Then
            Dim dPortLen As Double
            Dim dInsDia As Double
            dPortLen = flangeThick1 + parInsulationThickness

            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set -parFacetoFace / 2 + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Dim objInsPort As Object
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

        If Not CmpDblEqual(flangeThick2, 0) Then
            dPortLen = flangeThick2 + parInsulationThickness
            oStPoint.Set parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2 - dPortLen, 0, 0
            dInsDia = flangeDiam2 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
    ElseIf lPartDataBasis = 935 Or lPartDataBasis = 936 Or lPartDataBasis = 938 Or lPartDataBasis = 939 Or lPartDataBasis = 940 Or lPartDataBasis = 941 Or lPartDataBasis = 943 Then
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parPipingSpecialtyDiameter + 2 * parInsulationThickness, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        If Not CmpDblEqual(flangeThick1, 0) Then


            dPortLen = flangeThick1 + parInsulationThickness

            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set -parFacetoFace / 2 + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness

            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        If Not CmpDblEqual(flangeThick2, 0) Then
            dPortLen = flangeThick2 + parInsulationThickness
            oStPoint.Set parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2 - dPortLen, 0, 0
            dInsDia = flangeDiam2 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
    ElseIf lPartDataBasis = 942 Then
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        parFacetoFace = 2 * parFace1toCenter
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parPipingSpecialtyDiameter + 2 * parInsulationThickness, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing
        ' Insert your code for the Outlet Curve
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
        Dim dStartingPoint As Double
        Dim objCircle As Object
        dStartingPoint = parFace2toCenter

        Set objCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, pipeDiam / 2 + parInsulationThickness)

        Set oGeomFactory = Nothing
        '   Revolve the circle about the Z axis passing by the origin

        Set ocv = New AutoMath.DVector

        oCP.Set -parFace2toCenter, -parFace2toCenter, 0
        ocv.Set 0, 0, -1
        Dim parAngle As Double
        parAngle = PI / 2
        Dim objCurvedSurface As Object
        Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                               oCP, parAngle, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", objCurvedSurface
        Set objCurvedSurface = Nothing
        If Not CmpDblEqual(flangeThick1, 0) Then
            dPortLen = flangeThick1 + parInsulationThickness
            oStPoint.Set -parFace1toCenter, 0, 0
            oEnPoint.Set -parFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        If Not CmpDblEqual(flangeThick2, 0) Then
            dPortLen = flangeThick2 + parInsulationThickness
            oStPoint.Set 0, -parFace2toCenter, 0
            oEnPoint.Set 0, -parFace2toCenter + dPortLen, 0
            dInsDia = flangeDiam2 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
    End If
    Set oCP = Nothing
    Set ocv = Nothing
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
              Err.HelpFile, Err.HelpContext

End Sub


