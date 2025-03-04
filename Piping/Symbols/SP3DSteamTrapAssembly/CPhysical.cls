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
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DSteamTrapAssembly.SteamTA
'   Author:         MP
'   Creation Date:  Wednesday, Oct 15 2008
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'    15.Oct.2008    MP     CR-151135  Provide steam trap fitting unit symbols per Yarway catalog
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
    Dim PipeDiam1 As Double
    Dim PipeDiam2 As Double
    Dim PipeDiam3 As Double
    Dim FlangeThick1 As Double
    Dim FlangeThick2 As Double
    Dim FlangeThick3 As Double
    Dim CptOffset1 As Double
    Dim CptOffset2 As Double
    Dim CptOffset3 As Double
    Dim FlangeDiam1 As Double
    Dim FlangeDiam2 As Double
    Dim FlangeDiam3 As Double
    Dim Depth1 As Double
    Dim Depth2 As Double
    Dim Depth3 As Double
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnpoint As AutoMath.DPosition
    Dim oDirvec As AutoMath.DVector
    
    Set oStPoint = New DPosition
    Set oEnpoint = New DPosition
    Set oDirvec = New AutoMath.DVector

    Dim lPartDataBasis As Long
    Dim oPipeComponent As IJDPipeComponent

    Dim iOutput As Long
    
    Dim ObjCylinder1 As Object
    Dim objValveCyl1 As Object
    Dim objValveSphere As Object
    Dim objValveCyl2 As Object
    Dim objOperatorStem As Object
    Dim objOperatorHead As Object
    Dim ObjBody2 As Object
    Dim objNozzle1 As Object
    Dim objNozzle2 As Object
    Dim objNozzle3 As Object
    Dim objOp2Cyl1 As Object
    Dim objOp2Sphere As Object
    Dim objOp2Cyl2 As Object
    Dim objOp2Stem As Object
    Dim objOp2Head As Object
    Dim objOp3Cyl1 As Object
    Dim objOp3Sphere As Object
    Dim objOp3Cyl2 As Object
    Dim objOp3Stem As Object
    Dim objOp3Head As Object

    Dim parL As Double                      'FacetoFace
    Dim parH As Double                      'OperatorHeight
    Dim parB As Double                      'TrapLength1
    Dim parW As Double                      'OperatorDiameter
    Dim parInsulationThickness As Double    'InsulationThickness
    Dim parK As Double                      'TrapLength2

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parL = arrayOfInputs(2)    'FacetoFace
    parH = arrayOfInputs(3)    'OperatorHeight
    parB = arrayOfInputs(4)    'TrapLength1
    parW = arrayOfInputs(5)    'OperatorDiameter
    parInsulationThickness = arrayOfInputs(6)

    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    If lPartDataBasis = STP_Cock_Bypass Then
        parK = arrayOfInputs(7)    'Traplength2
    End If

    iOutput = 0

    If lPartDataBasis <= 1 Or lPartDataBasis = STA_DEFAULT Or _
                    lPartDataBasis = STA_Cock Or lPartDataBasis = STP_Cock_Bypass Then

        'Insert code for output1 (cylinder1)
        oStPoint.Set -parL / 2, 0, 0
        oEnpoint.Set parL / 2, 0, 0

        RetrieveParameters 1, oPartFclt, m_OutputColl, PipeDiam1, FlangeThick1, FlangeDiam1, CptOffset1, Depth1
        Dim dBody2Dia As Double
        dBody2Dia = 0.8 * PipeDiam1

        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, PipeDiam1, True)

        'Set Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder1
        Set ObjCylinder1 = Nothing

        'Insert code for output2 (valvecyl1)
        Dim dValveDia As Double
        dValveDia = 1.25 * PipeDiam1

        oStPoint.Set (parL / 4) - (parW * 0.4), 0, 0
        oEnpoint.Set (parL / 4) - (0.25 * parW), 0, 0

        Set objValveCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dValveDia, True)

        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveCyl1
        Set objValveCyl1 = Nothing

        'insert code for output3(ValveSpehere)
        Dim SphereRad As Double
        SphereRad = dValveDia / 2
        oStPoint.Set parL / 4, 0, 0

        Set objValveSphere = PlaceSphere(m_OutputColl, oStPoint, SphereRad)

        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveSphere
        Set objValveSphere = Nothing

        'insert code for output4 (ValveCyl2)
        oStPoint.Set (parL / 4) + (0.25 * parW), 0, 0
        oEnpoint.Set (parL / 4) + (0.4 * parW), 0, 0

        Set objValveCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dValveDia, True)

        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveCyl2
        Set objValveCyl2 = Nothing

        'insert code for output5 (operator stem)
        Dim StemDia As Double
        StemDia = 0.1 * parW

        oStPoint.Set parL / 4, 0, 0
        oEnpoint.Set parL / 4, 0.9 * parH, 0

        Set objOperatorStem = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, StemDia, True)

        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOperatorStem
        Set objOperatorStem = Nothing

        'insert code for output6
        oStPoint.Set parL / 4, 0.9 * parH, 0
        oEnpoint.Set parL / 4, parH, 0

        Set objOperatorHead = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, parW, True)

        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOperatorHead
        Set objOperatorHead = Nothing

        If lPartDataBasis = STA_Cock Or lPartDataBasis = STP_Cock_Bypass Then
            'insert code for op2-cyl1
            oStPoint.Set -(parL / 4) + (parW * 0.4), 0, 0
            oEnpoint.Set -(parL / 4) + (0.25 * parW), 0, 0

            Set objOp2Cyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dValveDia, True)

            'set the output
            m_OutputColl.AddOutput "Op2Cyl1", objOp2Cyl1
            Set objOp2Cyl1 = Nothing

            'insert code for op2-ValveSpehere
            oStPoint.Set -parL / 4, 0, 0

            Set objOp2Sphere = PlaceSphere(m_OutputColl, oStPoint, SphereRad)

            'set the output
            m_OutputColl.AddOutput "Op2Sphere", objOp2Sphere
            Set objValveSphere = Nothing
            Set objOp2Sphere = Nothing

            'insert code for Op2cyl2
            oStPoint.Set -(parL / 4) - (0.25 * parW), 0, 0
            oEnpoint.Set -(parL / 4) - (0.4 * parW), 0, 0

            Set objOp2Cyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dValveDia, True)

            'set the output
            m_OutputColl.AddOutput "Op2Cyl2", objOp2Cyl2
            Set objOp2Cyl2 = Nothing

            'insert code for Op2Stem
            oStPoint.Set -parL / 4, 0, 0
            oEnpoint.Set -parL / 4, 0.9 * parH, 0

            Set objOp2Stem = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, StemDia, True)

            'set the output
            m_OutputColl.AddOutput "Op2Stem", objOp2Stem
            Set objOp2Stem = Nothing

            'insert code for Op2head
            oStPoint.Set -parL / 4, 0.9 * parH, 0
            oEnpoint.Set -parL / 4, parH, 0

            Set objOp2Head = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, parW, True)

            'set the output
            m_OutputColl.AddOutput "Op2Head", objOp2Head
            Set objOp2Head = Nothing
            
            'insert code for op3-cyl1
            oStPoint.Set -0.25 * parW * Cos(PI / 4), -0.5 * parB + 0.25 * parW * Sin(PI / 4), 0
            oEnpoint.Set -0.4 * parW * Cos(PI / 4), -0.5 * parB + 0.4 * parW * Sin(PI / 4), 0

            Set objOp3Cyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dValveDia, True)

            'set the output
            m_OutputColl.AddOutput "Op3Cyl1", objOp3Cyl1
            Set objOp3Cyl1 = Nothing

            'insert code for op3-ValveSpehere
            oStPoint.Set 0, -0.5 * parB, 0

            Set objOp3Sphere = PlaceSphere(m_OutputColl, oStPoint, SphereRad)

            'set the output
            m_OutputColl.AddOutput "Op3Sphere", objOp3Sphere
            Set objValveSphere = Nothing
            Set objOp3Sphere = Nothing

            'insert code for Op3cyl2
            oStPoint.Set 0.25 * parW * Cos(PI / 4), -0.5 * parB - 0.25 * parW * Sin(PI / 4), 0
            oEnpoint.Set 0.4 * parW * Cos(PI / 4), -0.5 * parB - 0.4 * parW * Sin(PI / 4), 0

            Set objOp3Cyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dValveDia, True)

            'set the output
            m_OutputColl.AddOutput "Op3Cyl2", objOp3Cyl2
            Set objOp3Cyl2 = Nothing

            'insert code for Op3Stem
            oStPoint.Set 0, -0.5 * parB, 0
            oEnpoint.Set -0.5 * parB, -parB, 0

            Set objOp3Stem = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, StemDia, True)

            'set the output
            m_OutputColl.AddOutput "Op3Stem", objOp3Stem
            Set objOp3Stem = Nothing

            'insert code for Op3head
            oStPoint.Set -0.5 * parB, -parB, 0
            'oEnpoint.Set -0.5 * parB, -parB, 0
             oEnpoint.Set -(0.5 * parB - 0.1 * parH * Cos(PI / 4)), _
                                    -(parB - 0.1 * parH * Sin(PI / 4)), 0


            Set objOp3Head = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, parW, True)

            'set the output
            m_OutputColl.AddOutput "Op3Head", objOp3Head
            Set objOp3Head = Nothing

            If lPartDataBasis = STP_Cock_Bypass Then
                'insert code for Bypass
                Dim oByPassCyl1 As Object
                Dim oByPassCyl2 As Object
                Dim oByPassCyl3 As Object
                Dim oByPassCyl4 As Object
                Dim oByPassCyl5 As Object
                Dim oOp4Sphere As Object
                Dim oOp4Cyl1 As Object
                Dim oOp4Cyl2 As Object
                Dim oOp4Stem As Object
                Dim oOp4Head As Object

                'insert code for Bypass cyl1
                oStPoint.Set -parL / 3, 0, 0
                oEnpoint.Set -parL / 3, 0, -parK * 2 / 3

                Set oByPassCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dBody2Dia, True)

                'set the output
                m_OutputColl.AddOutput "ByPassCyl1", oByPassCyl1
                Set oByPassCyl1 = Nothing

                'insert code for Bypasscyl2 (elbow)
                Dim oCircle As IngrGeom3D.Circle3d
                Dim oCenterPoint As New AutoMath.DPosition
                Dim oGeomFactory As IngrGeom3D.GeometryFactory

                Set oGeomFactory = New IngrGeom3D.GeometryFactory
                oCenterPoint.Set -parL / 3 + parK / 3, 0, -2 * parK / 3
                oDirvec.Set 0, -1, 0

                Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, -parL / 3, 0, -parK * 2 / 3, 0, 0, 1, PipeDiam1 / 2)

                Set oByPassCyl2 = PlaceRevolution(m_OutputColl, oCircle, oDirvec, oCenterPoint, PI / 2, True)

                'set the output
                m_OutputColl.AddOutput "Elbow1", oByPassCyl2
                Set oByPassCyl2 = Nothing
                Set oCenterPoint = Nothing
                Set oDirvec = Nothing
                Set oCircle = Nothing

                'insert code for BypassCyl3
                oStPoint.Set -parL / 3 + parK / 3, 0, -parK
                oEnpoint.Set parL / 3 - parK / 3, 0, -parK

                Set oByPassCyl3 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dBody2Dia, True)

                'set the output
                m_OutputColl.AddOutput "ByPassCyl3", oByPassCyl3
                Set oByPassCyl3 = Nothing

                'insert code for Op4Cyl1
                oStPoint.Set parL / 5 - 0.4 * parW, 0, -parK
                oEnpoint.Set parL / 5 - 0.25 * parW, 0, -parK

                Set oOp4Cyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dValveDia, True)

                m_OutputColl.AddOutput "Op4Cyl1", oOp4Cyl1
                Set oOp4Cyl1 = Nothing

                'insert code for Op4Sphere
                oStPoint.Set parL / 5, 0, -parK

                Set oOp4Sphere = PlaceSphere(m_OutputColl, oStPoint, SphereRad)

                'set the output
                m_OutputColl.AddOutput "Op4Sphere", oOp4Sphere
                Set oOp4Sphere = Nothing

                'insert code for Op4cyl2
                oStPoint.Set parL / 5 + parW * 0.25, 0, -parK
                oEnpoint.Set parL / 5 + parW * 0.4, 0, -parK

                Set oOp4Cyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dValveDia, True)

                'set the output
                m_OutputColl.AddOutput "Op4Cyl2", oOp4Cyl2
                Set oOp4Cyl2 = Nothing

                'insert code for Op4Stem
                oStPoint.Set parL / 5, 0, -parK
                oEnpoint.Set parL / 5, parH * 0.9, -parK

                Set oOp4Stem = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, StemDia, True)

                'set the output
                m_OutputColl.AddOutput "Op4Stem", oOp4Stem
                Set oOp4Stem = Nothing

                'insert code for Op4Head
                oStPoint.Set parL / 5, 0.9 * parH, -parK
                oEnpoint.Set parL / 5, parH, -parK

                Set oOp4Head = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, parW, True)

                'set the output
                m_OutputColl.AddOutput "Op4Head", oOp4Head
                Set oOp4Head = Nothing

                'insert code for Bypasscyl4 (elbow)
                oCenterPoint.Set parL / 3 - parK / 3, 0, -2 * parK / 3
                Set oDirvec = New AutoMath.DVector
                oDirvec.Set 0, 1, 0

                Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parL / 3, 0, -parK * 2 / 3, 0, 0, 1, PipeDiam1 / 2)

                Set oByPassCyl4 = PlaceRevolution(m_OutputColl, oCircle, oDirvec, oCenterPoint, PI / 2, True)

                'set the output
                m_OutputColl.AddOutput "Elbow2", oByPassCyl4
                Set oByPassCyl4 = Nothing
                Set oCenterPoint = Nothing
                Set oDirvec = Nothing
                Set oCircle = Nothing
                Set oGeomFactory = Nothing

                'insert code for BypassCyl5
                oStPoint.Set parL / 3, 0, 0
                oEnpoint.Set parL / 3, 0, -parK * 2 / 3

                Set oByPassCyl5 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dBody2Dia, True)

                'set the output
                m_OutputColl.AddOutput "BypassCyl5", oByPassCyl5
                Set oByPassCyl5 = Nothing

            End If

        End If

        'insert code for output7(body2)
        If lPartDataBasis = STA_Cock Or lPartDataBasis = STP_Cock_Bypass Then
            oStPoint.Set -0.5 * parB, 0, 0
            oEnpoint.Set -(0.5 * parB - Cos(PI / 4) * parB), -Cos(PI / 4) * parB, 0
        Else
            oStPoint.Set -0.5 * parB, 0, 0
            oEnpoint.Set -0.3 * parB, -0.2 * parB, 0
        End If
        Set ObjBody2 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dBody2Dia, True)

        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody2
        Set ObjBody2 = Nothing

        'insert code for output8 nozzle1
        Set oDirvec = New AutoMath.DVector
        oDirvec.Set -1, 0, 0

        oStPoint.Set -parL / 2 + Depth1 - CptOffset1, 0, 0

        Set objNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDirvec, oStPoint)

        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
        Set objNozzle1 = Nothing
        
        Set oDirvec = Nothing

        'insert code for output9 nozzle2
        RetrieveParameters 2, oPartFclt, m_OutputColl, PipeDiam2, FlangeThick2, FlangeDiam2, CptOffset2, Depth2
        Set oDirvec = New AutoMath.DVector
        oDirvec.Set 1, 0, 0
        oStPoint.Set parL / 2 - Depth2 + CptOffset2, 0, 0

        Set objNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDirvec, oStPoint)

        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
        Set objNozzle2 = Nothing
        Set oDirvec = Nothing

        'insert code for output10 nozzle3
        RetrieveParameters 3, oPartFclt, m_OutputColl, PipeDiam3, FlangeThick3, FlangeDiam3, CptOffset3, Depth3
        Set oDirvec = New AutoMath.DVector
        oDirvec.Set Cos(PI / 4), -Sin(PI / 4), 0
        Dim dAngle As Double
        dAngle = PI / 4
        Dim dNozzleLength As Double
        If lPartDataBasis = STA_Cock Or lPartDataBasis = STP_Cock_Bypass Then
            dNozzleLength = parB
        Else
            dNozzleLength = Sqr((0.2 * parB) ^ 2 + (0.2 * parB) ^ 2)
        End If
        oStPoint.Set -0.5 * parB + (dNozzleLength + CptOffset3 - Depth3) * Cos(dAngle), _
            -(dNozzleLength + CptOffset3 - Depth3) * Sin(dAngle), 0
        Set objNozzle3 = CreateNozzle(3, oPartFclt, m_OutputColl, oDirvec, oStPoint)

        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle3
        Set objNozzle3 = Nothing
        Set oDirvec = Nothing
        Set oStPoint = Nothing
        Set oEnpoint = Nothing
    End If
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
