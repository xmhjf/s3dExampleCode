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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, Feb 14 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the Style 443-B and 443-BS in the 440-Series-6page.pdf
'    Proco's catalog available at URL http://www.bigriverrubber.com/proco/expjointlocation.htm
'    For the part data basis 935- 944 - Source - www.ejsus.com
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  15.May.2008     KKC  CR-136574  Enhance expansion joint symbol to provide more realistic dimensional basis
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
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
    Case 935
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        
    Case 936
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parCenterSpoollength = arrayOfInputs(4)
    Case 937
        parFacetoFace = arrayOfInputs(2)
        parOuterCasingLength = arrayOfInputs(5)
        parPurgePortLenoffset = arrayOfInputs(6)
        parPurgePortHeioffset = arrayOfInputs(7)
        parDrainPortLenoffset = arrayOfInputs(8)
        parDrainPortHeioffset = arrayOfInputs(9)
    Case 938
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parAngleFlangeWidth = arrayOfInputs(10)

    Case 939
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parCenterSpoollength = arrayOfInputs(4)

    Case 940
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parHingeWidth = arrayOfInputs(11)

    Case 941
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parGimbalWidth = arrayOfInputs(12)

    Case 942
        parPipingSpecialtyDiameter = arrayOfInputs(3)
        parFace1toCenter = arrayOfInputs(13)
        parFace2toCenter = arrayOfInputs(14)
        parFlangeWidth = arrayOfInputs(15)
    Case 943
        parFacetoFace = arrayOfInputs(2)
        parPipingSpecialtyDiameter = arrayOfInputs(3)

    Case 944
        parFacetoFace = arrayOfInputs(2)
        parOuterCasingLength = arrayOfInputs(5)
        parPurgePortLenoffset = arrayOfInputs(6)
        parPurgePortHeioffset = arrayOfInputs(7)
        parDrainPortLenoffset = arrayOfInputs(8)
        parDrainPortHeioffset = arrayOfInputs(9)

    End Select
    Select Case lPartDataBasis
    Case Is <= 1
        iOutput = 0
        '   Assumption: Origin of the symbol is taken at mid-point of parFacetoFace

        '   Insert your code for output 1(Left Hand Side Flange)
        '   Assumptions:
        '    1. The Body of 'Limit bolt Outside Ear' is taken to be cylinder, and
        '    2. Its thickness is taken to be equal to that of the flange thickness
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1

        Dim stPoint As AutoMath.DPosition
        Dim enPoint As AutoMath.DPosition

        Set stPoint = New AutoMath.DPosition
        Set enPoint = New AutoMath.DPosition

        stPoint.Set -parFacetoFace / 2, 0, 0
        enPoint.Set stPoint.x + flangeThick, 0, 0
        Dim ObjFlCylinder As Object
        Set ObjFlCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parPipingSpecialtyDiameter, True)
        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlCylinder
        Set ObjFlCylinder = Nothing

        '   Insert your code for output 2(Body of Expansion Joint)
        '    Assumption: The Bellows outside diameter is taken to be equal to 1.2 time the pipe diameter
        Dim dBellowOD As Double

        dBellowOD = pipeDiam * 1.2
        stPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        enPoint.Set parFacetoFace / 2 - flangeThick, 0, 0

        Dim ObjBodyCylinder As Object
        Set ObjBodyCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, dBellowOD, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCylinder
        Set ObjBodyCylinder = Nothing

        '   Insert your code for output 3(Right Hand Side Flange)
        stPoint.Set parFacetoFace / 2, 0, 0
        enPoint.Set stPoint.x - flangeThick, 0, 0

        Set ObjFlCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parPipingSpecialtyDiameter, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlCylinder
        Set ObjFlCylinder = Nothing

        Set stPoint = Nothing
        Set enPoint = Nothing

        ' Place Nozzle 1
        Dim oPlacePoint As AutoMath.DPosition
        Dim objNozzle As GSCADNozzleEntities.IJDNozzle

        Set oPlacePoint = New AutoMath.DPosition
        oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0

        Set objNozzle = CreateNozzleJustaCircle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle

        Set objNozzle = Nothing
        Set oDir = Nothing
        Set oPlacePoint = Nothing

    Case 935

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
        Dim OPipePort As IJCatalogPipePort
        Dim oCollection As IJDCollection
        ' Checking the terminaion class for one of the ports
        Set oCollection = oPartFclt.GetNozzles()
        Set OPipePort = oCollection.Item(1)
        If Not OPipePort.TerminationClass = 5 Then
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0

            Dim ObjCylinder As Object
            Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
            Set ObjCylinder = Nothing

            ' Insert your code for the Bellow geometry
            Dim objCircle As Object
            Dim dCircleDia As Double
            Dim dStartingPoint As Double
            dCircleDia = (0.5 * parFacetoFace) / 9    ' Assumption: Five bellow rings and four spaces in between
            dStartingPoint = parPipingSpecialtyDiameter / 2 - dCircleDia / 2
            Set objCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                              -(parFacetoFace / 4 - dCircleDia / 2), 0, dStartingPoint, 0, 1, 0, dCircleDia / 2)

            Set oGeomFactory = Nothing
            'First Bellow Ring
            '   Revolve the circle about the X axis passing by the origin

            Set ocv = New AutoMath.DVector
            Set oCP = New AutoMath.DPosition
            oCP.Set -(parFacetoFace / 4 - dCircleDia / 2), 0, 0
            ocv.Set -1, 0, 0
            Dim objCurvedSurface As Object
            Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                                   oCP, 2 * PI, True)

            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
            Set objCurvedSurface = Nothing

            ' Cylinder
            Dim dCircleDia1 As Double
            dCircleDia1 = parPipingSpecialtyDiameter - dCircleDia / 2
            oStPoint.Set -parFacetoFace / 4, 0, 0
            oEnPoint.Set -(parFacetoFace / 4 - dCircleDia), 0, 0
            Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            '   Set the output
            m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
            Set ObjCylinder = Nothing
            ' Next Bellow ring
            Dim dTransDistance As Double
            Dim i As Long
            For i = 1 To 4
                Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                                       oCP, 2 * PI, True)
                Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
                dTransDistance = 2 * i * dCircleDia
                oTransMatrix.LoadIdentity
                oDir.Set dTransDistance, 0, 0
                oTransMatrix.Translate oDir
                objCurvedSurface.Transform oTransMatrix
                ObjCylinder.Transform oTransMatrix
                iOutput = iOutput + 1
                m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
                m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
                Set objCurvedSurface = Nothing
                Set ObjCylinder = Nothing
            Next i

            ' Place Nozzle 1
            Set oPlacePoint = New AutoMath.DPosition

            oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
            oDir.Set -1, 0, 0
            Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

            '   Set the output
            m_OutputColl.AddOutput "objNozz_", objNozzle
            Set objNozzle = Nothing

            ' Place Nozzle 2
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
            oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
            oDir.Set 1, 0, 0
            Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "objNozz_", objNozzle

            Set objNozzle = Nothing
            Set oDir = Nothing
            Set oPlacePoint = Nothing

        ElseIf OPipePort.TerminationClass = 5 Then

            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1

            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
            Set ObjCylinder = Nothing

            ' Insert your code for the Bellow geometry
            Dim dStartingPointz As Double
            Dim dBellowLength As Double
            dBellowLength = parFacetoFace - 2 * flangeThick

            dCircleDia = dBellowLength / 11    ' Assumption: Five bellow rings and four spaces in between
            dStartingPointz = parPipingSpecialtyDiameter / 2 - dCircleDia / 2
            Dim dStartingPointx As Double
            dStartingPointx = parFacetoFace / 2 - (3 * dCircleDia) / 2 - flangeThick
            Set objCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                              -dStartingPointx, 0, dStartingPointz, 0, 1, 0, dCircleDia / 2)

            Set oGeomFactory = Nothing
            'First Bellow Ring
            '   Revolve the circle about the X axis passing by the origin

            Set ocv = New AutoMath.DVector
            Set oCP = New AutoMath.DPosition
            oCP.Set -dStartingPointx, 0, 0
            ocv.Set -1, 0, 0
            Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                                   oCP, 2 * PI, True)

            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
            Set objCurvedSurface = Nothing

            ' Cylinder
            dCircleDia1 = parPipingSpecialtyDiameter - dCircleDia / 2
            oStPoint.Set -(dStartingPointx + dCircleDia / 2), 0, 0
            oEnPoint.Set -(dStartingPointx - dCircleDia / 2), 0, 0
            Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            '   Set the output
            m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
            Set ObjCylinder = Nothing
            ' Next Bellow ring
            For i = 1 To 4
                Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                                       oCP, 2 * PI, True)
                Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
                dTransDistance = 2 * i * dCircleDia
                oTransMatrix.LoadIdentity
                oDir.Set dTransDistance, 0, 0
                oTransMatrix.Translate oDir
                objCurvedSurface.Transform oTransMatrix
                ObjCylinder.Transform oTransMatrix
                iOutput = iOutput + 1
                m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
                m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
                Set objCurvedSurface = Nothing
                Set ObjCylinder = Nothing
            Next i

            ' Place Nozzle 1
            Set oPlacePoint = New AutoMath.DPosition
            oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
            oDir.Set -1, 0, 0
            Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

            '   Set the output
            m_OutputColl.AddOutput "objNozz_", objNozzle
            Set objNozzle = Nothing

            ' Place Nozzle 2
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
            oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
            oDir.Set 1, 0, 0
            Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

            '   Set the output
            m_OutputColl.AddOutput "objNozz_", objNozzle

            Set objNozzle = Nothing
            Set oDir = Nothing
            Set oPlacePoint = Nothing
        End If
    Case 936
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1

        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Insert your code for the Bellow geometry

        dBellowLength = parFacetoFace - 2 * flangeThick - parCenterSpoollength
        dCircleDia = dBellowLength / 12    ' Assumption: Five bellow rings and four spaces in between
        dStartingPointz = parPipingSpecialtyDiameter / 2 - dCircleDia / 2
        Dim dStartingPointx1 As Double
        dStartingPointx = parFacetoFace / 2 - (3 * dCircleDia) / 2 - flangeThick
        dStartingPointx1 = parCenterSpoollength / 2 + dCircleDia / 2
        Set objCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPointx, 0, dStartingPointz, 0, 1, 0, dCircleDia / 2)
        Dim ObjCircle1 As Object
        Set ObjCircle1 = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          dStartingPointx1, 0, dStartingPointz, 0, 1, 0, dCircleDia / 2)
        Set oGeomFactory = Nothing
        'First Bellow Ring
        '   Revolve the circle about the X axis passing by the origin

        Set ocv = New AutoMath.DVector
        Set oCP = New AutoMath.DPosition
        oCP.Set -dStartingPointx, 0, 0
        ocv.Set -1, 0, 0
        Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                               oCP, 2 * PI, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
        Set objCurvedSurface = Nothing

        Set ocv = New AutoMath.DVector
        Set oCP = New AutoMath.DPosition
        oCP.Set dStartingPointx1, 0, 0
        ocv.Set 1, 0, 0
        Dim objCurvedSurface1 As Object
        Set objCurvedSurface1 = PlaceRevolution(m_OutputColl, ObjCircle1, ocv, _
                                                oCP, 2 * PI, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", objCurvedSurface1
        Set objCurvedSurface1 = Nothing

        ' Cylinder
        dCircleDia1 = parPipingSpecialtyDiameter - dCircleDia / 2
        oStPoint.Set -(dStartingPointx + dCircleDia / 2), 0, 0
        oEnPoint.Set -(dStartingPointx - dCircleDia / 2), 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        For i = 1 To 2
            Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                                   oCP, 2 * PI, True)
            Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            dTransDistance = 2 * i * dCircleDia
            oTransMatrix.LoadIdentity
            oDir.Set dTransDistance, 0, 0
            oTransMatrix.Translate oDir
            objCurvedSurface.Transform oTransMatrix
            ObjCylinder.Transform oTransMatrix
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
            m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
            Set objCurvedSurface = Nothing
            Set ObjCylinder = Nothing
        Next i

        oStPoint.Set dStartingPointx1 - dCircleDia / 2, 0, 0
        oEnPoint.Set dStartingPointx1 + dCircleDia / 2, 0, 0
        Dim ObjCylinder1 As Object
        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder1
        Set ObjCylinder1 = Nothing


        ' Next Bellow ring


        For i = 1 To 2
            Set objCurvedSurface1 = PlaceRevolution(m_OutputColl, ObjCircle1, ocv, _
                                                    oCP, 2 * PI, True)
            Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            dTransDistance = 2 * i * dCircleDia
            oTransMatrix.LoadIdentity
            oDir.Set dTransDistance, 0, 0
            oTransMatrix.Translate oDir
            objCurvedSurface1.Transform oTransMatrix
            ObjCylinder1.Transform oTransMatrix
            m_OutputColl.AddOutput "CurSur_", objCurvedSurface1
            m_OutputColl.AddOutput "ObjCyl1_", ObjCylinder1
            Set objCurvedSurface1 = Nothing
            Set ObjCylinder1 = Nothing
        Next i

        ' Place Nozzle 1
        Set oPlacePoint = New AutoMath.DPosition

        oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        Set oDir = Nothing
        Set oPlacePoint = Nothing
    Case 937

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1

        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Cylinder 1

        oStPoint.Set -parOuterCasingLength / 2, 0, 0
        oEnPoint.Set parOuterCasingLength / 2, 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 2 * pipeDiam, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
        Set ObjCylinder = Nothing

        ' Place Nozzle 1
        Set oPlacePoint = New AutoMath.DPosition

        oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        If CmpDblEqual(parDrainPortHeioffset, 0) Then parDrainPortHeioffset = parPurgePortHeioffset
        If CmpDblEqual(parDrainPortLenoffset, 0) Then parDrainPortLenoffset = parPurgePortLenoffset
 
        ' Place Nozzle 3

        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set -parFacetoFace / 2 + parPurgePortLenoffset, parPurgePortHeioffset + sptOffset - depth, 0
        oDir.Set 0, 1, 0
        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, parPurgePortHeioffset)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 4
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set parFacetoFace / 2 - parDrainPortLenoffset, -(parDrainPortHeioffset + sptOffset - depth), 0
        oDir.Set 0, -1, 0
        Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, oPlacePoint, parDrainPortHeioffset)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        Set oDir = Nothing
        Set oPlacePoint = Nothing
    Case 938
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1

        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Insert your code for the Bellow geometry

        dCircleDia = (0.5 * parFacetoFace) / 9    ' Assumption: Five bellow rings and four spaces in between
        dStartingPoint = parPipingSpecialtyDiameter / 2 - dCircleDia / 2
        Set objCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -(parFacetoFace / 4 - dCircleDia / 2), 0, dStartingPoint, 0, 1, 0, dCircleDia / 2)

        Set oGeomFactory = Nothing
        'First Bellow Ring
        '   Revolve the circle about the X axis passing by the origin

        Set ocv = New AutoMath.DVector
        Set oCP = New AutoMath.DPosition
        oCP.Set -(parFacetoFace / 4 - dCircleDia / 2), 0, 0
        ocv.Set -1, 0, 0
        Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                               oCP, 2 * PI, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
        Set objCurvedSurface = Nothing

        ' Cylinder
        dCircleDia1 = parPipingSpecialtyDiameter - dCircleDia / 2
        oStPoint.Set -parFacetoFace / 4, 0, 0
        oEnPoint.Set -(parFacetoFace / 4 - dCircleDia), 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing
        ' Next Bellow ring
        For i = 1 To 4
            Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                                   oCP, 2 * PI, True)
            Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            dTransDistance = 2 * i * dCircleDia
            oTransMatrix.LoadIdentity
            oDir.Set dTransDistance, 0, 0
            oTransMatrix.Translate oDir
            objCurvedSurface.Transform oTransMatrix
            ObjCylinder.Transform oTransMatrix
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
            m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
            Set objCurvedSurface = Nothing
            Set ObjCylinder = Nothing
        Next i
        ' Place geometry of the box for holding the Tie Rod
        Dim ObjBox As Object
        oStPoint.Set -0.4 * parFacetoFace, -parAngleFlangeWidth / 2, -parAngleFlangeWidth / 2
        oEnPoint.Set -0.35 * parFacetoFace, parAngleFlangeWidth / 2, parAngleFlangeWidth / 2
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        oStPoint.Set 0.4 * parFacetoFace, -parAngleFlangeWidth / 2, -parAngleFlangeWidth / 2
        oEnPoint.Set 0.35 * parFacetoFace, parAngleFlangeWidth / 2, parAngleFlangeWidth / 2
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing


        ' Place geometry of the box 1 for holding the Tie Rod
        oStPoint.Set 0.4 * parFacetoFace, -parAngleFlangeWidth / 2, -parAngleFlangeWidth / 2
        oEnPoint.Set 0.35 * parFacetoFace, parAngleFlangeWidth / 2, parAngleFlangeWidth / 2
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        '   Set the output
        m_OutputColl.AddOutput "ObjCylinder_", ObjBox
        Set ObjBox = Nothing

        ' Place the geometry for the Tie Rods
        oStPoint.Set -0.45 * parFacetoFace, -0.35 * parAngleFlangeWidth, -0.35 * parAngleFlangeWidth
        oEnPoint.Set 0.45 * parFacetoFace, -0.35 * parAngleFlangeWidth, -0.35 * parAngleFlangeWidth
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parAngleFlangeWidth, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Tie rod 2
        oStPoint.Set -0.45 * parFacetoFace, -0.35 * parAngleFlangeWidth, 0.35 * parAngleFlangeWidth
        oEnPoint.Set 0.45 * parFacetoFace, -0.35 * parAngleFlangeWidth, 0.35 * parAngleFlangeWidth
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parAngleFlangeWidth, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Tie rod 3
        oStPoint.Set -0.45 * parFacetoFace, 0.35 * parAngleFlangeWidth, -0.35 * parAngleFlangeWidth
        oEnPoint.Set 0.45 * parFacetoFace, 0.35 * parAngleFlangeWidth, -0.35 * parAngleFlangeWidth
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parAngleFlangeWidth, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Tie rod 4
        oStPoint.Set -0.45 * parFacetoFace, 0.35 * parAngleFlangeWidth, 0.35 * parAngleFlangeWidth
        oEnPoint.Set 0.45 * parFacetoFace, 0.35 * parAngleFlangeWidth, 0.35 * parAngleFlangeWidth
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parAngleFlangeWidth, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Place Nozzle 1
        Set oPlacePoint = New AutoMath.DPosition

        oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        m_OutputColl.AddOutput "objNozz_", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "objNozz_", objNozzle

        Set objNozzle = Nothing
        Set oDir = Nothing
        Set oPlacePoint = Nothing

    Case 939

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1

        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Insert your code for the Bellow geometry

        dBellowLength = parFacetoFace - 2 * flangeThick - parCenterSpoollength
        dCircleDia = dBellowLength / 12    ' Assumption: Five bellow rings and four spaces in between
        dStartingPointz = parPipingSpecialtyDiameter / 2 - dCircleDia / 2
        dStartingPointx = parFacetoFace / 2 - (3 * dCircleDia) / 2 - flangeThick
        dStartingPointx1 = parCenterSpoollength / 2 + dCircleDia / 2
        Set objCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPointx, 0, dStartingPointz, 0, 1, 0, dCircleDia / 2)
        Set ObjCircle1 = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          dStartingPointx1, 0, dStartingPointz, 0, 1, 0, dCircleDia / 2)

        'First Bellow Ring
        '   Revolve the circle about the X axis passing by the origin

        Set ocv = New AutoMath.DVector
        Set oCP = New AutoMath.DPosition
        oCP.Set -dStartingPointx, 0, 0
        ocv.Set -1, 0, 0
        Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                               oCP, 2 * PI, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
        Set objCurvedSurface = Nothing

        Set ocv = New AutoMath.DVector
        Set oCP = New AutoMath.DPosition
        oCP.Set dStartingPointx1, 0, 0
        ocv.Set 1, 0, 0
        Set objCurvedSurface1 = PlaceRevolution(m_OutputColl, ObjCircle1, ocv, _
                                                oCP, 2 * PI, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", objCurvedSurface1
        Set objCurvedSurface1 = Nothing

        ' Cylinder
        dCircleDia1 = parPipingSpecialtyDiameter - dCircleDia / 2
        oStPoint.Set -(dStartingPointx + dCircleDia / 2), 0, 0
        oEnPoint.Set -(dStartingPointx - dCircleDia / 2), 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        For i = 1 To 2
            Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                                   oCP, 2 * PI, True)
            Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            dTransDistance = 2 * i * dCircleDia
            oTransMatrix.LoadIdentity
            oDir.Set dTransDistance, 0, 0
            oTransMatrix.Translate oDir
            objCurvedSurface.Transform oTransMatrix
            ObjCylinder.Transform oTransMatrix
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
            m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
            Set objCurvedSurface = Nothing
            Set ObjCylinder = Nothing
        Next i

        oStPoint.Set dStartingPointx1 - dCircleDia / 2, 0, 0
        oEnPoint.Set dStartingPointx1 + dCircleDia / 2, 0, 0
        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl1_", ObjCylinder1
        Set ObjCylinder1 = Nothing

        ' Next Bellow ring
        For i = 1 To 2
            Set objCurvedSurface1 = PlaceRevolution(m_OutputColl, ObjCircle1, ocv, _
                                                    oCP, 2 * PI, True)
            Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            dTransDistance = 2 * i * dCircleDia
            oTransMatrix.LoadIdentity
            oDir.Set dTransDistance, 0, 0
            oTransMatrix.Translate oDir
            objCurvedSurface1.Transform oTransMatrix
            ObjCylinder1.Transform oTransMatrix
            m_OutputColl.AddOutput "CurSur_", objCurvedSurface1
            m_OutputColl.AddOutput "ObjCyl1_", ObjCylinder1
            Set objCurvedSurface1 = Nothing
            Set ObjCylinder1 = Nothing
        Next i
        ' Place the geometry of the symbol for the Tie Rod holder and the Tie rods

        Dim dLineStringPoints(0 To 14) As Double
        dLineStringPoints(0) = 0.5 * parFacetoFace - 0.8 * flangeThick
        dLineStringPoints(1) = 0.1 * flangeDiam
        dLineStringPoints(2) = 0.6 * flangeDiam

        dLineStringPoints(3) = 0.5 * parFacetoFace - 0.8 * flangeThick
        dLineStringPoints(4) = 0.1 * flangeDiam
        dLineStringPoints(5) = 0.45 * flangeDiam

        dLineStringPoints(6) = 0.5 * parFacetoFace - 0.8 * flangeThick
        dLineStringPoints(7) = -0.1 * flangeDiam
        dLineStringPoints(8) = 0.45 * flangeDiam

        dLineStringPoints(9) = 0.5 * parFacetoFace - 0.8 * flangeThick
        dLineStringPoints(10) = -0.1 * flangeDiam
        dLineStringPoints(11) = 0.6 * flangeDiam

        dLineStringPoints(12) = dLineStringPoints(0)
        dLineStringPoints(13) = dLineStringPoints(1)
        dLineStringPoints(14) = dLineStringPoints(2)

        Dim oLineString As IngrGeom3D.LineString3d
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStringPoints)

        Dim oAxisVect As AutoMath.DVector
        Set oAxisVect = New AutoMath.DVector
        oAxisVect.Set 1, 0, 0

        Dim dLineStringPoints1(0 To 14) As Double
        dLineStringPoints1(0) = -(0.5 * parFacetoFace - 0.8 * flangeThick)
        dLineStringPoints1(1) = 0.1 * flangeDiam
        dLineStringPoints1(2) = 0.6 * flangeDiam

        dLineStringPoints1(3) = -(0.5 * parFacetoFace - 0.8 * flangeThick)
        dLineStringPoints1(4) = 0.1 * flangeDiam
        dLineStringPoints1(5) = 0.45 * flangeDiam

        dLineStringPoints1(6) = -(0.5 * parFacetoFace - 0.8 * flangeThick)
        dLineStringPoints1(7) = -0.1 * flangeDiam
        dLineStringPoints1(8) = 0.45 * flangeDiam

        dLineStringPoints1(9) = -(0.5 * parFacetoFace - 0.8 * flangeThick)
        dLineStringPoints1(10) = -0.1 * flangeDiam
        dLineStringPoints1(11) = 0.6 * flangeDiam

        dLineStringPoints1(12) = dLineStringPoints1(0)
        dLineStringPoints1(13) = dLineStringPoints1(1)
        dLineStringPoints1(14) = dLineStringPoints1(2)

        Dim oLineString1 As IngrGeom3D.LineString3d
        Set oLineString1 = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStringPoints1)

        Dim oAxisVect1 As AutoMath.DVector
        Set oAxisVect1 = New AutoMath.DVector
        oAxisVect1.Set -1, 0, 0

        Dim ObjBody As Object
        Dim ObjBody1 As Object
        Set ObjBody = PlaceProjection(m_OutputColl, oLineString, oAxisVect, 0.6 * flangeThick, True)
        Set ObjBody1 = PlaceProjection(m_OutputColl, oLineString1, oAxisVect1, 0.6 * flangeThick, True)

        ' Transform the Holder to 60 deg with rotaion along x-axis

        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate PI / 3, oAxisVect
        ObjBody.Transform oTransMatrix
        ObjBody1.Transform oTransMatrix

        ' Set the Output
        m_OutputColl.AddOutput "ObjBody_", ObjBody
        Set ObjBody = Nothing

        m_OutputColl.AddOutput "ObjBody_", ObjBody1
        Set ObjBody1 = Nothing


        Set ObjBody = PlaceProjection(m_OutputColl, oLineString, oAxisVect, 0.6 * flangeThick, True)
        Set ObjBody1 = PlaceProjection(m_OutputColl, oLineString1, oAxisVect1, 0.6 * flangeThick, True)

        ' Transform the Holder to 330 deg with rotaion along x-axis

        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate -PI / 3, oAxisVect
        ObjBody.Transform oTransMatrix
        ObjBody1.Transform oTransMatrix

        ' Set the Output
        m_OutputColl.AddOutput "ObjBody_", ObjBody
        Set ObjBody = Nothing

        m_OutputColl.AddOutput "ObjBody_", ObjBody1
        Set ObjBody1 = Nothing

        Set ObjBody = PlaceProjection(m_OutputColl, oLineString, oAxisVect, 0.6 * flangeThick, True)
        Set ObjBody1 = PlaceProjection(m_OutputColl, oLineString1, oAxisVect1, 0.6 * flangeThick, True)
        ' Transform the Holder to 130 deg with rotaion along x-axis

        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate -2 * PI / 3, oAxisVect
        ObjBody.Transform oTransMatrix
        ObjBody1.Transform oTransMatrix

        ' Set the Output
        m_OutputColl.AddOutput "ObjBody_", ObjBody
        Set ObjBody = Nothing

        m_OutputColl.AddOutput "ObjBody_", ObjBody1
        Set ObjBody1 = Nothing

        Set ObjBody = PlaceProjection(m_OutputColl, oLineString, oAxisVect, 0.6 * flangeThick, True)
        Set ObjBody1 = PlaceProjection(m_OutputColl, oLineString1, oAxisVect1, 0.6 * flangeThick, True)

        ' Transform the Holder to 30 deg with rotaion along x-axis

        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate 2 * PI / 3, oAxisVect
        ObjBody.Transform oTransMatrix
        ObjBody1.Transform oTransMatrix

        ' Set the Output
        m_OutputColl.AddOutput "ObjBody_", ObjBody
        Set ObjBody = Nothing

        m_OutputColl.AddOutput "ObjBody_", ObjBody1
        Set ObjBody1 = Nothing

        ' Insert your code for the Cylinder Geometry (Tie Rods)

        oStPoint.Set 1.05 * (parFacetoFace / 2), 0, 0.55 * flangeDiam
        oEnPoint.Set -1.05 * (parFacetoFace / 2), 0, 0.55 * flangeDiam

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * flangeDiam, True)

        ' Transform the Holder to 60 deg with rotaion along x-axis

        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate PI / 3, oAxisVect
        ObjCylinder.Transform oTransMatrix

        ' Set the Output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * flangeDiam, True)

        ' Transform the Holder to 330 deg with rotaion along x-axis

        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate -PI / 3, oAxisVect
        ObjCylinder.Transform oTransMatrix

        ' Set the Output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * flangeDiam, True)

        ' Transform the Holder to 130 deg with rotaion along x-axis

        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate -2 * PI / 3, oAxisVect
        ObjCylinder.Transform oTransMatrix

        ' Set the Output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * flangeDiam, True)

        ' Transform the Holder to 30 deg with rotaion along x-axis

        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate 2 * PI / 3, oAxisVect
        ObjCylinder.Transform oTransMatrix

        ' Set the Output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing
        ' Place Nozzle 1
        Set oPlacePoint = New AutoMath.DPosition

        oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        Set oDir = Nothing
        Set oPlacePoint = Nothing
        
    Case 940

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1

        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Insert your code for the Bellow geometry

        dCircleDia = (0.5 * parFacetoFace) / 9    ' Assumption: Five bellow rings and four spaces in between
        dStartingPoint = parPipingSpecialtyDiameter / 2 - dCircleDia / 2
        Set objCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -(parFacetoFace / 4 - dCircleDia / 2), 0, dStartingPoint, 0, 1, 0, dCircleDia / 2)

        Set oGeomFactory = Nothing
        'First Bellow Ring
        '   Revolve the circle about the X axis passing by the origin

        Set ocv = New AutoMath.DVector
        Set oCP = New AutoMath.DPosition
        oCP.Set -(parFacetoFace / 4 - dCircleDia / 2), 0, 0
        ocv.Set -1, 0, 0
        Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                               oCP, 2 * PI, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
        Set objCurvedSurface = Nothing

        ' Cylinder
        dCircleDia1 = parPipingSpecialtyDiameter - dCircleDia / 2
        oStPoint.Set -parFacetoFace / 4, 0, 0
        oEnPoint.Set -(parFacetoFace / 4 - dCircleDia), 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing
        ' Next Bellow ring
        For i = 1 To 4
            Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                                   oCP, 2 * PI, True)
            Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            dTransDistance = 2 * i * dCircleDia
            oTransMatrix.LoadIdentity
            oDir.Set dTransDistance, 0, 0
            oTransMatrix.Translate oDir
            objCurvedSurface.Transform oTransMatrix
            ObjCylinder.Transform oTransMatrix
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
            m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
            Set objCurvedSurface = Nothing
            Set ObjCylinder = Nothing
        Next i
        ' Place geometry of the box for holding the Hinges

        oStPoint.Set -0.4 * parFacetoFace, -parHingeWidth / 2, -parHingeWidth / 2
        oEnPoint.Set -0.35 * parFacetoFace, parHingeWidth / 2, parHingeWidth / 2
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        oStPoint.Set 0.4 * parFacetoFace, -parHingeWidth / 2, -parHingeWidth / 2
        oEnPoint.Set 0.35 * parFacetoFace, parHingeWidth / 2, parHingeWidth / 2
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        ' Insert your code for the Hinges
        oStPoint.Set -0.4 * parFacetoFace, -0.1 * parHingeWidth, -parHingeWidth / 2
        oEnPoint.Set 0.4 * parFacetoFace, 0.1 * parHingeWidth, -0.4 * parHingeWidth
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        ' Insert your code for the Hinges
        oStPoint.Set -0.4 * parFacetoFace, -0.1 * parHingeWidth, parHingeWidth / 2
        oEnPoint.Set 0.4 * parFacetoFace, 0.1 * parHingeWidth, 0.4 * parHingeWidth
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        ' Place Nozzle 1
        Set oPlacePoint = New AutoMath.DPosition

        oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        m_OutputColl.AddOutput "objNozz_", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "objNozz_", objNozzle

        Set objNozzle = Nothing
        Set oDir = Nothing
        Set oPlacePoint = Nothing
    Case 941

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1

        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Insert your code for the Bellow geometry

        dCircleDia = (0.5 * parFacetoFace) / 9    ' Assumption: Five bellow rings and four spaces in between
        dStartingPoint = parPipingSpecialtyDiameter / 2 - dCircleDia / 2
        Set objCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -(parFacetoFace / 4 - dCircleDia / 2), 0, dStartingPoint, 0, 1, 0, dCircleDia / 2)

        Set oGeomFactory = Nothing
        'First Bellow Ring
        '   Revolve the circle about the X axis passing by the origin

        Set ocv = New AutoMath.DVector
        Set oCP = New AutoMath.DPosition
        oCP.Set -(parFacetoFace / 4 - dCircleDia / 2), 0, 0
        ocv.Set -1, 0, 0
        Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                               oCP, 2 * PI, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
        Set objCurvedSurface = Nothing

        ' Cylinder
        dCircleDia1 = parPipingSpecialtyDiameter - dCircleDia / 2
        oStPoint.Set -parFacetoFace / 4, 0, 0
        oEnPoint.Set -(parFacetoFace / 4 - dCircleDia), 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing
        ' Next Bellow ring
        For i = 1 To 4
            Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                                   oCP, 2 * PI, True)
            Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            dTransDistance = 2 * i * dCircleDia
            oTransMatrix.LoadIdentity
            oDir.Set dTransDistance, 0, 0
            oTransMatrix.Translate oDir
            objCurvedSurface.Transform oTransMatrix
            ObjCylinder.Transform oTransMatrix
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
            m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
            Set objCurvedSurface = Nothing
            Set ObjCylinder = Nothing
        Next i
        ' Place geometry of the box for holding the Tie Rod

        oStPoint.Set -0.4 * parFacetoFace, -parGimbalWidth / 2, -parGimbalWidth / 2
        oEnPoint.Set -0.35 * parFacetoFace, parGimbalWidth / 2, parGimbalWidth / 2
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        oStPoint.Set 0.4 * parFacetoFace, -parGimbalWidth / 2, -parGimbalWidth / 2
        oEnPoint.Set 0.35 * parFacetoFace, parGimbalWidth / 2, parGimbalWidth / 2
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        oStPoint.Set -0.05 * parFacetoFace, -0.1 * parGimbalWidth, -parGimbalWidth / 2
        oEnPoint.Set 0.4 * parFacetoFace, 0.1 * parGimbalWidth, -0.4 * parGimbalWidth
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        ' Insert your code for the Hinges
        oStPoint.Set -0.05 * parFacetoFace, -0.1 * parGimbalWidth, parGimbalWidth / 2
        oEnPoint.Set 0.4 * parFacetoFace, 0.1 * parGimbalWidth, 0.4 * parGimbalWidth
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        oStPoint.Set 0.05 * parFacetoFace, -parGimbalWidth / 2, -0.1 * parGimbalWidth
        oEnPoint.Set -0.4 * parFacetoFace, -0.4 * parGimbalWidth, 0.1 * parGimbalWidth
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        oStPoint.Set 0.05 * parFacetoFace, parGimbalWidth / 2, -0.1 * parGimbalWidth
        oEnPoint.Set -0.4 * parFacetoFace, 0.4 * parGimbalWidth, 0.1 * parGimbalWidth
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        oStPoint.Set -0.05 * parFacetoFace, parGimbalWidth / 2, -0.5 * parGimbalWidth
        oEnPoint.Set 0.05 * parFacetoFace, 0.4 * parGimbalWidth, 0.5 * parGimbalWidth
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        oStPoint.Set -0.05 * parFacetoFace, -parGimbalWidth / 2, -0.5 * parGimbalWidth
        oEnPoint.Set 0.05 * parFacetoFace, -0.4 * parGimbalWidth, 0.5 * parGimbalWidth
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        oStPoint.Set -0.05 * parFacetoFace, -0.5 * parGimbalWidth, parGimbalWidth / 2
        oEnPoint.Set 0.05 * parFacetoFace, 0.5 * parGimbalWidth, 0.4 * parGimbalWidth
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        oStPoint.Set -0.05 * parFacetoFace, -0.5 * parGimbalWidth, -parGimbalWidth / 2
        oEnPoint.Set 0.05 * parFacetoFace, 0.5 * parGimbalWidth, -0.4 * parGimbalWidth
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        '   Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        ' Place Nozzle 1
        Set oPlacePoint = New AutoMath.DPosition

        oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        m_OutputColl.AddOutput "objNozz_", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "objNozz_", objNozzle

        Set objNozzle = Nothing
        Set oDir = Nothing
        Set oPlacePoint = Nothing

    Case 942
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
        parFacetoFace = 2 * parFace1toCenter
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Insert your code for the Bellow geometry
        parCenterSpoollength = 2 * parFace2toCenter
        dBellowLength = parFacetoFace - 2 * flangeThick - parCenterSpoollength
        dCircleDia = dBellowLength / 12    ' Assumption: Five bellow rings and four spaces in between
        dStartingPointz = parPipingSpecialtyDiameter / 2 - dCircleDia / 2
        dStartingPointx = parFacetoFace / 2 - (3 * dCircleDia) / 2 - flangeThick
        dStartingPointx1 = parCenterSpoollength / 2 + dCircleDia / 2
        Set objCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPointx, 0, dStartingPointz, 0, 1, 0, dCircleDia / 2)
        Set ObjCircle1 = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          dStartingPointx1, 0, dStartingPointz, 0, 1, 0, dCircleDia / 2)

        'First Bellow Ring
        '   Revolve the circle about the X axis passing by the origin

        Set ocv = New AutoMath.DVector
        Set oCP = New AutoMath.DPosition
        oCP.Set -dStartingPointx, 0, 0
        ocv.Set -1, 0, 0
        Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                               oCP, 2 * PI, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
        Set objCurvedSurface = Nothing

        Set ocv = New AutoMath.DVector
        Set oCP = New AutoMath.DPosition
        oCP.Set dStartingPointx1, 0, 0
        ocv.Set 1, 0, 0
        Set objCurvedSurface1 = PlaceRevolution(m_OutputColl, ObjCircle1, ocv, _
                                                oCP, 2 * PI, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", objCurvedSurface1
        Set objCurvedSurface1 = Nothing

        ' Cylinder
        dCircleDia1 = parPipingSpecialtyDiameter - dCircleDia / 2
        oStPoint.Set -(dStartingPointx + dCircleDia / 2), 0, 0
        oEnPoint.Set -(dStartingPointx - dCircleDia / 2), 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        For i = 1 To 2
            Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                                   oCP, 2 * PI, True)
            Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            dTransDistance = 2 * i * dCircleDia
            oTransMatrix.LoadIdentity
            oDir.Set dTransDistance, 0, 0
            oTransMatrix.Translate oDir
            objCurvedSurface.Transform oTransMatrix
            ObjCylinder.Transform oTransMatrix
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
            m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
            Set objCurvedSurface = Nothing
            Set ObjCylinder = Nothing
        Next i

        oStPoint.Set dStartingPointx1 - dCircleDia / 2, 0, 0
        oEnPoint.Set dStartingPointx1 + dCircleDia / 2, 0, 0
        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl1_", ObjCylinder1
        Set ObjCylinder1 = Nothing

        ' Next Bellow ring
        For i = 1 To 2
            Set objCurvedSurface1 = PlaceRevolution(m_OutputColl, ObjCircle1, ocv, _
                                                    oCP, 2 * PI, True)
            Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            dTransDistance = 2 * i * dCircleDia
            oTransMatrix.LoadIdentity
            oDir.Set dTransDistance, 0, 0
            oTransMatrix.Translate oDir
            objCurvedSurface1.Transform oTransMatrix
            ObjCylinder1.Transform oTransMatrix
            m_OutputColl.AddOutput "CurSur_", objCurvedSurface1
            m_OutputColl.AddOutput "ObjCyl1_", ObjCylinder1
            Set objCurvedSurface1 = Nothing
            Set ObjCylinder1 = Nothing
        Next i

        ' Place Box holder for the Tie Rods
        oStPoint.Set parFacetoFace / 2, -parFlangeWidth / 2, -parFlangeWidth / 2
        oEnPoint.Set 0.45 * parFacetoFace, parFlangeWidth / 2, parFlangeWidth / 2

        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        ' Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        ' Place Box holder for the Tie Rods
        oStPoint.Set -parFacetoFace / 2, -parFlangeWidth / 2, -parFlangeWidth / 2
        oEnPoint.Set -0.45 * parFacetoFace, parFlangeWidth / 2, parFlangeWidth / 2

        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        ' Set the output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing
        ' Insert your code for the Tie Rod Geometery
        oStPoint.Set 1.05 * parFacetoFace / 2, -0.4 * parFlangeWidth, -0.4 * parFlangeWidth
        oEnPoint.Set -1.05 * parFacetoFace / 2, -0.4 * parFlangeWidth, -0.4 * parFlangeWidth

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parFlangeWidth, True)
        ' Set the Output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Insert your code for the Tie Rod Geometery
        oStPoint.Set 1.05 * parFacetoFace / 2, 0.4 * parFlangeWidth, 0.4 * parFlangeWidth
        oEnPoint.Set -1.05 * parFacetoFace / 2, 0.4 * parFlangeWidth, 0.4 * parFlangeWidth

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parFlangeWidth, True)
        ' Set the Output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Insert your code for the Tie Rod Geometery
        oStPoint.Set 1.05 * parFacetoFace / 2, 0.4 * parFlangeWidth, -0.4 * parFlangeWidth
        oEnPoint.Set -1.05 * parFacetoFace / 2, 0.4 * parFlangeWidth, -0.4 * parFlangeWidth

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parFlangeWidth, True)
        ' Set the Output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Insert your code for the Tie Rod Geometery
        oStPoint.Set 1.05 * parFacetoFace / 2, -0.4 * parFlangeWidth, 0.4 * parFlangeWidth
        oEnPoint.Set -1.05 * parFacetoFace / 2, -0.4 * parFlangeWidth, 0.4 * parFlangeWidth
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.1 * parFlangeWidth, True)

        ' Set the Output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Insert your code for the Outlet Curve
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
        dStartingPoint = parFace2toCenter

        Set objCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, pipeDiam / 2)

        Set oGeomFactory = Nothing
        '   Revolve the circle about the Z axis passing by the origin

        Set ocv = New AutoMath.DVector

        oCP.Set -parFace2toCenter, -parFace2toCenter, 0
        ocv.Set 0, 0, -1
        Dim parAngle As Double
        parAngle = PI / 2
        Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                               oCP, parAngle, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", objCurvedSurface
        Set objCurvedSurface = Nothing

        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
        Set oPlacePoint = New AutoMath.DPosition
        oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

        oPlacePoint.Set 0, -parFace2toCenter - sptOffset + depth, 0
        oDir.Set 0, -1, 0

        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        Set oPlacePoint = Nothing
        Set oDir = Nothing

    Case 943

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1

        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Insert your code for the Bellow geometry

        dBellowLength = 0.5 * parFacetoFace
        dCircleDia = dBellowLength / 17
        dStartingPointz = parPipingSpecialtyDiameter / 2 - dCircleDia / 2
        dStartingPointx = 0.45 * parFacetoFace - dCircleDia
        dStartingPointx1 = 0.3 * parFacetoFace + dCircleDia / 2
        Set objCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPointx, 0, dStartingPointz, 0, 1, 0, dCircleDia / 2)
        Set ObjCircle1 = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          dStartingPointx1, 0, dStartingPointz, 0, 1, 0, dCircleDia / 2)

        'First Bellow Ring
        '   Revolve the circle about the X axis passing by the origin

        Set ocv = New AutoMath.DVector
        Set oCP = New AutoMath.DPosition
        oCP.Set -dStartingPointx, 0, 0
        ocv.Set -1, 0, 0
        Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                               oCP, 2 * PI, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
        Set objCurvedSurface = Nothing

        Set ocv = New AutoMath.DVector
        Set oCP = New AutoMath.DPosition
        oCP.Set dStartingPointx1, 0, 0
        ocv.Set 1, 0, 0
        Set objCurvedSurface1 = PlaceRevolution(m_OutputColl, ObjCircle1, ocv, _
                                                oCP, 2 * PI, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", objCurvedSurface1
        Set objCurvedSurface1 = Nothing

        ' Cylinder
        dCircleDia1 = parPipingSpecialtyDiameter - dCircleDia / 2
        oStPoint.Set -(dStartingPointx + dCircleDia / 2), 0, 0
        oEnPoint.Set -(dStartingPointx - dCircleDia / 2), 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        For i = 1 To 2
            Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                                   oCP, 2 * PI, True)
            Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            dTransDistance = 2 * i * dCircleDia
            oTransMatrix.LoadIdentity
            oDir.Set dTransDistance, 0, 0
            oTransMatrix.Translate oDir
            objCurvedSurface.Transform oTransMatrix
            ObjCylinder.Transform oTransMatrix
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
            m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
            Set objCurvedSurface = Nothing
            Set ObjCylinder = Nothing
        Next i
        oStPoint.Set dStartingPointx1 - dCircleDia / 2, 0, 0
        oEnPoint.Set dStartingPointx1 + dCircleDia / 2, 0, 0
        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl1_", ObjCylinder1
        Set ObjCylinder1 = Nothing

        ' Next Bellow ring
        For i = 1 To 2
            Set objCurvedSurface1 = PlaceRevolution(m_OutputColl, ObjCircle1, ocv, _
                                                    oCP, 2 * PI, True)
            Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            dTransDistance = 2 * i * dCircleDia
            oTransMatrix.LoadIdentity
            oDir.Set dTransDistance, 0, 0
            oTransMatrix.Translate oDir
            objCurvedSurface1.Transform oTransMatrix
            ObjCylinder1.Transform oTransMatrix
            m_OutputColl.AddOutput "CurSur_", objCurvedSurface1
            m_OutputColl.AddOutput "ObjCyl1_", ObjCylinder1
            Set objCurvedSurface1 = Nothing
            Set ObjCylinder1 = Nothing
        Next i

        ' Insert your code for the geometry Balancing Bellows
        dStartingPointx = 0.1 * parFacetoFace - dCircleDia
        dStartingPointz = 2 * parPipingSpecialtyDiameter / 2 - dCircleDia / 2
        Set objCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPointx, 0, dStartingPointz, 0, 1, 0, dCircleDia / 2)

        'First Bellow Ring
        '   Revolve the circle about the X axis passing by the origin

        Set ocv = New AutoMath.DVector
        Set oCP = New AutoMath.DPosition
        oCP.Set -dStartingPointx, 0, 0
        ocv.Set -1, 0, 0
        Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                               oCP, 2 * PI, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurvedSurface
        Set objCurvedSurface = Nothing

        ' Cylinder
        dCircleDia1 = 2 * parPipingSpecialtyDiameter - dCircleDia / 2
        oStPoint.Set -(dStartingPointx + dCircleDia / 2), 0, 0
        oEnPoint.Set -(dStartingPointx - dCircleDia / 2), 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        For i = 1 To 3
            Set objCurvedSurface = PlaceRevolution(m_OutputColl, objCircle, ocv, _
                                                   oCP, 2 * PI, True)
            Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCircleDia1, True)
            dTransDistance = 2 * i * dCircleDia
            oTransMatrix.LoadIdentity
            oDir.Set dTransDistance, 0, 0
            oTransMatrix.Translate oDir
            objCurvedSurface.Transform oTransMatrix
            ObjCylinder.Transform oTransMatrix
            m_OutputColl.AddOutput "ObjCur_", objCurvedSurface
            m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
            Set objCurvedSurface = Nothing
            Set ObjCylinder = Nothing
        Next i

        dLineStringPoints(0) = -0.5 * parFacetoFace
        dLineStringPoints(1) = 1.25 * parPipingSpecialtyDiameter
        dLineStringPoints(2) = 1.25 * parPipingSpecialtyDiameter

        dLineStringPoints(3) = -0.5 * parFacetoFace
        dLineStringPoints(4) = -1.25 * parPipingSpecialtyDiameter
        dLineStringPoints(5) = 1.25 * parPipingSpecialtyDiameter

        dLineStringPoints(6) = -0.5 * parFacetoFace
        dLineStringPoints(7) = -1.25 * parPipingSpecialtyDiameter
        dLineStringPoints(8) = -1.25 * parPipingSpecialtyDiameter

        dLineStringPoints(9) = -0.5 * parFacetoFace
        dLineStringPoints(10) = 1.25 * parPipingSpecialtyDiameter
        dLineStringPoints(11) = -1.25 * parPipingSpecialtyDiameter

        dLineStringPoints(12) = dLineStringPoints(0)
        dLineStringPoints(13) = dLineStringPoints(1)
        dLineStringPoints(14) = dLineStringPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStringPoints)

        Set oAxisVect = New AutoMath.DVector
        oAxisVect.Set 1, 0, 0

        Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisVect, 0.05 * parFacetoFace, True)

        ' Set the Output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisVect, 0.05 * parFacetoFace, True)
        oTransMatrix.LoadIdentity
        oAxisVect.Set 1, 0, 0
        oTransMatrix.Rotate PI / 4, oAxisVect
        oAxisVect.Set 0.35 * parFacetoFace, 0, 0
        oTransMatrix.Translate oAxisVect
        ObjBox.Transform oTransMatrix

        ' Set the Output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisVect, 0.05 * parFacetoFace, True)
        oTransMatrix.LoadIdentity
        oAxisVect.Set 0.6 * parFacetoFace, 0, 0
        oTransMatrix.Translate oAxisVect
        ObjBox.Transform oTransMatrix

        ' Set the Output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisVect, 0.05 * parFacetoFace, True)
        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate PI / 4, oAxisVect
        oAxisVect.Set 0.95 * parFacetoFace, 0, 0
        oTransMatrix.Translate oAxisVect
        ObjBox.Transform oTransMatrix

        ' Set the Output
        m_OutputColl.AddOutput "ObjBox_", ObjBox
        Set ObjBox = Nothing

        ' Tie Rods for the expansion joints
        oStPoint.Set -0.5 * parFacetoFace, 0.95 * parPipingSpecialtyDiameter, 0.95 * parPipingSpecialtyDiameter
        oEnPoint.Set 0.2 * parFacetoFace, 0.95 * parPipingSpecialtyDiameter, 0.95 * parPipingSpecialtyDiameter

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.08 * parPipingSpecialtyDiameter, True)
        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.08 * parPipingSpecialtyDiameter, True)
        ' Set the Output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate PI / 4, oAxisVect
        oAxisVect.Set 0.3 * parFacetoFace, 0, 0
        oTransMatrix.Translate oAxisVect
        ObjCylinder1.Transform oTransMatrix
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder1
        Set ObjCylinder1 = Nothing

        ' Tie Rods for the expansion joints
        oStPoint.Set -0.5 * parFacetoFace, -0.95 * parPipingSpecialtyDiameter, 0.95 * parPipingSpecialtyDiameter
        oEnPoint.Set 0.2 * parFacetoFace, -0.95 * parPipingSpecialtyDiameter, 0.95 * parPipingSpecialtyDiameter

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.08 * parPipingSpecialtyDiameter, True)
        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.08 * parPipingSpecialtyDiameter, True)
        ' Set the Output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate PI / 4, oAxisVect
        oAxisVect.Set 0.3 * parFacetoFace, 0, 0
        oTransMatrix.Translate oAxisVect
        ObjCylinder1.Transform oTransMatrix
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder1
        Set ObjCylinder1 = Nothing

        ' Tie Rods for the expansion joints
        oStPoint.Set -0.5 * parFacetoFace, 0.95 * parPipingSpecialtyDiameter, -0.95 * parPipingSpecialtyDiameter
        oEnPoint.Set 0.2 * parFacetoFace, 0.95 * parPipingSpecialtyDiameter, -0.95 * parPipingSpecialtyDiameter

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.08 * parPipingSpecialtyDiameter, True)
        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.08 * parPipingSpecialtyDiameter, True)
        ' Set the Output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate PI / 4, oAxisVect
        oAxisVect.Set 0.3 * parFacetoFace, 0, 0
        oTransMatrix.Translate oAxisVect
        ObjCylinder1.Transform oTransMatrix
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder1
        Set ObjCylinder1 = Nothing

        ' Tie Rods for the expansion joints
        oStPoint.Set -0.5 * parFacetoFace, -0.95 * parPipingSpecialtyDiameter, -0.95 * parPipingSpecialtyDiameter
        oEnPoint.Set 0.2 * parFacetoFace, -0.95 * parPipingSpecialtyDiameter, -0.95 * parPipingSpecialtyDiameter

        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.08 * parPipingSpecialtyDiameter, True)
        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.08 * parPipingSpecialtyDiameter, True)
        ' Set the Output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder
        Set ObjCylinder = Nothing

        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate PI / 4, oAxisVect
        oAxisVect.Set 0.3 * parFacetoFace, 0, 0
        oTransMatrix.Translate oAxisVect
        ObjCylinder1.Transform oTransMatrix
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", ObjCylinder1
        Set ObjCylinder1 = Nothing

        ' Place Nozzle 1
        Set oPlacePoint = New AutoMath.DPosition

        oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        m_OutputColl.AddOutput "objNozz_", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle

        Set objNozzle = Nothing
        Set oDir = Nothing
        Set oPlacePoint = Nothing
    Case 944

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1

        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder
        Set ObjCylinder = Nothing

        ' Cylinder 1

        oStPoint.Set -parOuterCasingLength / 2, 0, 0
        oEnPoint.Set parOuterCasingLength / 2, 0, 0
        Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 2 * pipeDiam, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
        Set ObjCylinder = Nothing

        ' Place Nozzle 1
        Set oPlacePoint = New AutoMath.DPosition

        oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        If CmpDblEqual(parDrainPortHeioffset, 0) Then parDrainPortHeioffset = parPurgePortHeioffset
        If CmpDblEqual(parDrainPortLenoffset, 0) Then parDrainPortLenoffset = parPurgePortLenoffset

        ' Place Nozzle 3

        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set -parFacetoFace / 2 + parPurgePortLenoffset, parPurgePortHeioffset + sptOffset - depth, 0
        oDir.Set 0, 1, 0
        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, parPurgePortHeioffset)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 4
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set parFacetoFace / 2 - parDrainPortLenoffset, -(parDrainPortHeioffset + sptOffset - depth), 0
        oDir.Set 0, -1, 0
        Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, oPlacePoint, parDrainPortHeioffset)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        Set oDir = Nothing
        Set oPlacePoint = Nothing

    End Select

    Set oCP = Nothing
    Set ocv = Nothing
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
              Err.HelpFile, Err.HelpContext

End Sub

