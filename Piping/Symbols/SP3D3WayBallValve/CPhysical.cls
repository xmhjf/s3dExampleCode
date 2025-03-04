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
'   Copyright (c) 2003-2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Friday, Jan 4 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   3-way ball valve, T-type(963):www.xulongindustry.com/en/main/index/productshow.asp?cid=2043
'   3-way ball valve, L-type(964):www.xulongindustry.com/en/main/index/productshow.asp?cid=2042
'   3-way ball valve, Y-type(965):www.xulongindustry.com/en/main/index/productshow.asp?cid=2044

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   24.Aug.2004     MS     TR 62189         Delivered 3s-Way Ball Valve symbol does not place.A value of 0.0000001 is used in Z direction to avoid this problem.
'   03.Dec.2004     svsmylav                Defined new input parameter HandwheelAngle and modified valve operator direction vectors.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   08.SEP.2006     KKC     DI-95670        Replace names with initials in all revision history sheets and symbols
'   30.JAN.2008     PK      CR-135510       Enhanced the existing 3 way ball valve to support dimensional basis for
'                                               (i)Face-to-center dimension basis
'                                               (ii)Face-to-Face dimension and Face-to-Center dimension for port 3
'                                               (iii)Asymmetrical Face-to-Center dimension basis
' 05.May.2008      VRK                      CR-CP-141546  Enhance master 3-way ball valve symbol for Y-type, L-type, and T-type
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private Const MODULE = "Physical:"    'Used for error messages
Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub
Private Function ReturnMax3(A#, B#, C#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    ReturnMax3 = MaxValue
End Function

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim sptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double
    Dim pipeDiam2 As Double
    Dim flangeThick2 As Double
    Dim sptOffset2 As Double
    Dim flangeDiam2 As Double
    Dim depth2 As Double
    Dim pipeDiam3 As Double
    Dim flangeThick3 As Double
    Dim sptOffset3 As Double
    Dim flangeDiam3 As Double
    Dim depth3 As Double
    Dim CenterPos As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput As Double
    Dim ObjBodyBall As Object
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjBodyCone3 As Object
    Dim objCylinder As Object

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double

    'Variables defined by svsmylav
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    Dim BallDiameter As Double
    Dim parHandwheelAngle As Double

    Dim parFacetoFace As Double
    Dim parFacetoCenter As Double

    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dFace3toCenter As Double
    Dim oEllipse As IngrGeom3D.Ellipse3d
    Dim objEllTop As Object
    Dim dThickness As Double
    Dim dEllipseMajor As Double
    Dim ObjBox As Object
    Dim dLineStrPoints(0 To 14) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oAxisvec As New AutoMath.DVector
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New GeometryFactory

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFace1toCenter = arrayOfInputs(2)
    'parFace2toCenter = arrayOfInputs(3)
    'parFace3toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    'parFacetoFace = arrayOfInputs(7)
    'parFacetoCenter = arrayOfInputs(8)

    dEllipseMajor = 0.06
    
    iOutput = 0

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                       flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                       flangeThick2, flangeDiam2, sptOffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, _
                       flangeThick3, flangeDiam3, sptOffset3, depth3
                       
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartdatabasis As Long

    Set oPipeComponent = oPartFclt
    lPartdatabasis = oPipeComponent.PartDataBasis
    
    If lPartdatabasis = 963 Or lPartdatabasis = 964 Or _
                lPartdatabasis = 965 Then
        'Checking for zero flange diameter and thickness
        If CmpDblEqual(flangeDiam, 0) Then
            flangeDiam = 1.75 * pipeDiam
        End If
        If CmpDblEqual(flangeDiam2, 0) Then
            flangeDiam2 = 1.75 * pipeDiam2
        End If
        If CmpDblEqual(flangeDiam3, 0) Then
            flangeDiam3 = 1.75 * pipeDiam3
        End If
        
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = 0.0001
        End If
        If CmpDblEqual(flangeThick2, 0) Then
            flangeThick2 = 0.0001
        End If
        If CmpDblEqual(flangeThick3, 0) Then
            flangeThick3 = 0.0001
        End If
   End If
    
    Select Case lPartdatabasis

    Case Is <= 1, 13                        'Asymmetrical Face-to-Center dimension basis
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
        parFace3toCenter = arrayOfInputs(4)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        dFace3toCenter = parFace3toCenter
        If dFace1toCenter = 0 Then
            dFace1toCenter = dFace2toCenter
        ElseIf dFace2toCenter = 0 Then
            dFace2toCenter = dFace1toCenter
        End If

    Case 10                                 'Face-to-center dimension basis
        parFacetoCenter = arrayOfInputs(8)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
        dFace3toCenter = parFacetoCenter

    Case 11                                 'Face-to-Face dimension and Face-to-Center dimension for port 3
        parFace3toCenter = arrayOfInputs(4)
        parFacetoFace = arrayOfInputs(7)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
        dFace3toCenter = parFace3toCenter
        
    Case 963                                '3 Way Ball Valve, T-Type
        parFace3toCenter = arrayOfInputs(4)
        parFacetoFace = arrayOfInputs(7)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
        dFace3toCenter = parFace3toCenter
        
        If CmpDblGreaterthan(dEllipseMajor, dFace1toCenter / 2) Then
            dEllipseMajor = 0.35 * dFace1toCenter
        End If
        
        Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(m_OutputColl.ResourceManager, _
        CenterPos.x, CenterPos.y + 0.4 * parFacetoFace, CenterPos.z, 0, 1, 0, dEllipseMajor, 0, 0, 0.5)
        oAxisvec.Set 0, 1, 0
        Set objEllTop = PlaceProjection(m_OutputColl, oEllipse, oAxisvec, 0.05 * parFacetoFace, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEllTop
        Set oEllipse = Nothing
        Set objEllTop = Nothing
        
        'Top Disc
        oStPoint.Set 0, 0.3 * parFacetoFace, 0
        oEnPoint.Set 0, 0.4 * parFacetoFace, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * parFacetoFace, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
        
        oStPoint.Set -0.75 * (dFace1toCenter - flangeThick), 0, 0
        oEnPoint.Set 0.75 * (dFace2toCenter - flangeThick2), 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.2 * pipeDiam, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
        
        oStPoint.Set 0, 0.3 * parFacetoFace, 0
        oEnPoint.Set 0, -0.3 * parFacetoFace, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.4 * parFacetoFace, True)
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", objCylinder
        Set objCylinder = Nothing
        
        oStPoint.Set -0.75 * (dFace1toCenter - flangeThick), 0, 0
        oEnPoint.Set -dFace1toCenter + flangeThick, 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", objCylinder
        Set objCylinder = Nothing
        
        oStPoint.Set 0.75 * (dFace2toCenter - flangeThick2), 0, 0
        oEnPoint.Set dFace2toCenter - flangeThick2, 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam2, True)
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", objCylinder
        Set objCylinder = Nothing
        
        oStPoint.Set 0, 0, 0.3 * parFacetoFace
        oEnPoint.Set 0, 0, -0.3 * parFacetoFace
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.4 * parFacetoFace, True)
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", objCylinder
        Set objCylinder = Nothing
        
        oStPoint.Set 0, 0, 0.3 * parFacetoFace
        oEnPoint.Set 0, 0, oStPoint.z + 0.75 * ((dFace3toCenter - flangeThick3) - 0.3 * parFacetoFace)
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.2 * pipeDiam3, True)
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", objCylinder
        Set objCylinder = Nothing
        
        oStPoint.Set 0, 0, oEnPoint.z
        oEnPoint.Set 0, 0, (dFace3toCenter - flangeThick3)
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam3, True)
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", objCylinder
        Set objCylinder = Nothing
        
     Case 964                   '3 Way Ball Valve, L-Type
        parFace3toCenter = arrayOfInputs(4)
        parFacetoFace = arrayOfInputs(7)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
        dFace3toCenter = parFace3toCenter
        
        If CmpDblGreaterthan(dEllipseMajor, dFace1toCenter / 2) Then
            dEllipseMajor = 0.35 * dFace1toCenter
        End If
        
        Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(m_OutputColl.ResourceManager, _
                             CenterPos.x, CenterPos.y + 0.5 * (pipeDiam + flangeDiam) / 2, _
                             CenterPos.z, 0, 1, 0, 0, 0, dEllipseMajor, 0.5)
        oAxisvec.Set 0, 1, 0
        Set objEllTop = PlaceProjection(m_OutputColl, oEllipse, oAxisvec, 0.05 * parFacetoFace, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEllTop
        Set oEllipse = Nothing
        Set objEllTop = Nothing
        
        oStPoint.Set -dFace1toCenter + flangeThick, 0, 0
        oEnPoint.Set -0.5 * (dFace1toCenter + flangeThick), 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
        
        oStPoint.Set dFace2toCenter - flangeThick2, 0, 0
        oEnPoint.Set 0.6 * (dFace2toCenter - flangeThick2), 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam2, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
        
        oStPoint.Set -0.5 * (dFace1toCenter + flangeThick), 0, 0
        oEnPoint.Set 0.4 * (dFace2toCenter - flangeThick2), 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * (pipeDiam + flangeDiam), True)
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", objCylinder
        Set objCylinder = Nothing
        
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, 0, dFace3toCenter - flangeThick3
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.4 * (pipeDiam3 + flangeDiam3), True)
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", objCylinder
        Set objCylinder = Nothing
        
        'Box
        dThickness = 0.0625 * 0.0254 '0.0625in

        dLineStrPoints(0) = 0.4 * (dFace2toCenter - flangeThick2)
        dLineStrPoints(1) = -0.5 * flangeDiam2
        dLineStrPoints(2) = 0.5 * flangeDiam2
        
        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = 0.5 * flangeDiam2
        dLineStrPoints(5) = 0.5 * flangeDiam2

        'back points
        dLineStrPoints(6) = dLineStrPoints(0)
        dLineStrPoints(7) = 0.5 * flangeDiam2
        dLineStrPoints(8) = -0.5 * flangeDiam2

        dLineStrPoints(9) = dLineStrPoints(0)
        dLineStrPoints(10) = -0.5 * flangeDiam2
        dLineStrPoints(11) = -0.5 * flangeDiam2

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                        5, dLineStrPoints)
        oAxisvec.Set 1, 0, 0
        Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisvec, (0.1 * (dFace2toCenter - flangeThick2) - 0.5 * dThickness), True)
        'Set the output
        m_OutputColl.AddOutput "Objbox_", ObjBox
        Set ObjBox = Nothing
        
        'Box 2
        dLineStrPoints(0) = 0.5 * (dFace2toCenter - flangeThick2) + 0.5 * dThickness
        dLineStrPoints(1) = -0.5 * flangeDiam2
        dLineStrPoints(2) = 0.5 * flangeDiam2
        
        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = 0.5 * flangeDiam2
        dLineStrPoints(5) = 0.5 * flangeDiam2

        'back points
        dLineStrPoints(6) = dLineStrPoints(0)
        dLineStrPoints(7) = 0.5 * flangeDiam2
        dLineStrPoints(8) = -0.5 * flangeDiam2

        dLineStrPoints(9) = dLineStrPoints(0)
        dLineStrPoints(10) = -0.5 * flangeDiam2
        dLineStrPoints(11) = -0.5 * flangeDiam2

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                        5, dLineStrPoints)
        oAxisvec.Set 1, 0, 0
        Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisvec, (0.1 * (dFace2toCenter - flangeThick2) - 0.5 * dThickness), True)
        'Set the output
        m_OutputColl.AddOutput "Objbox_", ObjBox
        Set ObjBox = Nothing

    Case 965                            '3 Way Ball Valve, Y-Type
        parFacetoCenter = arrayOfInputs(8)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
        dFace3toCenter = parFacetoCenter
        
        Dim dFlangeDia1 As Double
        Dim dFlangeDia2 As Double
        Dim dFlangeDia3 As Double
        dFlangeDia1 = flangeDiam / Sqr(2)
        dFlangeDia2 = flangeDiam2 / Sqr(2)
        dFlangeDia3 = flangeDiam3 / Sqr(2)
        dThickness = 0.0625 * 0.0254 '0.0625in
        If CmpDblGreaterthan(dEllipseMajor, dFace1toCenter / 2) Then
            dEllipseMajor = 0.35 * dFace1toCenter
        End If

        Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(m_OutputColl.ResourceManager, _
            CenterPos.x, CenterPos.y + 0.5 * (dFlangeDia1 + pipeDiam) / 2, CenterPos.z, _
            0, 1, 0, _
            0, 0, dEllipseMajor, 0.5)
        oAxisvec.Set 0, 1, 0
        Set objEllTop = PlaceProjection(m_OutputColl, oEllipse, oAxisvec, 0.05 * dFace3toCenter, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEllTop
        Set oEllipse = Nothing
        Set objEllTop = Nothing

        oStPoint.Set -dFace1toCenter + flangeThick, 0, 0
        oEnPoint.Set -(0.5 * dFace1toCenter + flangeThick), 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing

        oStPoint.Set -(0.5 * dFace1toCenter + flangeThick), -0.5 * dFlangeDia1, 0.5 * dFlangeDia1
        oEnPoint.Set oStPoint.x + flangeThick - 0.5 * dThickness, 0.5 * dFlangeDia1, -0.5 * dFlangeDia1
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the output
        m_OutputColl.AddOutput "Objbox_", ObjBox
        Set ObjBox = Nothing

        oStPoint.Set -0.5 * dFace1toCenter + flangeThick, -0.5 * dFlangeDia1, 0.5 * dFlangeDia1
        oEnPoint.Set oStPoint.x - (flangeThick - 0.5 * dThickness), 0.5 * dFlangeDia1, -0.5 * dFlangeDia1
        Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the output
        m_OutputColl.AddOutput "Objbox_", ObjBox
        Set ObjBox = Nothing

        oStPoint.Set -0.5 * dFace1toCenter + flangeThick, 0, 0
        oEnPoint.Set 0, 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * (dFlangeDia1 + pipeDiam), True)
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", objCylinder
        Set objCylinder = Nothing

        'Port 2 body
        oStPoint.Set (dFace2toCenter - flangeThick2) * Cos(PI / 3), 0, -(dFace2toCenter - flangeThick2) * Sin(PI / 3)
        oEnPoint.Set (0.5 * dFace2toCenter + flangeThick2) * Cos(PI / 3), 0, -(0.5 * dFace2toCenter + flangeThick2) * Sin(PI / 3)
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam2, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
        
        'Line
        dLineStrPoints(0) = (0.5 * dFace2toCenter - 0.5 * dThickness)
        dLineStrPoints(1) = -0.5 * dFlangeDia2
        dLineStrPoints(2) = 0.5 * dFlangeDia2
        
        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = 0.5 * dFlangeDia2
        dLineStrPoints(5) = 0.5 * dFlangeDia2

        'back points
        dLineStrPoints(6) = dLineStrPoints(0)
        dLineStrPoints(7) = 0.5 * dFlangeDia2
        dLineStrPoints(8) = -0.5 * dFlangeDia2

        dLineStrPoints(9) = dLineStrPoints(0)
        dLineStrPoints(10) = -0.5 * dFlangeDia2
        dLineStrPoints(11) = -0.5 * dFlangeDia2

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                        5, dLineStrPoints)
        oAxisvec.Set -1, 0, 0
        Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisvec, (flangeThick2 - 0.5 * dThickness), True)
        Dim oTransMatrix As IJDT4x4
        Set oTransMatrix = New DT4x4
        oTransMatrix.LoadIdentity
        oAxisvec.Set 0, 1, 0
        oTransMatrix.Rotate PI / 3, oAxisvec
        ObjBox.Transform oTransMatrix
        'Set the output
        m_OutputColl.AddOutput "Objbox_", ObjBox
        Set ObjBox = Nothing
        
        'Box 2
        dLineStrPoints(0) = (0.5 * dFace2toCenter + 0.5 * dThickness)
        dLineStrPoints(1) = -0.5 * dFlangeDia2
        dLineStrPoints(2) = 0.5 * dFlangeDia2
        
        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = 0.5 * dFlangeDia2
        dLineStrPoints(5) = 0.5 * dFlangeDia2

        'back points
        dLineStrPoints(6) = dLineStrPoints(0)
        dLineStrPoints(7) = 0.5 * dFlangeDia2
        dLineStrPoints(8) = -0.5 * dFlangeDia2

        dLineStrPoints(9) = dLineStrPoints(0)
        dLineStrPoints(10) = -0.5 * dFlangeDia2
        dLineStrPoints(11) = -0.5 * dFlangeDia2

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                        5, dLineStrPoints)
        oAxisvec.Set 1, 0, 0
        Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisvec, (flangeThick2 - 0.5 * dThickness), True)
        Set oTransMatrix = New DT4x4
        oTransMatrix.LoadIdentity
        oAxisvec.Set 0, 1, 0
        oTransMatrix.Rotate PI / 3, oAxisvec
        ObjBox.Transform oTransMatrix
        'Set the output
        m_OutputColl.AddOutput "Objbox_", ObjBox
        Set ObjBox = Nothing
        
        oStPoint.Set (0.5 * dFace2toCenter - flangeThick2) * Cos(PI / 3), 0, -(0.5 * dFace2toCenter - flangeThick2) * Sin(PI / 3)
        oEnPoint.Set 0, 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * (dFlangeDia2 + pipeDiam2), True)
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", objCylinder
        Set objCylinder = Nothing

        'Port 3
        oStPoint.Set (dFace3toCenter - flangeThick3) * Sin(PI / 6), 0, (dFace3toCenter - flangeThick3) * Cos(PI / 6)
        oEnPoint.Set (0.5 * dFace3toCenter + flangeThick3) * Sin(PI / 6), 0, (0.5 * dFace3toCenter + flangeThick3) * Cos(PI / 6)
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam3, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing
        
        dLineStrPoints(0) = (0.5 * dFace3toCenter - 0.5 * dThickness)
        dLineStrPoints(1) = -0.5 * dFlangeDia3
        dLineStrPoints(2) = 0.5 * dFlangeDia3
        
        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = 0.5 * dFlangeDia3
        dLineStrPoints(5) = 0.5 * dFlangeDia3

        'back points
        dLineStrPoints(6) = dLineStrPoints(0)
        dLineStrPoints(7) = 0.5 * dFlangeDia3
        dLineStrPoints(8) = -0.5 * dFlangeDia3

        dLineStrPoints(9) = dLineStrPoints(0)
        dLineStrPoints(10) = -0.5 * dFlangeDia3
        dLineStrPoints(11) = -0.5 * dFlangeDia3

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                        5, dLineStrPoints)
        oAxisvec.Set -1, 0, 0
        Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisvec, (flangeThick3 - 0.5 * dThickness), True)
        Set oTransMatrix = New DT4x4
        oTransMatrix.LoadIdentity
        oAxisvec.Set 0, -1, 0
        oTransMatrix.Rotate PI / 3, oAxisvec
        ObjBox.Transform oTransMatrix
        'Set the output
        m_OutputColl.AddOutput "Objbox_", ObjBox
        Set ObjBox = Nothing
        
        'Box 2
        dLineStrPoints(0) = (0.5 * dFace3toCenter + 0.5 * dThickness)
        dLineStrPoints(1) = -0.5 * dFlangeDia3
        dLineStrPoints(2) = 0.5 * dFlangeDia2
        
        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = 0.5 * dFlangeDia3
        dLineStrPoints(5) = 0.5 * dFlangeDia3

        'back points
        dLineStrPoints(6) = dLineStrPoints(0)
        dLineStrPoints(7) = 0.5 * dFlangeDia3
        dLineStrPoints(8) = -0.5 * dFlangeDia3

        dLineStrPoints(9) = dLineStrPoints(0)
        dLineStrPoints(10) = -0.5 * dFlangeDia3
        dLineStrPoints(11) = -0.5 * dFlangeDia3

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                        5, dLineStrPoints)
        oAxisvec.Set 1, 0, 0
        Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisvec, (flangeThick3 - 0.5 * dThickness), True)
        Set oTransMatrix = New DT4x4
        oTransMatrix.LoadIdentity
        oAxisvec.Set 0, -1, 0
        oTransMatrix.Rotate PI / 3, oAxisvec
        ObjBox.Transform oTransMatrix
        'Set the output
        m_OutputColl.AddOutput "Objbox_", ObjBox
        Set ObjBox = Nothing
        
        oStPoint.Set (0.5 * dFace3toCenter - flangeThick3) * Sin(PI / 6), 0, (0.5 * dFace3toCenter - flangeThick3) * Cos(PI / 6)
        oEnPoint.Set 0, 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * (dFlangeDia3 + pipeDiam3), True)
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder_", objCylinder
        Set objCylinder = Nothing

    Case Else
        GoTo ErrorLabel

    End Select
        
    If lPartdatabasis <= 1 Or lPartdatabasis = 10 Or _
             lPartdatabasis = 11 Or lPartdatabasis = 13 Then
        'Insert your code for output 1(Ball of Valve Body)
        'Ball diameter is considered to be half of the maximum face to center distance
        BallDiameter = 0.5 * ReturnMax3(dFace1toCenter, dFace2toCenter, dFace3toCenter)
        Set ObjBodyBall = PlaceSphere(m_OutputColl, CenterPos, BallDiameter / 2)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyBall
        Set ObjBodyBall = Nothing
    End If
      
    If lPartdatabasis <= 1 Or lPartdatabasis = 10 Or _
       lPartdatabasis = 11 Or lPartdatabasis = 13 Then

        'Insert your code for output 2(Cone Port1 of Body)
        oStPoint.Set CenterPos.x - (dFace1toCenter - flangeThick), _
                     CenterPos.y, _
                     CenterPos.z
        'oEnPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
        'PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
        'value removes the 'Invalid property value' failure.
        oEnPoint.Set CenterPos.x, _
                     CenterPos.y, _
                     CenterPos.z
        Set ObjBodyCone1 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam / 2, 0, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
        Set ObjBodyCone1 = Nothing

        'Insert your code for output 3(Cone Port2 of Body)
        oStPoint.Set CenterPos.x + dFace2toCenter - flangeThick2, _
                     CenterPos.y, _
                     CenterPos.z
        'oEnPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
        'PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
        'value removes the 'Invalid property value' failure.
        oEnPoint.Set CenterPos.x, _
                     CenterPos.y, _
                     CenterPos.z
        Set ObjBodyCone2 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam2 / 2, 0, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
        Set ObjBodyCone2 = Nothing

        'Insert your code for output 4(Cone Port3 of Body)
        oStPoint.Set CenterPos.x, _
                     CenterPos.y + (dFace3toCenter - flangeThick3), _
                     CenterPos.z
        'Updated to fix problem reported by Franz Drews SR 1-29654801
        'oEnPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
        'PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
        'value removes the 'Invalid property value' failure.
        oEnPoint.Set CenterPos.x, _
                     CenterPos.y, _
                     CenterPos.z
        Set ObjBodyCone3 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam3 / 2, 0, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone3
        Set ObjBodyCone3 = Nothing

    End If
        Set oGeomFactory = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oAxisvec = Nothing
        Set oTransMatrix = Nothing
        
    'Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    If lPartdatabasis <= 1 Or lPartdatabasis = 10 Or _
       lPartdatabasis = 11 Or lPartdatabasis = 13 Or _
       lPartdatabasis = 963 Or lPartdatabasis = 964 Or lPartdatabasis = 965 Then

        oPlacePoint.Set CenterPos.x - dFace1toCenter - sptOffset + depth, _
                        CenterPos.y, _
                        CenterPos.z
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    End If

    'Place Nozzle 2
    If lPartdatabasis <= 1 Or lPartdatabasis = 10 Or _
       lPartdatabasis = 11 Or lPartdatabasis = 13 Or _
       lPartdatabasis = 963 Or lPartdatabasis = 964 Then

        oPlacePoint.Set CenterPos.x + dFace2toCenter + sptOffset2 - depth2, _
                        CenterPos.y, _
                        CenterPos.z
        oDir.Set 1, 0, 0
    ElseIf lPartdatabasis = 965 Then
        oPlacePoint.Set (CenterPos.x + dFace2toCenter + sptOffset2 - depth2) * Cos(PI / 3), _
                        CenterPos.y, _
                        -(CenterPos.z + dFace2toCenter + sptOffset2 - depth2) * Sin(PI / 3)
        oDir.Set Cos(PI / 3), 0, -Sin(PI / 3)
    End If
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    'Place Nozzle 3
    If lPartdatabasis <= 1 Or lPartdatabasis = 10 Or _
       lPartdatabasis = 11 Or lPartdatabasis = 13 Then
        oPlacePoint.Set CenterPos.x, _
                        CenterPos.y + (dFace3toCenter + sptOffset3 - depth3), _
                        CenterPos.z
        oDir.Set 0, 1, 0
    ElseIf lPartdatabasis = 963 Or lPartdatabasis = 964 Then
        oPlacePoint.Set CenterPos.x, _
                        CenterPos.y, _
                        CenterPos.z + dFace3toCenter + sptOffset3 - depth3
        oDir.Set 0, 0, 1
    ElseIf lPartdatabasis = 965 Then
        oPlacePoint.Set (CenterPos.x + dFace3toCenter + sptOffset3 - depth3) * Cos(PI / 3), _
                        CenterPos.y, _
                        (CenterPos.z + dFace3toCenter + sptOffset3 - depth3) * Sin(PI / 3)
        oDir.Set Sin(PI / 6), 0, Cos(PI / 6)
    End If
    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    'Insert your code for output 10 (Valve Operator)
    Dim oSymbolHelper As IJSymbolGeometryHelper
    Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
    oSymbolHelper.OutputCollection = m_OutputColl

    On Error Resume Next
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector

    oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
    oDirY.Set 0, 1, 0
    oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)

    Set oPipeComponent = oPartFclt
    On Error GoTo ErrorLabel
    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc   As IJPartOcc
    If Not oPipeComponent Is Nothing Then
'       Note: oPipeComponent.GetValveOperatorPart method call may fail in following two cases:
'        1. If valve operator catalog part number is not defined (to check for E_NOOPERATORDEFINED) or
'        2. If valve operator part is not found (to check for E_NOOPERATORPARTFOUND)
'        symbol should raise error for case 2.
    On Error GoTo HandleOperatorCodeLabel
        Set oOperatorPart = oPipeComponent.GetValveOperatorPart
        If Not oOperatorPart Is Nothing Then
            Dim OpOrigin As IJDPosition
            Set OpOrigin = New DPosition
            OpOrigin.Set 0, 0, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        End If
    End If
HandleOperatorCodeLabel:
'   The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
        If Err.Number = E_NOOPERATORPARTFOUND Then   'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If
    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
    Set CenterPos = Nothing
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing
    Set oLineString = Nothing
    Set OpOrigin = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
