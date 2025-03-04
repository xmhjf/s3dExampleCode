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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Thursday, Jun 13 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No. is: V18, Page No. D-17 PDS Piping Component Data Reference Guide.
' The Symbol is made of eleven output objects. Nine are physical aspect outputs and Two are insulation aspect outputs.
' The physical outputs are made using 'PlaceCone' and 'CreateNozzle' functions.
'The Insulation outputs are made using 'PlaceCylinder' functions.
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------         -----           ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav                Defined new input parameter HandwheelAngle and modified valve operator direction vectors.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   08.SEP.2006     KKC     DI-95670        Replace names with initials in all revision history sheets and symbols
'   17.JAN.2008     PK      CR-133006       Enhanced the existing 3 way plug valve to support dimensional basis for
'                                               (i)Face-to-center dimension basis
'                                               (ii)Face-to-Face dimension and Face-to-Center dimension for port 3
'                                               (iii)Asymmetrical Face-to-Center dimension basis
'                                           (Source:JIS F7381H.pdf)
'   21.Feb.2008     RRK     CR-33401        Enhanced the symbol to support the part data basis value-Plug valve, 3-way, partial jacket, specified by face-to-face dimension
'                                           and face-to-center dimension for branch (822) and Plug valve, 3-way, full jacket, oversized flanges, specified by face-to-face dimension and face-to-center dimension for branch(823)
'   05.Jun.2008     RRK      TR-143567,
'                               143575      Changed the order in which nozzles are created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl

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

    Dim pipeDiam4 As Double
    Dim flangeThick4 As Double
    Dim sptOffset4 As Double
    Dim flangeDiam4 As Double
    Dim depth4 As Double

    Dim pipeDiam5 As Double
    Dim flangeThick5 As Double
    Dim sptOffset5 As Double
    Dim flangeDiam5 As Double
    Dim depth5 As Double

    Dim iOutput As Double
    Dim ObjLeftCone As Object
    Dim ObjUpperCone As Object
    Dim ObjLowerCone As Object
    Dim ObjRightCone As Object
    Dim ObjTopCone As Object
    Dim ObjInsulCylinder1 As Object
    Dim ObjInsulCylinder2 As Object

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double
    Dim parInletYOffset As Double
    Dim parOutletYOffset As Double
    Dim parInletOutletZOffset As Double
    Dim parInletXOffset As Double
    Dim parInletZOffset As Double
    Dim parOutletZOffset As Double

    Dim parFacetoFace As Double
    Dim parFacetoCenter As Double

    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dFace3toCenter As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFace1toCenter = arrayOfInputs(2)     'A
    'parFace2toCenter = arrayOfInputs(3)     'B
    'parFace3toCenter = arrayOfInputs(4)     'C
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    'parFacetoFace  =arrayOfInputs(7)
    'parFacetoCenter = arrayOfInputs(8)
    'parInletYOffset = arrayOfInputs(9)
    'parOutletYOffset = arrayOfInputs(10)
    'parInletOutletZOffset = arrayOfInputs(11)
    'parInletXOffset = arrayOfInputs(12)
    'parInletZOffset = arrayOfInputs(13)
    'parOutletZOffset = arrayOfInputs(14)

    iOutput = 0

    Dim oPipeComponent As IJDPipeComponent
    Dim lPartdatabasis As Long

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    Set oPipeComponent = oPartFclt
    lPartdatabasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, sptOffset3, depth3

    Select Case lPartdatabasis
    Case Is <= 1, 13 'Asymmetrical Face-to-Center dimension basis
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
        parFace3toCenter = arrayOfInputs(4)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        dFace3toCenter = parFace3toCenter
        If CmpDblEqual(dFace1toCenter, 0) Then
            dFace1toCenter = dFace2toCenter
        ElseIf CmpDblEqual(dFace2toCenter, 0) Then
            dFace2toCenter = dFace1toCenter
        End If

    Case 10  'Face-to-center dimension basis
        parFacetoCenter = arrayOfInputs(8)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
        dFace3toCenter = parFacetoCenter

    Case 11 'Face-to-Face dimension and Face-to-Center dimension for port 3
        parFace3toCenter = arrayOfInputs(4)
        parFacetoFace = arrayOfInputs(7)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
        dFace3toCenter = parFace3toCenter

    Case 822 'Plug valve, 3-way, partial jacket, specified by face-to-face dimension and face-to-center dimension for branch
        parFace3toCenter = arrayOfInputs(4)
        parFacetoFace = arrayOfInputs(7)
        parInletYOffset = arrayOfInputs(9)
        parOutletYOffset = arrayOfInputs(10)
        parInletOutletZOffset = arrayOfInputs(11)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
        dFace3toCenter = parFace3toCenter

        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam4, sptOffset4, depth4
        RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, flangeThick5, flangeDiam5, sptOffset5, depth5

    Case 823 'Plug valve, 3-way, full jacket, oversized flanges, specified by face-to-face dimension and face-to-center dimension for branch
        parFace3toCenter = arrayOfInputs(4)
        parFacetoFace = arrayOfInputs(7)
        parInletYOffset = arrayOfInputs(9)
        parOutletYOffset = arrayOfInputs(10)
        parInletXOffset = arrayOfInputs(12)
        parInletZOffset = arrayOfInputs(13)
        parOutletZOffset = arrayOfInputs(14)

        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
        dFace3toCenter = parFace3toCenter

        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam4, sptOffset4, depth4
        RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, flangeThick5, flangeDiam5, sptOffset5, depth5

    Case Else
        GoTo ErrorLabel:
    End Select

    'Creating outputs based on part data basis values
    Select Case lPartdatabasis
    Case Is <= 1, 10, 11, 13
        'Insert your code for output 1(Left Cone)
        Dim BaseCenPoint As New AutoMath.DPosition
        Dim TopCenPoint As New AutoMath.DPosition

        BaseCenPoint.Set -dFace1toCenter + flangeThick, 0#, 0#

        'PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
        'value removes the 'Invalid property value' failure.

        TopCenPoint.Set 0#, 0#, 0
        Set ObjLeftCone = PlaceCone(m_OutputColl, BaseCenPoint, TopCenPoint, pipeDiam / 2, 0, True)

        'Set the output
        m_OutputColl.AddOutput "LeftCone", ObjLeftCone
        Set ObjLeftCone = Nothing
        Set BaseCenPoint = Nothing
        Set TopCenPoint = Nothing

        'Insert your code for output 2(Upper Cone)
        Dim UpperConeBasePt As New AutoMath.DPosition
        Dim UpperConeTopPt As New AutoMath.DPosition
        Dim coneDia As Double
        Dim coneHt As Double

        coneDia = dFace1toCenter * 0.5

        If CmpDblGreaterthan(pipeDiam2, pipeDiam) Then
            coneHt = pipeDiam2 * 0.5
        Else
            coneHt = pipeDiam * 0.5
        End If

        UpperConeBasePt.Set 0#, 0#, 0#
        UpperConeTopPt.Set 0#, coneHt, 0#

        Set ObjUpperCone = PlaceCone(m_OutputColl, UpperConeBasePt, UpperConeTopPt, coneDia / 2, 0, True)

        'Set the output
        m_OutputColl.AddOutput "UpperCone", ObjUpperCone
        Set ObjUpperCone = Nothing
        Set UpperConeBasePt = Nothing
        Set UpperConeTopPt = Nothing

        'Insert your code for output 3(Lower Cone)
        Dim LowerConeBasePt As New AutoMath.DPosition
        Dim LowerConeTopPt As New AutoMath.DPosition

        LowerConeBasePt.Set 0#, 0#, 0#
        LowerConeTopPt.Set 0#, -coneHt, 0#

        Set ObjLowerCone = PlaceCone(m_OutputColl, LowerConeBasePt, LowerConeTopPt, coneDia / 2, 0, True)

        'Set the output
        m_OutputColl.AddOutput "LowerCone", ObjLowerCone
        Set ObjLowerCone = Nothing
        Set LowerConeBasePt = Nothing
        Set LowerConeTopPt = Nothing

        'Insert your code for output 4(Right Cone)
        Dim RightConeBasePt As New AutoMath.DPosition
        Dim RightConeTopPt As New AutoMath.DPosition

        RightConeBasePt.Set dFace2toCenter - flangeThick2, 0#, 0#
        'PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
        'value removes the 'Invalid property value' failure.
        RightConeTopPt.Set 0#, 0#, 0

        Set ObjRightCone = PlaceCone(m_OutputColl, RightConeBasePt, RightConeTopPt, pipeDiam2 / 2, 0, True)

        'Set the output
        m_OutputColl.AddOutput "RightCone", ObjRightCone
        Set ObjRightCone = Nothing
        Set RightConeBasePt = Nothing
        Set RightConeTopPt = Nothing

        'Insert your code for output 5(Top Cone)
        Dim TopConeBasePt As New AutoMath.DPosition
        Dim TopConeTopPt As New AutoMath.DPosition

        TopConeBasePt.Set 0, dFace3toCenter - flangeThick3, 0#
        'PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
        'value removes the 'Invalid property value' failure.
        TopConeTopPt.Set 0#, 0#, 0

        Set ObjTopCone = PlaceCone(m_OutputColl, TopConeBasePt, TopConeTopPt, pipeDiam3 / 2, 0, True)

        'Set the output
        m_OutputColl.AddOutput "TopCone", ObjTopCone
        Set ObjTopCone = Nothing
        Set TopConeBasePt = Nothing
        Set TopConeTopPt = Nothing
        
        'Place Nozzle 1
        oPlacePoint.Set -dFace1toCenter - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
    
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        'Place Nozzle 2
        oPlacePoint.Set dFace2toCenter + sptOffset2 - depth2, 0, 0
        oDir.Set 1, 0, 0
    
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        'Place Nozzle 3
        oPlacePoint.Set 0, dFace3toCenter + sptOffset3 - depth3, 0
        oDir.Set 0, 1, 0
    
        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        Set oPlacePoint = Nothing
        Set oDir = Nothing

    Case 822
        'Place Jacket Body
        Dim objJacketBody As Object
        Dim oStPoint As Object
        Dim oEnPoint As Object
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition

        Dim dJacketBoxLength As Double
        Dim dJacketBoxWidth As Double
        Dim dJacketBoxTopPos As Double
        Dim dJacketBoxBottomPos As Double

        Dim dMaxJacketBoxLength As Double
        dMaxJacketBoxLength = (parFacetoFace - flangeThick - flangeThick2)

        Dim dMinJacketBoxLength As Double
        dMinJacketBoxLength = pipeDiam3

        'Taking the average of maximum and minimum lengths as jacket box length
        dJacketBoxLength = (dMaxJacketBoxLength + dMinJacketBoxLength) / 2

        'Setting jacket box width equal to twice of the Z-offset of inlet minus flange thickness of jacket inlet
        dJacketBoxWidth = 2 * (parInletOutletZOffset - flangeThick4)

        'Check for case where flange thick ness of jacket inlet  is zero
        If CmpDblEqual(flangeThick4, 0) Then dJacketBoxWidth = 1.8 * parInletOutletZOffset

        'Setting the jacket box top position
        If CmpDblGreaterthan(pipeDiam4, flangeDiam4) Then
            dJacketBoxTopPos = 1.01 * (parInletYOffset + (pipeDiam4 / 2))
        Else
            dJacketBoxTopPos = 1.01 * (parInletYOffset + (flangeDiam4 / 2))
        End If

        'Setting the jacket box bottom position
        dJacketBoxBottomPos = 0.95 * (parFace3toCenter - flangeThick3)

        oStPoint.Set -dJacketBoxLength / 2, -dJacketBoxBottomPos, -dJacketBoxWidth / 2
        oEnPoint.Set dJacketBoxLength / 2, dJacketBoxTopPos, dJacketBoxWidth / 2

        Set objJacketBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        ' Set the output
        m_OutputColl.AddOutput "JacketBody", objJacketBody
        Set objJacketBody = Nothing

        'Place Jacket Top
        Dim objTopBox As Object

        Dim dTopBoxLength As Double
        Dim dTopBoxWidth As Double
        Dim dBoxTopPos As Double
        Dim dBoxBottomPos As Double

        'Assumptions for jacket top box
        dTopBoxLength = dJacketBoxLength
        dTopBoxWidth = dJacketBoxWidth
        dBoxBottomPos = dJacketBoxTopPos + 0.01 * pipeDiam
        dBoxTopPos = dBoxBottomPos + 0.2 * pipeDiam

        'Check to see that the box top position exceeds pipe outside diameter
        If CmpDblLessThan(dBoxTopPos, 0.5 * pipeDiam) Then
            dBoxTopPos = 0.55 * pipeDiam
        End If

        oStPoint.Set -dTopBoxLength / 2, dBoxBottomPos, -dTopBoxWidth / 2
        oEnPoint.Set dTopBoxLength / 2, dBoxTopPos, dTopBoxWidth / 2

        Set objTopBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        ' Set the output
        m_OutputColl.AddOutput "TopBox", objTopBox
        Set objTopBox = Nothing

        'Place valve body(cylinder 1) joining port 1 and port 2
        Dim objCylinder1 As Object

        Dim dCylinderDiam As Double
        dCylinderDiam = pipeDiam

        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set dFace2toCenter, 0, 0

        Set objCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylinderDiam, True)

        ' Set the output
        m_OutputColl.AddOutput "ValveBodyCylinder1", objCylinder1
        Set objCylinder1 = Nothing

        'Place valve body(cylinder2) from port 3
        Dim objCylinder2 As Object

        dCylinderDiam = pipeDiam3

        oStPoint.Set 0, -dJacketBoxBottomPos, 0
        oEnPoint.Set 0, -dFace3toCenter, 0

        'Check to see that creation of cylinder doesn't fail
        If CmpDblEqual(oStPoint.x, oEnPoint.x) Then
            oEnPoint.x = oStPoint.x - 0.0001
        End If

        Set objCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylinderDiam, True)

        ' Set the output
        m_OutputColl.AddOutput "ValveBodyCylinder2", objCylinder2
        Set objCylinder2 = Nothing

        'Place Nozzle 1
        oPlacePoint.Set -dFace1toCenter - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
    
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        'Place Nozzle 2
        oPlacePoint.Set dFace2toCenter + sptOffset2 - depth2, 0, 0
        oDir.Set 1, 0, 0
    
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        'Place Nozzle 3

        oPlacePoint.Set 0, -dFace3toCenter - sptOffset3 + depth3, 0
        oDir.Set 0, -1, 0
    
        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        ' Place Nozzle 4(Jacket Inlet)
        oPlacePoint.Set 0, parInletYOffset, parInletOutletZOffset + sptOffset4 - depth4
        oDir.Set 0, 0, 1

        Dim dNozzleLength As Double
        dNozzleLength = parInletOutletZOffset - (dJacketBoxWidth / 2)

        Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)
        ' Set the output
        m_OutputColl.AddOutput "JacketInlet", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 5 (Jacket Outlet)
        oPlacePoint.Set 0, -parOutletYOffset, parInletOutletZOffset + sptOffset5 - depth5
        oDir.Set 0, 0, 1
        dNozzleLength = parInletOutletZOffset - (dJacketBoxWidth / 2)

        Set objNozzle = CreateNozzleWithLength(5, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)
        ' Set the output
        m_OutputColl.AddOutput "JacketOutlet", objNozzle
        Set objNozzle = Nothing
        Set oPlacePoint = Nothing
        Set oDir = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing

    Case 823

        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition

        'Place valve body(cylinder 1) joining port 1 and port 2
        'Computing greater of pipe and flange diameters for jacket piping ports
        Dim dMaxDia5 As Double
        dMaxDia5 = GreaterDim(pipeDiam5, flangeDiam5)

        Dim dMaxDia4 As Double
        dMaxDia4 = GreaterDim(pipeDiam4, flangeDiam4)
        
        'Computing the the maximum cylinder diameter 1
        Dim dMaxCylDia1 As Double
        If CmpDblEqual(flangeThick4, 0) Then
            dMaxCylDia1 = 2 * Sqr((parOutletYOffset - dMaxDia5 / 2) ^ 2 + (1.8 * parOutletZOffset) ^ 2)
        Else
            dMaxCylDia1 = 2 * Sqr((parOutletYOffset - dMaxDia5 / 2) ^ 2 + (parOutletZOffset - flangeThick4) ^ 2)
        End If
        
        'Computing the greater of pipe and flange diameters for port 1
        Dim dMaxCylDia2 As Double
        dMaxCylDia2 = GreaterDim(flangeDiam, pipeDiam)
        
        'Computing the limiting cylinder diameter
        Dim dMaxCylDia As Double
        dMaxCylDia = SmallerDim(dMaxCylDia1, 0.8 * dMaxCylDia2)
            
        'Computing the cylinder diameter
        dCylinderDiam = 2 * Sqr(parInletZOffset ^ 2 + (parInletYOffset + dMaxDia4 / 2) ^ 2)
        
        'Check to see that cylinder diameter doesn't exceed limiting cylinder diameter
        If CmpDblGreaterthan(dCylinderDiam, dMaxCylDia) Then
            dCylinderDiam = 0.9 * dMaxCylDia
        End If

        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set dFace2toCenter, 0, 0

        Set objCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylinderDiam, True)

        ' Set the output
        m_OutputColl.AddOutput "ValveBodyCylinder1", objCylinder1
        Set objCylinder1 = Nothing

        'Place valve body(cylinder2) from port 3 to center
        Dim dCylinderDiam2 As Double
        
        'Setting the cylinder diameter 2 based on flange thickness
        If CmpDblEqual(flangeThick5, 0) Then
            dCylinderDiam2 = 1.8 * parOutletZOffset
        Else
            dCylinderDiam2 = 2 * (parOutletZOffset - flangeThick5)
        End If
        
        'Calculating the limiting cylinder 2 diameter
        dMaxCylDia1 = parFacetoFace - flangeThick - flangeThick2
        dMaxCylDia2 = GreaterDim(flangeDiam, pipeDiam)
        dMaxCylDia = SmallerDim(dMaxCylDia1, 0.8 * dMaxCylDia2)
        
        'Check to see that cylinder 2 diameter doesn't exceed limiting diameter
        If CmpDblGreaterthan(dCylinderDiam2, dMaxCylDia) Then
            dCylinderDiam2 = 0.9 * dMaxCylDia
        End If
        
        'Setting the start point of the cylinder
        If CmpDblGreaterthan(dCylinderDiam2, dCylinderDiam) Then
            oStPoint.Set 0, -dCylinderDiam / 2, 0
        Else
            oStPoint.Set 0, 0, 0
        End If

        oEnPoint.Set 0, -dFace3toCenter, 0

        Set objCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylinderDiam2, True)

        ' Set the output
        m_OutputColl.AddOutput "ValveBodyCylinder2", objCylinder2
        Set objCylinder2 = Nothing

        'Creating cone when cylinder diameter 2 is greater than cylinder diameter 1
        If CmpDblGreaterthan(dCylinderDiam2, dCylinderDiam) Then
            Dim objCone As Object
            oStPoint.Set 0, -dCylinderDiam / 2, 0
            oEnPoint.Set 0, 0, 0

            Set objCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, dCylinderDiam2 / 2, 0, True)
            ' Set the output
            m_OutputColl.AddOutput "Cone", objCone
            Set objCone = Nothing
        End If

        'Creating Box
        'Assumptions for jacket top box
        dTopBoxLength = 0.5 * (parFacetoFace - flangeThick - flangeThick2)
        dTopBoxWidth = 0.8 * dCylinderDiam
        dBoxBottomPos = 0.45 * dCylinderDiam
        dBoxTopPos = 0.55 * dCylinderDiam

        oStPoint.Set -dTopBoxLength / 2, dBoxBottomPos, -dTopBoxWidth / 2
        oEnPoint.Set dTopBoxLength / 2, dBoxTopPos, dTopBoxWidth / 2

        Set objTopBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        ' Set the output
        m_OutputColl.AddOutput "TopBox", objTopBox
        Set objTopBox = Nothing

        'Place Nozzle 1
        oPlacePoint.Set -dFace1toCenter - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
    
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        'Place Nozzle 2
        oPlacePoint.Set dFace2toCenter + sptOffset2 - depth2, 0, 0
        oDir.Set 1, 0, 0
    
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        'Place Nozzle 3

        oPlacePoint.Set 0, -dFace3toCenter - sptOffset3 + depth3, 0
        oDir.Set 0, -1, 0
    
        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        ' Place Nozzle 4(Jacket inlet)
        oPlacePoint.Set parInletXOffset, parInletYOffset, parInletZOffset + sptOffset4 - depth4
        oDir.Set 0, 0, 1

        dNozzleLength = parInletZOffset

        Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)
        ' Set the output
        m_OutputColl.AddOutput "JacketInlet", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 5(Jacket outlet)
        oPlacePoint.Set 0, -parOutletYOffset, parOutletZOffset + sptOffset5 - depth5
        oDir.Set 0, 0, 1
        dNozzleLength = parOutletZOffset

        Set objNozzle = CreateNozzleWithLength(5, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)
        ' Set the output
        m_OutputColl.AddOutput "JacketOutlet", objNozzle
        Set objNozzle = Nothing
        Set oPlacePoint = Nothing
        Set oDir = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
    End Select

    'Insert your code for output 11 (Valve Operator)
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

    'Setting the operator orientation based on part data basis values
    If lPartdatabasis = 822 Or lPartdatabasis = 823 Then
        oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
        oDirY.Set 0, 1, 0
        oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)
    Else
        oDirX.Set Cos(parHandwheelAngle), Sin(parHandwheelAngle), 0
        oDirY.Set 0, 0, 1
        oDirZ.Set -Sin(parHandwheelAngle), Cos(parHandwheelAngle), 0
    End If

    Set oPipeComponent = oPartFclt
    On Error GoTo ErrorLabel
    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc As IJPartOcc

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
    If Err.Number = E_NOOPERATORPARTFOUND Then    'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If

    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next

End Sub

Private Function GreaterDim(ByVal Dim1 As Double, ByVal Dim2 As Double) As Double
        If CmpDblGreaterthanOrEqualTo(Dim1, Dim2) Then
            GreaterDim = Dim1
        Else
            GreaterDim = Dim2
        End If
    
End Function
