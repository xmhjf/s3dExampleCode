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
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav                Defined new input parameter HandwheelAngle.
'   08.SEP.2006     KKC     DI-95670        Replace names with initials in all revision history sheets and symbols
'   17.JAN.2008     PK      CR-133006       Enhanced the existing 3 way plug valve to support dimensional basis for
'                                               (i)Face-to-center dimension basis
'                                               (ii)Face-to-Face dimension and Face-to-Center dimension for port 3
'                                               (iii)Asymmetrical Face-to-Center dimension basis
'                                           (Source:JIS F7381H.pdf)
'   21.Feb.2008     RRK          CR-33401   Enhanced the symbol to support the part data basis value-Plug valve, 3-way, partial jacket, specified by face-to-face dimension
'                                           and face-to-center dimension for branch (822) and Plug valve, 3-way, full jacket, oversized flanges, specified by face-to-face dimension and face-to-center dimension for branch(823)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

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

    Dim pipeDiam5        As Double
    Dim flangeThick5     As Double
    Dim sptOffset5       As Double
    Dim flangeDiam5      As Double
    Dim depth5           As Double

    Dim iOutput                As Double
    Dim ObjLeftCone            As Object
    Dim ObjUpperCone           As Object
    Dim ObjLowerCone           As Object
    Dim ObjRightCone           As Object
    Dim ObjTopCone             As Object
    Dim ObjInsulCylinder1      As Object
    Dim ObjInsulCylinder2      As Object
    Dim parFace1toCenter       As Double
    Dim parFace2toCenter       As Double
    Dim parFace3toCenter       As Double
    Dim parInsulationThickness As Double
    Dim parInletYOffset        As Double
    Dim parOutletYOffset       As Double
    Dim parInletOutletZOffset  As Double
    Dim parInletXOffset        As Double
    Dim parInletZOffset        As Double
    Dim parOutletZOffset       As Double

    Dim parHandwheelAngle      As Double

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

    Dim oPipeComponent          As IJDPipeComponent
    Dim lPartdatabasis          As Long

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

    Case 10 'Face-to-center dimension basis
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
    
    Case 823                                'Plug valve, 3-way, full jacket, oversized flanges, specified by face-to-face dimension and face-to-center dimension for branch
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

    iOutput = 0
    'Insert your code for output 6(Insul Cylinder1)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDiameter As Double
    Dim InsulationDiameter1 As Double
    Dim InsulationDiameter2 As Double

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
    
    Dim InsulationDiameter3 As Double
    If CmpDblGreaterthan(flangeDiam3, pipeDiam3) Then
        InsulationDiameter3 = flangeDiam3 + parInsulationThickness * 2
    Else
        InsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
    End If
    
    Select Case lPartdatabasis
    Case Is <= 1, 10, 11, 13 'Setting end point for part data basis values 5, 10 and 13
        'Place Insulation Cylinder 1
        stPoint.Set -dFace1toCenter, 0, 0
        enPoint.Set dFace2toCenter, 0, 0
        Set ObjInsulCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder1
        Set ObjInsulCylinder1 = Nothing
        
        'Place Insulation Cylinder 2
        stPoint.Set 0, 0, 0
        enPoint.Set 0, dFace3toCenter, 0
        
        Set ObjInsulCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter3, True)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder2
        Set ObjInsulCylinder2 = Nothing
        
    Case 822
        Dim dJacketBoxLength As Double

        Dim dMaxJacketBoxLength As Double
        dMaxJacketBoxLength = (parFacetoFace - flangeThick - flangeThick2)

        Dim dMinJacketBoxLength As Double
        dMinJacketBoxLength = pipeDiam3

        'Taking the average of maximum and minimum lengths as jacket box length
        dJacketBoxLength = (dMaxJacketBoxLength + dMinJacketBoxLength) / 2
        
        'Place Insulation Cylinder 1
        stPoint.Set -dFace1toCenter, 0, 0
        enPoint.Set -dJacketBoxLength / 2, 0, 0

        Set ObjInsulCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter1, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder1
        Set ObjInsulCylinder1 = Nothing
        
        'Place Insulation Cylinder 2
        'Setting the jacket box bottom position
        Dim dJacketBoxBottomPos As Double
        dJacketBoxBottomPos = 0.95 * (parFace3toCenter - flangeThick3)
        stPoint.Set 0, -dJacketBoxBottomPos, 0
        enPoint.Set 0, -dFace3toCenter, 0

        'Check to see that creation of cylinder doesn't fail
        If CmpDblEqual(stPoint.x, enPoint.x) Then
            enPoint.x = stPoint.x - 0.0001
        End If
        
        Set ObjInsulCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter3, True)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder2
        Set ObjInsulCylinder2 = Nothing
        
        ''Place Insulation Cylinder 3
        Dim ObjInsulCylinder3 As Object
        stPoint.Set dJacketBoxLength / 2, 0, 0
        enPoint.Set dFace2toCenter, 0, 0
        Set ObjInsulCylinder3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter2, True)

        ' Set the output
        m_OutputColl.AddOutput "InsulationCyl3", ObjInsulCylinder3
        Set ObjInsulCylinder3 = Nothing

        'Place Jacket Body Insulation
        Dim objInsJacketBox As Object

        Dim dJacketBoxWidth As Double
        Dim dJacketBoxTopPos As Double

        'Setting jacket box width equal to twice of the Z-offset of inlet minus flange thickness of jacket inlet
        dJacketBoxWidth = 2 * (parInletOutletZOffset - flangeThick4)

        'Check for case where flange thickness of jacket inlet  is zero
        If CmpDblEqual(flangeThick4, 0) Then dJacketBoxWidth = 1.8 * parInletOutletZOffset

        'Setting the jacket box top position
        If CmpDblGreaterthan(pipeDiam4, flangeDiam4) Then
            dJacketBoxTopPos = 1.01 * (parInletYOffset + (pipeDiam4 / 2))
        Else
            dJacketBoxTopPos = 1.01 * (parInletYOffset + (flangeDiam4 / 2))
        End If

        'Setting the jacket box bottom position
        dJacketBoxBottomPos = 0.95 * (parFace3toCenter - flangeThick3)

        Dim dTopBoxLength As Double
        Dim dTopBoxWidth  As Double
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

        Dim dInsJacketBoxLength As Double
        dInsJacketBoxLength = dJacketBoxLength + 2 * parInsulationThickness

        If CmpDblGreaterthan(dInsJacketBoxLength, parFacetoFace) Then
            dInsJacketBoxLength = parFacetoFace
        End If

        stPoint.Set -dInsJacketBoxLength / 2, -dJacketBoxBottomPos, -dJacketBoxWidth / 2 - parInsulationThickness
        enPoint.Set dInsJacketBoxLength / 2, dBoxTopPos + parInsulationThickness, dJacketBoxWidth / 2 + parInsulationThickness

        Set objInsJacketBox = PlaceBox(m_OutputColl, stPoint, enPoint)

    ' Set the output
        m_OutputColl.AddOutput "InsJacketBody", objInsJacketBox
        Set objInsJacketBox = Nothing
    
    Case 823
    'Place insulation for valve body(cylinder 1) joining port 1 and port 2
        Dim dCylinderDiam As Double
        
        'Computing greater of pipe and flange diameters for jacket piping ports
        Dim dMaxDia5 As Double
        dMaxDia5 = GreaterDim(pipeDiam5, flangeDiam5)
        
        Dim dMaxDia4 As Double
        dMaxDia4 = GreaterDim(pipeDiam4, flangeDiam4)
        
        'Computing the the maximum cylinder diameter 1 based on flange thickness
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
        
        stPoint.Set -dFace1toCenter, 0, 0
        enPoint.Set dFace2toCenter, 0, 0
        
        Set ObjInsulCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dCylinderDiam + 2 * parInsulationThickness, True)
    
    ' Set the output
        m_OutputColl.AddOutput "InsValveBodyCylinder1", ObjInsulCylinder1
        Set ObjInsulCylinder1 = Nothing
        
        'Place valve body(cylinder2) from port 3
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
            stPoint.Set 0, -dCylinderDiam / 2, 0
        Else
            stPoint.Set 0, 0, 0
        End If
        
        enPoint.Set 0, -dFace3toCenter, 0
        
        Dim dInsulCyl2Dia As Double
        dInsulCyl2Dia = dCylinderDiam2 + 2 * parInsulationThickness
        
        If CmpDblGreaterthan(dInsulCyl2Dia, parFacetoFace) Then dInsulCyl2Dia = parFacetoFace
        
        Set ObjInsulCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulCyl2Dia, True)
    
    ' Set the output
        m_OutputColl.AddOutput "ValveBodyCylinder2", ObjInsulCylinder2
        Set ObjInsulCylinder2 = Nothing
        
        'Creating cone when cylinder diameter 2 is greater than cylinder diameter 1
        If CmpDblGreaterthan(dCylinderDiam2, dCylinderDiam) Then
            Dim objInsCone As Object
            stPoint.Set 0, -dCylinderDiam / 2, 0
            enPoint.Set 0, 0, 0
            
            Dim dInsConeDia As Double
            dInsConeDia = dCylinderDiam2 + 2 * parInsulationThickness
            
            If CmpDblGreaterthan(dInsConeDia, parFacetoFace) Then dInsConeDia = parFacetoFace
            
            Set objInsCone = PlaceCone(m_OutputColl, stPoint, enPoint, dInsConeDia / 2, 0, True)
            ' Set the output
            m_OutputColl.AddOutput "InsCone", objInsCone
            Set objInsCone = Nothing
        End If
        
        'Creating Insulation for Box
        'Assumptions for jacket top box
        Dim objInsTopBox As Object
        dTopBoxLength = 0.5 * (parFacetoFace - flangeThick - flangeThick2)
        dTopBoxWidth = 0.8 * dCylinderDiam
        dBoxBottomPos = 0.45 * dCylinderDiam
        dBoxTopPos = 0.55 * dCylinderDiam
        
        stPoint.Set -dTopBoxLength / 2, dBoxBottomPos, -dTopBoxWidth / 2 - parInsulationThickness
        enPoint.Set dTopBoxLength / 2, dBoxTopPos + parInsulationThickness, dTopBoxWidth / 2 + parInsulationThickness
        
        Set objInsTopBox = PlaceBox(m_OutputColl, stPoint, enPoint)
        
        ' Set the output
        m_OutputColl.AddOutput "InsTopBox", objInsTopBox
        Set objInsTopBox = Nothing
        
        ' Insert your code for Insulated Port1
        If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
            
            Dim ObjInsulatedPort1 As Object
            stPoint.Set -dFace1toCenter, 0, 0
            enPoint.Set -(dFace1toCenter - flangeThick - parInsulationThickness), 0, 0
        
            'Check to see that header insulation doesn't exceed face to face dimension
            If CmpDblGreaterthan(flangeThick + parInsulationThickness, parFacetoFace) Then
                enPoint.Set dFace2toCenter, 0, 0
            End If
            
            Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter1, True)
        
            ' Set the output
            m_OutputColl.AddOutput "InsulatedPort1", ObjInsulatedPort1
            Set ObjInsulatedPort1 = Nothing
        End If
        
        If CmpDblGreaterthan(flangeThick2, LINEAR_TOLERANCE) Then
            ' Insert your code for Insulated Port 2
            Dim ObjInsulatedPort2 As Object
            stPoint.Set dFace2toCenter, 0, 0
            enPoint.Set (dFace2toCenter - flangeThick2 - parInsulationThickness), 0, 0
            
            'Check to see that header insulation doesn't exceed face to face dimension
            If CmpDblGreaterthan(flangeThick2 + parInsulationThickness, parFacetoFace) Then
                enPoint.Set -dFace1toCenter, 0, 0
            End If
            
            Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter2, True)
        
            ' Set the output
            m_OutputColl.AddOutput "InsulatedPort2", ObjInsulatedPort2
            Set ObjInsulatedPort2 = Nothing
        End If
        
        If CmpDblGreaterthan(flangeThick3, LINEAR_TOLERANCE) Then
            ' Insert your code for Insulated Port 3
            
            Dim ObjInsulatedPort3 As Object
            stPoint.Set 0, -dFace3toCenter, 0
            enPoint.Set 0, -(dFace3toCenter - flangeThick3 - parInsulationThickness), 0
            
            'Check to see that branch insulation doesn't exceed face to center 3 dimension
            If CmpDblGreaterthan(flangeThick3 + parInsulationThickness, dFace3toCenter) Then
                enPoint.Set 0, 0, 0
            End If
            
            Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter3, True)
            
            ' Set the output
            m_OutputColl.AddOutput "InsulatedPort3", ObjInsulatedPort3
            Set ObjInsulatedPort3 = Nothing
        End If
        
    End Select
    
    Set stPoint = Nothing
    Set enPoint = Nothing

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
