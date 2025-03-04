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
'   Author:          NN
'   Creation Date:  Friday, Nov 24 2000
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----              ------------------
'   09.Jul.2003     SymbolTeam(India) Copyright Information, Header  is added.
'   22.Aug.2005     svsmylav          TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'  08.SEP.2006     KKC                DI-95670 Replace names with initials in all revision history sheets and symbols
'  07.Jan.2008     RRK                CR-131505 Enhanced the symbol to support part data basis values of 608, 609 and 610
'  07.JUL.2008     MA                 CR-145695 Implemented part data basis for the value 13
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI           As Double
Private Const MODULE = "Insulation:"    'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim cptoffset As Double
    Dim flangeDiam As Double
    Dim depth As Double

    Dim flangeThick2 As Double
    Dim flangeDiam2 As Double
    Dim cptoffset2 As Double
    Dim depth2 As Double

    Dim pipeDiam3 As Double
    Dim flangeThick3 As Double
    Dim flangeDiam3 As Double
    Dim cptoffset3 As Double
    Dim depth3 As Double

    Dim CenterPos As New AutoMath.DPosition
    CenterPos.Set 0, 0, 0

    Dim iOutput As Double

    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedBranch As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object
    Dim ObjInsulatedPort3 As Object

    Dim parFacetoFace As Double
    Dim parFace2toBranch As Double
    Dim parInsulationThickness As Double
    Dim parSeattoSeat As Double
    Dim parSeat1toCenter As Double
    Dim parSeat3toCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parAngle As Double

    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dFace3toCenter As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parFace2toBranch = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
'    parSeattoSeat = arrayOfInputs(5)
'    parSeat1toCenter = arrayOfInputs(6)
'    parSeat3toCenter = arrayOfInputs(7)
'    parFace1toCenter = arrayOfInputs(8)
'    parFace2toCenter = arrayOfInputs(9)
'    parFace3toCenter = arrayOfInputs(10)
'    parAngle = arrayOfInputs(11)

    iOutput = 0
    
    parAngle = PI / 4

    'Check to see that old instances of the symbol do not fail
    If UBound(arrayOfInputs) >= 11 Then parAngle = arrayOfInputs(11)

    iOutput = 0

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam2, cptoffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, cptoffset3, depth3

    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Select Case lPartDataBasis
    Case Is <= 1    'Default
        parFacetoFace = arrayOfInputs(2)
        parFace2toBranch = arrayOfInputs(3)

        dFace1toCenter = parFacetoFace - parFace2toBranch
        dFace2toCenter = parFace2toBranch
        dFace3toCenter = parFace2toBranch
        
    Case 13    'Asymmetrical Face-to-Center dimension basis
               'Face to center dimension for port 3 is same as that for port 2
        parFace1toCenter = arrayOfInputs(8)
        parFace2toCenter = arrayOfInputs(9)
        
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        dFace3toCenter = parFace2toCenter

    Case 608   'Y branches, single, specified by header face-to-face dimension, face-to-center 1, and face-to-center 3
        parFacetoFace = arrayOfInputs(2)
        parFace1toCenter = arrayOfInputs(8)
        parFace3toCenter = arrayOfInputs(10)

        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFacetoFace - dFace1toCenter
        dFace3toCenter = parFace3toCenter

    Case 609   'Y branches, single, specified by face-to-center 1, face-to-center 2, and face-to-center 3
        parFace1toCenter = arrayOfInputs(8)
        parFace2toCenter = arrayOfInputs(9)
        parFace3toCenter = arrayOfInputs(10)

        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        dFace3toCenter = parFace3toCenter

    Case 610  'Y branches, single, specified by header seat-to-seat dimension, seat-to-center 1, and seat-to-center 3
        parSeattoSeat = arrayOfInputs(5)
        parSeat1toCenter = arrayOfInputs(6)
        parSeat3toCenter = arrayOfInputs(7)

        dFace1toCenter = parSeat1toCenter + depth - cptoffset
        dFace2toCenter = parSeattoSeat - parSeat1toCenter + depth2 + cptoffset2
        dFace3toCenter = parSeat3toCenter + depth3 - cptoffset3

    Case Else
        GoTo ErrorLabel    'Invalid specification.
    End Select

    'Insert your code for output 1(Insulated Body)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    Dim parInsulationDiameter As Double

    stPoint.Set -(dFace1toCenter - flangeThick), 0, 0
    enPoint.Set dFace2toCenter - flangeThick2, 0, 0

    parInsulationDiameter = pipeDiam + parInsulationThickness * 2

    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing

    'Insert your code for output 2(Insulated Port1)
    Dim dFacetoFace As Double
    dFacetoFace = dFace1toCenter + dFace2toCenter
    
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    
    stPoint.Set -dFace1toCenter, 0, 0
    enPoint.Set -(dFace1toCenter - flangeThick - parInsulationThickness), 0, 0
    
    'Check to see that header insulation doesn't exceed face to face dimension
    If CmpDblGreaterthan(flangeThick + parInsulationThickness, dFacetoFace) Then
        enPoint.Set dFace2toCenter, 0, 0
    End If
    
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing

    'Insert your code for output 3(Insulated Port2)
    If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = 0.0001
    
    stPoint.Set dFace2toCenter, 0, 0
    enPoint.Set (dFace2toCenter - flangeThick2 - parInsulationThickness), 0, 0
    
    'Check to see that header insulation doesn't exceed face to face dimension
    If CmpDblGreaterthan(flangeThick2 + parInsulationThickness, dFacetoFace) Then
        enPoint.Set -dFace1toCenter, 0, 0
    End If
    
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    
    If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then parInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
    
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing

    'Insert your code for output 4(Insulated Branch)
    Dim BranchPoint As New AutoMath.DPosition
    
    If CmpDblEqual(flangeThick3, 0) Then flangeThick3 = 0.0001
    
    parInsulationDiameter = pipeDiam3 + parInsulationThickness * 2
    
    BranchPoint.Set (dFace3toCenter - flangeThick3) * Cos(parAngle), (dFace3toCenter - flangeThick3) * Cos(parAngle), CenterPos.z
    
    Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, CenterPos, BranchPoint, parInsulationDiameter, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
    Set ObjInsulatedBranch = Nothing

    'Insert your code for output 5 (Insulated Port3)
    stPoint.Set dFace3toCenter * Cos(parAngle), dFace3toCenter * Sin(parAngle), 0
    enPoint.Set (dFace3toCenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), _
                    (dFace3toCenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
    
    'Check to see that branch insulation doesn't exceed face to center 3 dimension
    If CmpDblGreaterthan(flangeThick3 + parInsulationThickness, dFace3toCenter) Then
        enPoint.Set 0, 0, 0
    End If
                    
    parInsulationDiameter = pipeDiam3 + parInsulationThickness * 2
    
    If CmpDblGreaterthan(flangeDiam3, pipeDiam3) Then parInsulationDiameter = flangeDiam3 + parInsulationThickness * 2

    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set BranchPoint = Nothing
    Set CenterPos = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
