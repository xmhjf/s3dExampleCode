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
'   Author:          svsmylav
'   Creation Date:  Wednesday, Jan 9 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav                Defined new input parameter HandwheelAngle.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   19.Aug.2005     svsmylav    TR-83739:   Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'   08.SEP.2006     KKC         DI-95670    Replace names with initials in all revision history sheets and symbols
'   21.FEB.2008     PK          CR-135702   Enhanced the existing 3 way control valve to support dimensional basis for
'                                               (i)Face-to-center dimension basis
'                                               (ii)Face-to-Face dimension and Face-to-Center dimension for port 3
'                                               (iii)Asymmetrical Face-to-Center dimension basis
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
    ''
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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput             As Double
    Dim ObjBodyCone1        As Object
    Dim ObjBodyCone2        As Object
    Dim ObjBodyCone3        As Object
    Dim ObjInsulatedCone1   As Object
    Dim ObjInsulatedPort1   As Object
    Dim ObjInsulatedCone2   As Object
    Dim ObjInsulatedPort2   As Object
    Dim ObjInsulatedCone3   As Object
    Dim ObjInsulatedPort3   As Object
    
    Dim parFace1toCenter        As Double
    Dim parFace2toCenter        As Double
    Dim parFace3toCenter        As Double
    Dim parAngle                As Double
    Dim parInsulationThickness  As Double
    
    Dim parFacetoFace       As Double
    Dim parFacetoCenter     As Double
    
    Dim dFace1toCenter      As Double
    Dim dFace2toCenter      As Double
    Dim dFace3toCenter      As Double
    
    'Variables defined by svsmylav
    Dim stPoint                 As New AutoMath.DPosition
    Dim enPoint                 As New AutoMath.DPosition
    Dim parInsDiameterPort      As Double
    Dim parInsulationDiameter   As Double

    Dim parHandwheelAngle   As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFace1toCenter = arrayOfInputs(2)
    'parFace2toCenter = arrayOfInputs(3)
    'parFace3toCenter = arrayOfInputs(4)
    parAngle = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)
    'parFacetoFace = arrayOfInputs(8)
    'parFacetoCenter = arrayOfInputs(9)
    
    iOutput = 0
    
    Dim oPipeComponent      As IJDPipeComponent
    Dim lPartdatabasis      As Long
    
    Set oPipeComponent = oPartFclt
    lPartdatabasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Select Case lPartdatabasis
    
        Case Is <= 1, 13                        'default case and Asymmetrical Face-to-Center dimension basis
            parFace1toCenter = arrayOfInputs(2)
            parFace2toCenter = arrayOfInputs(3)
            parFace3toCenter = arrayOfInputs(4)
            parAngle = arrayOfInputs(5)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dFace3toCenter = parFace3toCenter
            If CmpDblEqual(dFace1toCenter, 0) Then
                dFace1toCenter = dFace2toCenter
            ElseIf CmpDblEqual(dFace2toCenter, 0) Then
                dFace2toCenter = dFace1toCenter
            End If
                      
        Case 10                                 'Face-to-center dimension basis
            parFacetoCenter = arrayOfInputs(9)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dFace3toCenter = parFacetoCenter
                        
        Case 11                                 'Face-to-Face dimension and Face-to-Center dimension for port 3
            parFace3toCenter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(8)
            dFace1toCenter = parFacetoFace / 2
            dFace2toCenter = parFacetoFace / 2
            dFace3toCenter = parFace3toCenter
            
        Case Else
            GoTo ErrorLabel:
            
    End Select
    'Insert your code for output 4(Insulation Cone1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
    stPoint.Set CenterPos.x - (dFace1toCenter - flangeThick), CenterPos.y, CenterPos.z
    
    'PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
    'value removes the 'Invalid property value' failure.
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, parInsulationDiameter / 2, 0, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCone1
    Set ObjInsulatedCone1 = Nothing
    
    'Insert your code for output 5(Insulation Port1)
    stPoint.Set CenterPos.x - dFace1toCenter, CenterPos.y, CenterPos.z
    If CmpDblGreaterthan(flangeThick + parInsulationThickness, dFace1toCenter) Then
        enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    Else
        enPoint.Set CenterPos.x - (dFace1toCenter - flangeThick - parInsulationThickness), _
                CenterPos.y, _
                CenterPos.z
    End If
    parInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsDiameterPort = flangeDiam + parInsulationThickness * 2
                
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameterPort, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing
    
    'Insert your code for output 6(Insulation Cone2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
    stPoint.Set CenterPos.x + dFace2toCenter - flangeThick, CenterPos.y, CenterPos.z
    
    'PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
    'value removes the 'Invalid property value' failure.
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, parInsulationDiameter / 2, 0, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCone2
    Set ObjInsulatedCone2 = Nothing
    
    'Insert your code for output 7(Insulation Port2)
    stPoint.Set CenterPos.x + dFace2toCenter, CenterPos.y, CenterPos.z
    If CmpDblGreaterthan(flangeThick + parInsulationThickness, dFace2toCenter) Then
        enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    Else
        enPoint.Set CenterPos.x + dFace2toCenter - flangeThick - parInsulationThickness, _
                CenterPos.y, _
                CenterPos.z
    End If
    
    parInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsDiameterPort = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameterPort, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    
    'Insert your code for output 8(Insulation Cone3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
    stPoint.Set CenterPos.x + (dFace3toCenter - flangeThick) * Cos(parAngle), _
                CenterPos.y + (dFace3toCenter - flangeThick) * Sin(parAngle), _
                CenterPos.z
    
    'PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
    'value removes the 'Invalid property value' failure.
    enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    
    parInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsDiameterPort = flangeDiam + parInsulationThickness * 2
    
    'If Insulation Diameter for port is greater than sum of Face 1 to Center and Face 2 to Center
    'then setting insulation diameter equal to sum of Face 1 to Center and Face 2 to Center
    If CmpDblGreaterthan(parInsDiameterPort, dFace1toCenter + dFace2toCenter) Then
        parInsDiameterPort = dFace1toCenter + dFace2toCenter
    End If
    
    'This check is to ensure that the base diameter for insulation cone of port 3 does not
    'exceed the insualtion diameter for port 3
    If CmpDblGreaterthan(parInsulationDiameter, parInsDiameterPort) Then
        Set ObjInsulatedCone3 = PlaceCone(m_OutputColl, stPoint, enPoint, parInsDiameterPort / 2, 0, True)
    Else
        Set ObjInsulatedCone3 = PlaceCone(m_OutputColl, stPoint, enPoint, parInsulationDiameter / 2, 0, True)
    End If
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCone3
    Set ObjInsulatedCone3 = Nothing
    
    'Insert your code for output 9(Insulation Port3)
    stPoint.Set CenterPos.x + dFace3toCenter * Cos(parAngle), _
                CenterPos.y + dFace3toCenter * Sin(parAngle), CenterPos.z
    If CmpDblGreaterthan(flangeThick + parInsulationThickness, dFace3toCenter) Then
        enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    Else
        enPoint.Set CenterPos.x + (dFace3toCenter - flangeThick - parInsulationThickness) * Cos(parAngle), _
                CenterPos.y + (dFace3toCenter - flangeThick - parInsulationThickness) * Sin(parAngle), _
                CenterPos.z
    End If
    
    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameterPort, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
End Sub
