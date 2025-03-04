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
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Thursday, Jan 17 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle.
'   26.Jul.2006     dkl                     CR-28369 Added implementation of PartDataBasis 5
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  04.05.2007      KKC         CR-117167  Implemented Part data basis for 290, 295 and 300
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Dim PI           As Double

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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjValveBodyBox As Object
    Dim ObjInsulatedCyl As Object
    Dim ObjInsulatedBox As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parValveHeight As Double
    Dim parInsulationThickness As Double
    Dim parFacetoFace As Double
    Dim parFrameHeight As Double
    Dim parFrameWidth As Double
    Dim parFrameDepth As Double
    Dim parValveWidth As Double
    Dim dFrameDepth1 As Double
    Dim dFrameDepth2 As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsulationDiameter   As Double
    Dim pipeDiam1               As Double
    Dim flangeDiam1             As Double
    Dim pipeDiam2               As Double
    Dim flangeDiam2             As Double
   
    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)    'Set only if PartDataBasis=1 or 10
'    parFace2toCenter = arrayOfInputs(3)    'Set only if PartDataBasis=1 or 10
    parValveHeight = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
'    parFacetoFace = arrayOfInputs(7)       'Set only if PartDataBasis=5
'    parFrameHeight = arrayOfInputs(8)
'    parFrameWidth = arrayOfInputs(9)
'    parFrameDepth = arrayOfInputs(10)
'    parValveWidth = arrayOfInputs(11)
    iOutput = 0
    
 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Select Case lPartDataBasis
    Case 5, 295  'Applicable for 5 "Face-to-face dimension basis"
        parFacetoFace = arrayOfInputs(7)
        parFace1toCenter = parFacetoFace / 2  'parFace1toCenter and parFace2toCenter: These dimensions are calculated and used in computation.
        parFace2toCenter = parFacetoFace / 2
        
        If UBound(arrayOfInputs) >= 10 Then
            parFrameDepth = arrayOfInputs(10)
        End If
        
        If CmpDblEqual(parFrameDepth, 0) Then
           dFrameDepth1 = parFacetoFace / 2
           dFrameDepth2 = parFacetoFace / 2
        Else
           dFrameDepth1 = parFrameDepth / 2
           dFrameDepth2 = parFrameDepth / 2
        End If
           
    Case Is <= 1, 10  'Applicable for  1 "Default" & 10 -"Face-to-center dimension basis"
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
        
        If UBound(arrayOfInputs) >= 10 Then
            parFrameDepth = arrayOfInputs(10)
        End If
        
        If CmpDblEqual(parFrameDepth, 0) Then
           dFrameDepth1 = parFace1toCenter
           dFrameDepth2 = parFace2toCenter
        Else
           dFrameDepth1 = parFrameDepth / 2
           dFrameDepth2 = parFrameDepth / 2
        End If
    Case 290
        parFacetoFace = arrayOfInputs(7)
        parFrameHeight = arrayOfInputs(8)
        parFrameWidth = arrayOfInputs(9)
        parFrameDepth = arrayOfInputs(10)
        parValveWidth = arrayOfInputs(11)
        If CmpDblEqual(parFrameHeight, 0) Then parFrameHeight = parValveHeight * 0.6
        If CmpDblEqual(parFrameWidth, 0) Then parFrameWidth = parValveWidth / 2
        If CmpDblEqual(parFrameDepth, 0) Then parFrameDepth = parFacetoFace / 2
    Case 300
        parFacetoFace = arrayOfInputs(7)
        parFrameWidth = arrayOfInputs(9)
        parFrameDepth = arrayOfInputs(10)
        If CmpDblEqual(parFrameWidth, 0) Then parFrameWidth = parValveWidth * 0.6
        If CmpDblEqual(parFrameDepth, 0) Then parFrameDepth = parFacetoFace / 2
    Case Else
        GoTo ErrorLabel
    End Select
 ' Insert your code for output 2(Insulation for Cylinder)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth

'   Maximum of pipeDiam1, flangeDiam1, pipeDiam2 and flangeDiam2 is used to compute
'   insulation diameter
    parInsulationDiameter = ReturnMax4(pipeDiam1, flangeDiam1, pipeDiam2, flangeDiam2) + _
                            2 * parInsulationThickness
                            
    If lPartDataBasis = 5 Or lPartDataBasis = 290 Or _
                    lPartDataBasis = 295 Or lPartDataBasis = 300 Then
        parFace1toCenter = parFacetoFace / 2
        parFace2toCenter = parFacetoFace / 2
    End If
    
    stPoint.Set CenterPos.x - parFace1toCenter, _
                CenterPos.y, _
                CenterPos.z
    enPoint.Set CenterPos.x + parFace2toCenter, _
                CenterPos.y, _
                CenterPos.z
    Set ObjInsulatedCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCyl
    Set ObjInsulatedCyl = Nothing
    
    If (lPartDataBasis <= 1) Or (lPartDataBasis = 5) Or (lPartDataBasis = 295) Or (lPartDataBasis = 10) Then
     ' Insert your code for output 3(Insulation for Box)
        stPoint.Set CenterPos.x - parFace1toCenter - parInsulationThickness, _
                    CenterPos.y - parInsulationThickness, _
                    CenterPos.z - pipeDiam1 / 2 - parInsulationThickness
        enPoint.Set CenterPos.x + parFace2toCenter + parInsulationThickness, _
                    CenterPos.y + parValveHeight + parInsulationThickness, _
                    CenterPos.z + pipeDiam1 / 2 + parInsulationThickness
        Set ObjInsulatedBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBox
        Set ObjInsulatedBox = Nothing
    ElseIf (lPartDataBasis = 290) Then
       ' Insert your code for output 3(Insulation for Box)
        stPoint.Set CenterPos.x - parFacetoFace / 2 + flangeThick - parInsulationThickness, _
                    CenterPos.y - parValveHeight / 2 - parInsulationThickness, _
                    CenterPos.z - parValveWidth / 2 - parInsulationThickness
        enPoint.Set CenterPos.x + parFacetoFace / 2 - flangeThick + parInsulationThickness, _
                    CenterPos.y + parValveHeight / 2 + parInsulationThickness, _
                    CenterPos.z + parValveWidth / 2 + parInsulationThickness
        Set ObjInsulatedBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    
    ' Set the outputs
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBox
        Set ObjInsulatedBox = Nothing
    
    ' Insert your code for output 5( Insulation for Frame for Operator support)
        Dim objFrame As Object
        Dim oBaseCenter As AutoMath.DPosition
        Set oBaseCenter = New AutoMath.DPosition
        oBaseCenter.Set 0, parValveHeight / 2, 0
        
        Set objFrame = PlaceTrapezoid(m_OutputColl, oBaseCenter, parValveWidth + 2 * parInsulationThickness, _
                              parFrameDepth + 2 * parInsulationThickness, parFrameWidth + 2 * parInsulationThickness, parFrameDepth + 2 * parInsulationThickness, _
                                        parFrameHeight + parInsulationThickness, True, -PI / 2, 0, PI / 2)
                                        
        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFrame
        Set objFrame = Nothing
        Set stPoint = Nothing
        Set enPoint = Nothing
        Set oBaseCenter = Nothing
    ElseIf (lPartDataBasis = 300) Then
        ' Insert your code for output 3(Insulation for Box)
        stPoint.Set CenterPos.x - parFrameDepth / 2 - parInsulationThickness, _
                    CenterPos.y - parInsulationThickness, _
                    CenterPos.z - parFrameWidth / 2 - parInsulationThickness
        enPoint.Set CenterPos.x + parFrameDepth / 2 + parInsulationThickness, _
                    CenterPos.y + parValveHeight + parInsulationThickness, _
                    CenterPos.z + parFrameWidth / 2 + parInsulationThickness
        Set ObjInsulatedBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBox
        Set ObjInsulatedBox = Nothing
    End If
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    If CmpDblLessThan(MaxValue, D) Then MaxValue = D
    ReturnMax4 = MaxValue
End Function
