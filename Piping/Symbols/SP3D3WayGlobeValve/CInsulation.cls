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
'   Creation Date:  Thursday, Jun 13 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model: V13, Page No. D-15 of PDS Piping Component Data Reference Guide.
' This Symbol is created with two insualation aspect outputs using 'PlaceCylinder' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav                Defined new input parameter HandwheelAngle.
'   08.SEP.2006     KKC      DI-95670       Replace names with initials in all revision history sheets and symbols
'   21.FEB.2008     PK       CR-135701      Enhanced the existing 3 way globe valve to support dimensional basis for
'                                               (i)Face-to-center dimension basis
'                                               (ii)Face-to-Face dimension and Face-to-Center dimension for port 3
'                                               (iii)Asymmetrical Face-to-Center dimension basis
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub
Private Function ReturnMax3(A#, B#, C#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    ReturnMax3 = MaxValue
End Function
Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    If CmpDblLessThan(MaxValue, D) Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim pipeDiam2       As Double
    Dim flangeDiam2     As Double
    
    Dim iOutput         As Double
    Dim ObjInsHorBody   As Object
    Dim ObjInsVertBody  As Object
    
    Dim parFace1toCenter        As Double
    Dim parFace2toCenter        As Double
    Dim parFace3toCenter        As Double
    Dim parInsulationThickness  As Double

    Dim stPoint                 As New AutoMath.DPosition
    Dim enPoint                 As New AutoMath.DPosition
    Dim ellipseMajorRadius      As Double
    Dim EllipseMinorRadius      As Double
    Dim parInsulationDiameter   As Double

    Dim parHandwheelAngle   As Double
    
    Dim parFacetoFace       As Double
    Dim parFacetoCenter     As Double
    
    Dim dFace1toCenter      As Double
    Dim dFace2toCenter      As Double
    Dim dFace3toCenter      As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFace1toCenter = arrayOfInputs(2)
    'parFace2toCenter = arrayOfInputs(3)
    'parFace3toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    'parFacetoFace = arrayOfInputs(7)
    'parFacetoCenter = arrayOfInputs(8)
    
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
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dFace3toCenter = parFace3toCenter
            If CmpDblEqual(dFace1toCenter, 0) Then
                dFace1toCenter = dFace2toCenter
            ElseIf CmpDblEqual(dFace2toCenter, 0) Then
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
        
        Case Else
            GoTo ErrorLabel:
            
    End Select

    'Insert your code for output 1(Insulation for Valve Body - Horizontal)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth

    'Maximum of pipeDiam1, flangeDiam1, pipeDiam2 and flangeDiam2 is used to compute
    'insulation diameter
    parInsulationDiameter = ReturnMax4(pipeDiam, flangeDiam, pipeDiam2, flangeDiam2) + _
                            2 * parInsulationThickness
    stPoint.Set -dFace1toCenter, 0, 0
    enPoint.Set dFace2toCenter, 0, 0
    Set ObjInsHorBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsHorBody
    Set ObjInsHorBody = Nothing
    
    'Insert your code for output 2(Insulation for Valve Body - Vertical)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    stPoint.Set 0, dFace3toCenter, 0
    'Extend the insulation along negative Y axis by ellipse minor radius plus parInsulationThickness
    'Consider Ellipse Major Axis is 0.6 times the larger face to center distance /2 and
    'Minor axis as 0.4 times the same center distance.
    
    ellipseMajorRadius = (0.6 * ReturnMax3(dFace1toCenter, dFace2toCenter, dFace3toCenter)) / 2
    EllipseMinorRadius = 0.4 * ellipseMajorRadius
    enPoint.Set 0, -EllipseMinorRadius - parInsulationThickness, 0
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    If CmpDblLessThan(parInsulationDiameter, (ellipseMajorRadius + parInsulationThickness) * 2) Then
        parInsulationDiameter = (ellipseMajorRadius + parInsulationThickness) * 2
    End If
    
    'Insulation diameter for port 3 is equated to twice the minimum of Face 1 to Center
    'and Face 2 to Center
    Dim dMinValue As Double
    If CmpDblLessThan(dFace1toCenter, dFace2toCenter) Then
        dMinValue = dFace1toCenter
    Else
        dMinValue = dFace2toCenter
    End If

    If CmpDblGreaterthan(parInsulationDiameter / 2, dMinValue) Then
        parInsulationDiameter = 2 * dMinValue
    End If
    
    Set ObjInsVertBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsVertBody
    Set ObjInsVertBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
