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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday 30, Jul 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Symbol Model No.: V41, Page No. D-25 of PDS Piping Component Data Reference Guide.
'   This Symbol is created with one insualation aspect outputs using 'PlaceCylinder' function.
'   Change History:

'   dd.mmm.yyyy         who                 change description
'   -----------         -----               ------------------
'   02.Jan.2006     KKK       CR-91106  Updated the symbol to be based on either face-to-face
'                                           or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)       Modified to accept default value of lPartDataBasis to be less than or equal to 1.
'   25.sep.06       KKC            TR-102869 Modified symbol code to use Face to Center or Face to Face to Dimension.
'                                            Modified symbol code to use Face to Face as a input parameter.
'   18.Jul.2007     MA             CR-123497 Implemented Part data basis to use Face 1 to Center and Face 2 to Center Dimensions.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
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

    Dim pipeDiam2        As Double
    Dim flangeDiam2      As Double

    Dim iOutput     As Double
    Dim ObjInsulCylinder As Object

    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    Dim EllipseMinorRadius As Double
    Dim dInsulationDia As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'    parFacetoCenter = arrayOfInputs(4)
'    parFace1toCenter = arrayOfInputs(5)
'    parFace2toCenter = arrayOfInputs(6)

    iOutput = 0

'   Compute insulation diameter
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 5  'Face-to-face dimension basis
        parFacetoFace = arrayOfInputs(2)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
    
    Case 10    'Face-to-center dimension basis
        parFacetoCenter = arrayOfInputs(4)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    
    Case 13   'Asymmetrical Face-to-Center dimension basis
        parFace1toCenter = arrayOfInputs(5)
        parFace2toCenter = arrayOfInputs(6)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
    
    Case Else
        GoTo ErrorLabel:
        
    End Select
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth

' Insert your code for output 3(Insulation Cylinder)
    stPoint.Set -dFace1toCenter, 0, 0
    enPoint.Set dFace2toCenter, 0, 0

    dInsulationDia = ReturnMax4(pipeDiam, flangeDiam, pipeDiam2, flangeDiam2) + 2 * parInsulationThickness

'   Extend the insulation along negative Y axis by ellipse minor radius plus parInsulationThickness
'   Consider Ellipse Major Axis is 0.6 times the larger face to center distance /2 and
'   Minor axis as 0.4 times the same center distance.
    EllipseMinorRadius = (0.4 * (dFace1toCenter + dFace2toCenter) / 2) / 2
    If CmpDblLessThan(dInsulationDia, (EllipseMinorRadius + 2 * parInsulationThickness)) Then
        dInsulationDia = EllipseMinorRadius + 2 * parInsulationThickness
    End If
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
    Set ObjInsulCylinder = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPipeComponent = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
