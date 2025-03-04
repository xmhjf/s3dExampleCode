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
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No.: V42, Page No. D-26 of PDS Piping Component Data Reference Guide.
'    This Symbol is created with two insualation aspect outputs using 'PlaceCylinder' function.
'
'   Change History:
'   dd.mmm.yyyy     who                            change description
'   -----------     ---                                      ------------------
'   02.Jan.2006     KKK                           CR-91106  Updated the symbol to be based on either face-to-face
'                                                                  or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)       Modified to accept default value of lPartDataBasis to be less than or equal to 1.
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   20.Jul.2007     MA             CR-123497 Implemented Part data basis to use Face 1 to Center and Face 2 to Center Dimensions.
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

    Dim iOutput     As Double
    Dim ObjInsulCylinder As Object

    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double

    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    Dim EllipseMajorRadius      As Double
    Dim EllipseMinorRadius      As Double
    Dim InsulationDia   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'    parFace1toCenter = arrayOfInputs(4)
'    parFace2toCenter = arrayOfInputs(5)

    iOutput = 0

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 10    'Face-to-center dimension basis
        parFacetoCenter = arrayOfInputs(2)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    
    Case 13    'Asymmetrical Face-to-Center dimension basis
        parFace1toCenter = arrayOfInputs(4)
        parFace2toCenter = arrayOfInputs(5)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
 
    Case Else
        GoTo ErrorLabel:
        
    End Select
 
 ' Insert your code for output 3(Insulation Cylinder)
    InsulationDia = pipeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then InsulationDia = flangeDiam + 2 * parInsulationThickness
    
    stPoint.Set -dFace1toCenter, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
    Set ObjInsulCylinder = Nothing

' Insert your code for output 3(Insulation Cylinder)
    stPoint.Set 0, dFace2toCenter, 0
'   Extend the insulation along negative Y axis by ellipse minor radius plus parInsulationThickness
'   Consider Ellipse Majoj Axis is 0.6 times the larger face to center distance /2 and
'   Minor axis as 0.4 times the same center distance.
    EllipseMajorRadius = (0.6 * dFace1toCenter) / 2
    EllipseMinorRadius = (0.4 * dFace1toCenter) / 2
    enPoint.Set 0, -EllipseMinorRadius - parInsulationThickness, 0
    InsulationDia = pipeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then InsulationDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblLessThan(InsulationDia, (EllipseMajorRadius + parInsulationThickness) * 2) Then
        InsulationDia = (EllipseMajorRadius + parInsulationThickness) * 2
    End If
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

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
