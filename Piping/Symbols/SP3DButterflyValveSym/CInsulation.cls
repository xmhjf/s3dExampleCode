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
'   Author:          BG
'   Creation Date:  Thursday, Jan 24 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004      svsmylav              Defined new input parameter HandwheelAngle.
'   02.Jan.2006       KKK                    CR-91106  Updated the symbol to be based on either face-to-face
'                                           or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)       Modified to accept default value of lPartDataBasis to be less than or equal to 1.
'   25.sep.06       KKC            TR-102869 Modified symbol code to use Face to Center or Face to Face to Dimension.
'                                            Modified symbol code to use Face to Face as a input parameter.
'   13.Jul.2007     MA             CR-123497 Implemented Part data basis to use Face 1 to Center and Face 2 to Center Dimensions.
'   26.Mar.2008     RUK            CR-139160 Butterfly valve, specified by face-to-face and face-to-center 1
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
    
    Dim iOutput     As Double
    Dim ObjInsulationCylinder As Object
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle   As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parHandwheelAngle = arrayOfInputs(4)
'    parFacetoCenter = arrayOfInputs(5)
'    parFace1toCenter = arrayOfInputs(6)
'    parFace2toCenter = arrayOfInputs(7)
    
    iOutput = 0

' Insert your code for output 4(Cylinder)
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim dInsulationDiameter As Double
    Dim dInsulationDia1 As Double
    Dim dInsulationDia2 As Double
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                     sptOffset, depth
                                                                            
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                sptOffset2, depth2
    
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
        parFacetoCenter = arrayOfInputs(5)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    
    Case 13   'Asymmetrical Face-to-Center dimension basis
        parFace1toCenter = arrayOfInputs(6)
        parFace2toCenter = arrayOfInputs(7)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        
    Case 894    'Butterfly valve, specified by face-to-face and face-to-center 1
        parFacetoFace = arrayOfInputs(2)
        parFace1toCenter = arrayOfInputs(6)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFacetoFace - dFace1toCenter
    
    Case Else
        GoTo ErrorLabel:
        
    End Select

    oStPoint.Set -dFace1toCenter, 0, 0
    oEnPoint.Set dFace2toCenter, 0, 0
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        dInsulationDia1 = flangeDiam + parInsulationThickness * 2
    Else
        dInsulationDia1 = pipeDiam + parInsulationThickness * 2
    End If
    
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
        dInsulationDia2 = flangeDiam2 + parInsulationThickness * 2
    Else
        dInsulationDia2 = pipeDiam2 + parInsulationThickness * 2
    End If
    
    If CmpDblGreaterthan(dInsulationDia1, dInsulationDia2) Then
        dInsulationDiameter = dInsulationDia1
    Else
        dInsulationDiameter = dInsulationDia2
    End If
    
    Set ObjInsulationCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationCylinder
    Set ObjInsulationCylinder = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oPipeComponent = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
