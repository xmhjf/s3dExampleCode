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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         MA
'   Creation Date:  Friday, May 16 2008
'   Description:
'   Source: For PDB value 988: Forged Steel Valves Catalog, Bonney Forge, www.bonneyforge.com
'   For PDB value 13: Vogt Catalog
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   16.May.2008     MA      CR-141770 Created the symbol
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CInsulation:" 'Used for error messages

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
    Dim pipeDiam2       As Double
    Dim flangeThick2    As Double
    Dim sptOffset2      As Double
    Dim flangeDiam2     As Double
    Dim depth2          As Double
    
    Dim iOutput     As Double
    Dim objInsulationCylinder As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)
'    parFace2toCenter = arrayOfInputs(3)
'    parFacetoFace = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    
    iOutput = 0

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                        sptOffset2, depth2
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim dInsulationDiameter As Double
    Dim dInsulationDiameter1 As Double
    Dim dInsulationDiameter2 As Double
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        dInsulationDiameter1 = flangeDiam + parInsulationThickness * 2
    Else
        dInsulationDiameter1 = pipeDiam + parInsulationThickness * 2
    End If
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
        dInsulationDiameter2 = flangeDiam2 + parInsulationThickness * 2
    Else
        dInsulationDiameter2 = pipeDiam2 + parInsulationThickness * 2
    End If
    If CmpDblGreaterthan(dInsulationDiameter1, dInsulationDiameter2) Then
        dInsulationDiameter = dInsulationDiameter1
    Else
        dInsulationDiameter = dInsulationDiameter2
    End If
    
    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    ' This symbol is based on the following Part data basis values that govern its geometry
    ' (Part data Basis value -988): Extended body gate valve, specified by face-to-face, face-to-center 1
    ' (Part data Basis value -13): Asymmetrical Face-to-Center dimension basis
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 988
        parFace1toCenter = arrayOfInputs(2)
        parFacetoFace = arrayOfInputs(4)
        
        'Insert your code for output 4(InsulationCylinder)
        Dim dWallThick As Double
        RetrievePipeWallThick 1, oPartFclt, dWallThick
        Dim dInnerDia1 As Double
        dInnerDia1 = pipeDiam - (2 * dWallThick)
        
        oStPoint.Set 0, dInnerDia1 / 2 + parFacetoFace, 0
        oEnPoint.Set 0, dInnerDia1 / 2, 0
        Set objInsulationCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter2, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulationCylinder
        Set objInsulationCylinder = Nothing
    
    Case 13
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
    
        oStPoint.Set -parFace1toCenter, 0, 0
        oEnPoint.Set parFace2toCenter, 0, 0
        Set objInsulationCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulationCylinder
        Set objInsulationCylinder = Nothing
    
    Case Else
        GoTo ErrorLabel:
        
    End Select
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
