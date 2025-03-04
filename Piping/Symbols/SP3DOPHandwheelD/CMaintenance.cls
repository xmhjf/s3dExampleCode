VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   ProgID:         SP3DOPHandwheelD.HandwheelD
'   Author:         RUK
'   Creation Date:  Monday, 11 August 2008
'
'   Description:
'   This Symbol details were taken from Appendix E-138 of the Design Document
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   11.Aug.2008     RUK     CR-146837  Provide needle valve symbol
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:"    'Used for error messages

Private Sub Class_Initialize()
    '''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    
    Dim parOperatorHeight As Double
    Dim parOffsetFrmValCen As Double
    Dim parHandwheelDiameter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOffsetFrmValCen = arrayOfInputs(3)
    parHandwheelDiameter = arrayOfInputs(4)
    
    Dim iOutput As Long
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition

    Dim dStemDia As Double
    Dim dSpokeDia As Double
    Dim dHwTorDia As Double
    
    dStemDia = 0.05 * parHandwheelDiameter
    dSpokeDia = 0.05 * parHandwheelDiameter
   
    If dStemDia < 0.015 Then dStemDia = 0.015
    If dStemDia > 0.05 Then dStemDia = 0.05
    
    dSpokeDia = dStemDia
    dHwTorDia = dStemDia

    'Handwheel Stem
    Dim objHwStem As Object
    oStPoint.Set 0, parOffsetFrmValCen, 0
    oEnPoint.Set 0, parOperatorHeight - dHwTorDia / 2, 0
    
    Set objHwStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemDia, True)
    
    'Add to the Ouput
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHwStem
    Set objHwStem = Nothing
    
    'Handwheel
    Dim objHandwheel As Object
    oStPoint.Set 0, parOperatorHeight, 0
    oEnPoint.Set 0, parOperatorHeight - dHwTorDia, 0
    
    Set objHandwheel = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parHandwheelDiameter, True)
    
    'Add to the Ouput
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandwheel
    Set objHandwheel = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
               
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
