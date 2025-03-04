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
'
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CMaintainence.cls
'   Author:         KKC
'   Creation Date:  Tuesday, April 10, 2007
'   Description: This symbol is graphical implementation of the Technegate hand wheel operator
'   This symbol implements following partdatabasis
'   Default
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.05.2007      KKC     Created: CR-117167  Create valve symbols for use in mining industry
'   22.08.2008      RUK     CR-148066  Provide more realistic handwheel operator symbol
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
    Dim iOutput As Double

    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parOffsetFrmValCen As Double
    Dim parHandwheelDiameter As Double
    Dim parHandwheelToStemTop As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parOperatorHeight = arrayOfInputs(2)
    'parOperatorDiameter = arrayOfInputs(3)
    'parOffsetFrmValCen = arrayOfInputs(4)
    'parHandwheelDiameter = arrayOfInputs(5)
    'parHandwheelToStemTop = arrayOfInputs(6)

    'Initialize
    iOutput = 0

    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition

    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oCenter = New DPosition

    Dim dStemDia As Double
    Dim dSpokeDia As Double
    Dim dHwTorDia As Double
    
    'Checking for the PartDataBasis
    Dim oOperator As IJValveOperator
    Dim lPartDataBasis As Long
    Set oOperator = oPartFclt
    lPartDataBasis = oOperator.DimensionalBasis
    Set oOperator = Nothing

    Select Case lPartDataBasis
    Case Is <= 1
        parOperatorHeight = arrayOfInputs(2)
        parOperatorDiameter = arrayOfInputs(3)
        'Insert your code for output 1(Cylinder)
        Dim ObjMaintCyl As Object
        oStPoint.Set 0, parOperatorHeight * 0.2, 0
        oEnPoint.Set 0, parOperatorHeight, 0
    
        Set ObjMaintCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parOperatorDiameter, True)
    
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl
        Set ObjMaintCyl = Nothing
    
    Case 50
        parOperatorHeight = arrayOfInputs(2)
        parOffsetFrmValCen = arrayOfInputs(4)
        parHandwheelDiameter = arrayOfInputs(5)
        parHandwheelToStemTop = arrayOfInputs(6)
        
        dStemDia = 0.05 * parHandwheelDiameter
        dSpokeDia = 0.05 * parHandwheelDiameter
       
        If CmpDblLessThan(dStemDia, 0.015) Then dStemDia = 0.015
        If CmpDblGreaterthan(dStemDia, 0.05) Then dStemDia = 0.05
        
        dSpokeDia = dStemDia
        dHwTorDia = dStemDia
    
        'Handwheel Stem
        Dim objHwStem As Object
        oStPoint.Set 0, parOffsetFrmValCen, 0
        oEnPoint.Set 0, parOperatorHeight, 0
        
        Set objHwStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemDia, True)
        
        'Add to the Ouput
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHwStem
        Set objHwStem = Nothing
        
        'Handwheel
        Dim objHandwheel As Object
        oStPoint.Set 0, parOperatorHeight - parHandwheelToStemTop + dHwTorDia / 2, 0
        oEnPoint.Set 0, parOperatorHeight - parHandwheelToStemTop - dHwTorDia / 2, 0
        
        Set objHandwheel = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parHandwheelDiameter, True)
        
        'Add to the Ouput
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "Handwheel", objHandwheel
        Set objHandwheel = Nothing
    
    Case Else
        GoTo ErrorLabel
    End Select
        
    'Remove the References
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
