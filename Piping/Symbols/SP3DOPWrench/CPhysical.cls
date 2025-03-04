VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Friday, Sept 10 2007
'   Description     :  Wrench type operator to be used with 3 way diverter combination valve
'                      of Tyco Flow Control
'   Reference       :  http://www.tycoflowcontrol-pc.com/ld/F605_4_07.pdf
'
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   24.Aug.2007     PK          CR-126718:Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim parOperatorHeight   As Double
    Dim parLeverLength      As Double
    
    Dim iOutput              As Integer
    Dim dFlangeCenterHeight  As Double
    Dim dFlangeWidth         As Double
    Dim dFlangeThickness     As Double
    Dim dStemVisibleHeight   As Double
    Dim dCylinderHeight      As Double
    Dim dWrenchThickness     As Double
    
    Dim oCentPos   As AutoMath.DPosition
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    Set oCentPos = New AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
       
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parLeverLength = arrayOfInputs(3)
    
    iOutput = 0
    oCentPos.Set 0, 0, 0
        
    'Insert your code for Flange(Output 1)
    Dim oFlange As Object
    
    'Assumptions:1)Connection point of two Flanges is assumed to be at 25% of Operator Height
    '            2)Width of the flange is assumed to be 40% of Lever Length
    '            3)Flange Thickness is assumed to be 1/32 times of Operator Height
    dFlangeCenterHeight = 0.25 * parOperatorHeight
    dFlangeWidth = 0.4 * parLeverLength
    dFlangeThickness = parOperatorHeight / 32
    
    oStPoint.Set oCentPos.x - dFlangeWidth / 2, oCentPos.y + dFlangeCenterHeight - dFlangeThickness, _
                                                oCentPos.z - dFlangeWidth / 2
    oEnPoint.Set oCentPos.x + dFlangeWidth / 2, oCentPos.y + dFlangeCenterHeight + dFlangeThickness, _
                                                oCentPos.z + dFlangeWidth / 2
    Set oFlange = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set Output 1
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oFlange
    Set oFlange = Nothing
        
    'Insert your code for Body(Output 2)
    Dim oBody As Object
    
    'Assumption:1)Cylinder Height is assumed to be 75% of Operator Height
    dCylinderHeight = 3 * parOperatorHeight / 4
    
    oStPoint.Set 0, 0, 0
    oEnPoint.Set oCentPos.x, oCentPos.y + dCylinderHeight, oCentPos.z
    Set oBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.3 * parLeverLength, True)
        
    'Set Output 2
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
    Set oBody = Nothing
    
    'Insert your code for Yoke Brush Wrench(Output 3)
    Dim oYokeBrushWrench As Object
    
    'Assumptions:1)Wrench Thickness is assumed to be 1/16 times of Operator Height
    dWrenchThickness = parOperatorHeight / 16
    
    oStPoint.Set oCentPos.x - parLeverLength / 2, oCentPos.y + dCylinderHeight, oCentPos.z + dWrenchThickness / 2
    oEnPoint.Set oCentPos.x + parLeverLength / 2, oCentPos.y + dCylinderHeight + dWrenchThickness, oCentPos.z - dWrenchThickness / 2
    Set oYokeBrushWrench = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
     'Set Output 3
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oYokeBrushWrench
    Set oYokeBrushWrench = Nothing

    'Insert your code for Visible Stem(Output 4)
    Dim oVisibleStem As Object
    
    'Assumption:Stem Visible Height is assumed to be 12.5% of Operator Height
    dStemVisibleHeight = parOperatorHeight / 8
    
    oStPoint.Set oCentPos.x, oCentPos.y + dCylinderHeight + dWrenchThickness, oCentPos.z
    oEnPoint.Set oCentPos.x, oCentPos.y + dCylinderHeight + dWrenchThickness + dStemVisibleHeight, oCentPos.z
    Set oVisibleStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.3 * dFlangeWidth, True)
    
     'Set Output 4
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oVisibleStem
    Set oVisibleStem = Nothing
    
    'Insert your code for StemWrench(Output5)
    Dim oStemWrench As Object
    oStPoint.Set oCentPos.x - parLeverLength / 2, oCentPos.y + dCylinderHeight + dWrenchThickness + dStemVisibleHeight, oCentPos.z + dWrenchThickness / 2
    oEnPoint.Set oCentPos.x + parLeverLength / 2, oCentPos.y + parOperatorHeight, oCentPos.z - dWrenchThickness / 2
    Set oStemWrench = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
     'Set Output 5
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oStemWrench
    Set oStemWrench = Nothing
    
    Set oCentPos = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
      
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub


