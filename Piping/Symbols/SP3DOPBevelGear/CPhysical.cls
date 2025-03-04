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
'   CPhysical.cls
'   Author:         KKC
'   Creation Date:  Tuesday, April 10 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol details were taken from Technegate knife gate valve bevel gear operator
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------  -----        ------------------
'   04.05.2007      KKC     Created: CR-117167  Create valve symbols for use in mining industry
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Const MIN_LENGTH = 0.0001

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    
    Dim iOutput As Double
    
    Dim ObjStem As Object
    Dim ObjOPCyl As Object
    Dim ObjArmWheel As Object
    Dim ObjArmShaft As Object
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
        
    Dim parArmLength As Double
    Dim dStemDia As Double
    Dim dExtnRodDia As Double
    Dim dHandleDia As Double
    Dim dHandWheelTk As Double
    
    Dim oStartPoint As AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    
    Set oStartPoint = New AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parArmLength = arrayOfInputs(4)
    
    iOutput = 0
    
    dStemDia = parArmLength / 15
    '   Check that the stem diameter is within 15mm and 50mm
    If CmpDblLessThan(dStemDia, 0.015) Then dStemDia = 0.015
    If CmpDblGreaterthan(dStemDia, 0.05) Then dStemDia = 0.05
    ' Assumed that both Extension Rod and Stem have the same Diameter
    dExtnRodDia = dStemDia
    dHandleDia = dStemDia
    dHandWheelTk = dStemDia
    
' Insert your code for output 1(Stem)
    oStartPoint.Set 0, 0, 0
    oEndPoint.Set 0, parOperatorHeight * 0.68, 0
    
    Set ObjStem = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, dStemDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 2(Bevel Gear Actuator Cylinder)
    oStartPoint.Set 0, parOperatorHeight * 0.625, 0
    oEndPoint.Set 0, parOperatorHeight * 0.775, 0
    
    Set ObjOPCyl = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parArmLength * 0.4, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOPCyl
    Set ObjOPCyl = Nothing
    
 ' Insert your code for output 3(Bevel Gear Shaft)
    oStartPoint.Set 0, parOperatorHeight * 0.7, 0
    oEndPoint.Set parArmLength - dHandWheelTk, oStartPoint.y, oStartPoint.z
                        
   Set ObjArmShaft = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, dHandleDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArmShaft
    Set ObjArmShaft = Nothing
    
 ' Insert your code for output 4(Arm Wheel)
    oStartPoint.Set oEndPoint.x, oEndPoint.y, oEndPoint.z
    oEndPoint.Set parArmLength, oStartPoint.y, oStartPoint.z
                        
    Set ObjArmWheel = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parOperatorDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArmWheel
    Set ObjArmWheel = Nothing
    
' Insert your code for output 5(Stem Housing)
    oStartPoint.Set 0, parOperatorHeight * 0.72, 0
    oEndPoint.Set 0, parOperatorHeight, 0
    Dim ObjStemHousing As Object
    Set ObjStemHousing = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parArmLength * 0.2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemHousing
    Set ObjStemHousing = Nothing
      
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
