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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         RRK
'   Creation Date:  Thursday, Jul 24 2006
'   Description:
'    Physical Aspect of Electric Actuator (Operator data is at www.valtorc.com/auto_butter.htm)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private PI As Double
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
  
    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
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
    iOutput = 0
    
    Dim parActuatorHeight As Double
    Dim parActuatorLength As Double
    Dim parActuatorWidth As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parActuatorHeight = arrayOfInputs(2)
    parActuatorLength = arrayOfInputs(3)
    parActuatorWidth = arrayOfInputs(4)
    
    Dim stPoint  As Object
    Dim enPoint  As Object
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
     
     m_oGeomHelper.OutputCollection = m_OutputColl

' Place Output 1(Stem)
'------------------------------------------------------------------------------------------------------------------------------------------
    Dim dStemMinimumDiameter As Double
    
'Assumption:
    dStemMinimumDiameter = parActuatorWidth * 0.1
    
'   Check that the stem diameter is within 15mm and 50mm
   If dStemMinimumDiameter < 0.015 Then dStemMinimumDiameter = 0.015
   If dStemMinimumDiameter > 0.05 Then dStemMinimumDiameter = 0.05
   
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parActuatorHeight, 0
    
    ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dStemMinimumDiameter
    
    
' Place Output 2(Handwheel)
'------------------------------------------------------------------------------------------------------------------------------------------
    Dim dHandWheelDia As Double
    Dim dHandWheelThickness As Double
    
    dHandWheelDia = 0.6 * parActuatorLength
    dHandWheelThickness = 0.1 * parActuatorWidth
    
    stPoint.Set 0, parActuatorHeight, 0
    enPoint.Set 0, parActuatorHeight - dHandWheelThickness, 0
    
    ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dHandWheelDia
    

    
' Insert your code for output 3(Actuator)
'------------------------------------------------------------------------------------------------------------------------------------------
    Dim dBoxPosY As Double
    Dim dBoxHt As Double
    
'Assumption:
    dBoxPosY = 0.5 * parActuatorHeight
    dBoxHt = parActuatorHeight - dBoxPosY - 2 * dHandWheelThickness

    stPoint.Set parActuatorLength / 3, dBoxPosY, parActuatorWidth / 2
    enPoint.Set -2 * parActuatorLength / 3, dBoxPosY + dBoxHt, -parActuatorWidth / 2
    
    Dim ObjBox As Object
    Set ObjBox = New IngrGeom3D.Projection3d
    Set ObjBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox
    Set ObjBox = Nothing
    
   
    
' Insert your code for output 4(Cylinder)
'------------------------------------------------------------------------------------------------------------------------------------------
    Dim dCylHt As Double
    Dim dCylDia As Double
    
    dCylHt = 0.1 * parActuatorHeight
    dCylDia = 0.5 * parActuatorWidth
    
    stPoint.Set 0, dBoxPosY, 0
    enPoint.Set 0, dBoxPosY - dCylHt, 0
    
    ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylDia

    Set stPoint = Nothing
    Set enPoint = Nothing
    
   
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
