VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CSimplePhysical.cls
'   Author:         MS
'   Creation Date:  Monday June 14, 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

    Private m_oSymGeomHelper As IJSymbolGeometryHelper
    Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
    Private Sub Class_Terminate()
        Set m_oSymGeomHelper = Nothing
    End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
   
    Dim iOutput     As Double
    Dim parFloorThickness As Double
    Dim parPostBaseToValveCenter As Double
    Dim parFloorLevelToValveCenter As Double
    Dim parTopOpeningDiameter As Double
    Dim parBottomOpeningDiameter As Double
    
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFloorThickness = arrayOfInputs(2)
    parFloorLevelToValveCenter = arrayOfInputs(3)
    parTopOpeningDiameter = arrayOfInputs(4)
    parBottomOpeningDiameter = arrayOfInputs(5)
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0
   
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

' Insert your code for output 1( Top Portion of the Floor Box)
    'Assuming the bottom portion height of the FloorBox is 0.2 times the Floor Thickness
    
    stPoint.Set 0, parFloorLevelToValveCenter, 0
    enPoint.Set 0, parFloorLevelToValveCenter - 0.8 * parFloorThickness, 0
    ' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parTopOpeningDiameter

' Insert your code for output 2(Bottom portion of the Floor Box)
    stPoint.Set 0, parFloorLevelToValveCenter - 0.8 * parFloorThickness, 0
    enPoint.Set 0, parFloorLevelToValveCenter - parFloorThickness, 0
    ' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parBottomOpeningDiameter
       
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

