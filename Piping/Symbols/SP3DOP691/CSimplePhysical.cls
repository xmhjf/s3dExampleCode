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
'
'   CSimplePhysical.cls
'   Author:          MS
'   Creation Date:
'   Description:
'This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages

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
    
    Dim ObjFrontHeadBonnet As Object
    Dim ObjRearHeadMiterTop15degBody As Object
    
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    Dim ostpoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
        
'Insert your code for output 1(Body)
    ostpoint.Set 0, 0, 0
    oEndPoint.Set 0, 0.7 * parOperatorHeight, 0

' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), ostpoint, oEndPoint, 0.5 * parOperatorDiameter
        
'Insert your code for output 2(Body)
    ostpoint.Set 0, 0.7 * parOperatorHeight, 0
    oEndPoint.Set 0, parOperatorHeight, 0

' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), ostpoint, oEndPoint, parOperatorDiameter

    Set ostpoint = Nothing
    Set oEndPoint = Nothing
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

