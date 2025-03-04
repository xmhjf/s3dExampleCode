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
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:          MS
'   Creation Date:  Wednesday, September 22 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   08.22.2004      MS  Added Maintenace Envelope to the operator
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   22.Feb.2007         KKC           TR-112500 Maintenance envelopes for some valve operators do not use capped ends
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double

    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parStemAngle As Double

'   Inputs
    Set oPartFclt = arrayOfInputs(1)
    parStemAngle = arrayOfInputs(2)
    parOperatorHeight = arrayOfInputs(3)
    parOperatorDiameter = arrayOfInputs(4)
    
'   Variables defined by svsmylav
    Dim CenterPos           As New AutoMath.DPosition
    Dim ConeCenterBasePt    As New AutoMath.DPosition
    Dim ConeCenterTopPt     As New AutoMath.DPosition
    
'   Initialize
    iOutput = 0

'   Prepare data points
    CenterPos.Set 0, 0, 0
    
    ConeCenterBasePt.Set CenterPos.x - parOperatorHeight * Cos(parStemAngle), _
                         CenterPos.y + parOperatorHeight * Sin(parStemAngle), _
                         CenterPos.z
    ConeCenterTopPt.Set CenterPos.x, CenterPos.y, CenterPos.z
    
' Insert your code for output 1(Cone)
'   Place Cone
    Dim objMaintenanceCone     As IngrGeom3D.Cone3d
    Set objMaintenanceCone = PlaceCone(m_OutputColl, ConeCenterBasePt, ConeCenterTopPt, parOperatorDiameter, 0, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMaintenanceCone
    Set objMaintenanceCone = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub



