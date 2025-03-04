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
'   Creation Date:  Thursday, September 08 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   08.09.2004      MS  Added Maintenace Envelope to the operator
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
    Dim parStemAngle As Double
    Dim parOperatorDiameter As Double
    Dim parStemExtension As Double
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parStemAngle = arrayOfInputs(3)
    parOperatorDiameter = arrayOfInputs(4)

'   Initialize
    iOutput = 0
    parStemExtension = parOperatorHeight * 0.1

    
' Insert your code for output 1(Cone)
    Dim ObjMaintCone As Object
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    stPoint.Set -(parOperatorHeight - parStemExtension) * Cos(parStemAngle), _
                    (parOperatorHeight - parStemExtension) * Sin(parStemAngle), 0
    enPoint.Set 0, 0, 0
    Set ObjMaintCone = PlaceCone(m_OutputColl, stPoint, enPoint, parOperatorDiameter / 2, 0, True)


' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCone
    Set ObjMaintCone = Nothing

 ' Insert your code for output 2(Cylinder)
    Dim ObjMaintCyl As Object
    stPoint.Set -(parOperatorHeight - parStemExtension) * Cos(parStemAngle), (parOperatorHeight - parStemExtension) * Sin(parStemAngle), 0
    enPoint.Set -(parOperatorHeight) * Cos(parStemAngle), (parOperatorHeight) * Sin(parStemAngle), 0
    Set ObjMaintCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl
    Set ObjMaintCyl = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
