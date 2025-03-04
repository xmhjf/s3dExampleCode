VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   COperation.cls
'   Author:         MA
'   Creation Date:  Monday, Dec 03 2007
'   Description:
'   This symbol is created for CR-129442 (Exhaust Air Nozzle)
'   Source: Toilet Exhaust Air Nozzle.pdf.
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Operation:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

   Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Integer
    
    Dim parWidth As Double
    Dim parBodyDiameter As Double
    Dim parBodyLength As Double
    Dim parRingDiameter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parBodyDiameter = arrayOfInputs(3)
    parBodyLength = arrayOfInputs(4)
    parRingDiameter = arrayOfInputs(5)
    
    'Insert your code for Output 5 (Operation Cylinder)
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set 0.6 * parBodyLength, 0, 0
    oEnPoint.Set parBodyLength, 0, 0
    
    Dim objCyl As Object
    Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth, True)
   
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
    Set objCyl = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub



