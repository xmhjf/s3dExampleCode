VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          dkl
'   Creation Date:  August 14, 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect.
'   This symbol is based on information on Studded Cross in Design document and at www.woodcousa.com.
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------    -----        ------------------
'   14.Aug.2007     dkl     CR-125106 Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptoffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput As Integer
  
    Dim parLength As Double
    Dim parWidth As Double
    Dim parHeight As Double
    Dim parInsulationThickness As Double
   
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parLength = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parHeight = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0

 'Provide insulation to the top and bottom faces of the Studded cross.
    Dim oStPoint As AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
    oStPoint.Set -parLength / 2, -parHeight / 2 - parInsulationThickness, -parWidth / 2
    oEndPoint.Set parLength / 2, parHeight / 2 + parInsulationThickness, parWidth / 2
    
    Dim objInsCrossBody As Object
    Set objInsCrossBody = PlaceBox(m_OutputColl, oStPoint, oEndPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCrossBody
    Set objInsCrossBody = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
