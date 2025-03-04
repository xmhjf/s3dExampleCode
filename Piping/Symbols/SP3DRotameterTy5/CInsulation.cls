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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Tuesday, Sep 17 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR5,
' PDMS Category Name=IFNK). Physical aspect consist of one vertical cylinder, One Horizontal Box
' and two nozzles, one facing -ive and the other facing +ive direction of X-axis.
' Insulation aspect consist of a cylinder along Y-axis and the Box along X-axis
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim iOutput     As Double
    
    Dim ObjVerCylinIns As Object
    Dim ObjHoriBoxIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)   'D2
    parFace2toCenter = arrayOfInputs(3)  'D3
    parInstrumentHeight = arrayOfInputs(4)   'D4
    parInstrumentHeight1 = arrayOfInputs(5)      'D5
    parInstrumentDiameter = arrayOfInputs(6)      'D6
    parInstrumentHeight2 = arrayOfInputs(7)      'D7
    parInstrumentHeight3 = arrayOfInputs(8)      'D8
    parInstrumentWidth = arrayOfInputs(9)        'D9
    parInstrumentWidth1 = arrayOfInputs(10)      'D10
    parInsulationThickness = arrayOfInputs(11)
    
    iOutput = 0

 ' Insert your code for output 3(Horizontal Box Insulation)
    Dim startPoint       As AutoMath.DPosition
    Dim endPoint         As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set -parFace1toCenter - parInsulationThickness, _
                            -parInstrumentHeight2 - parInsulationThickness, _
                            parInstrumentWidth + parInsulationThickness
    endPoint.Set parFace2toCenter + parInsulationThickness, _
                            parInstrumentHeight3 + parInsulationThickness, _
                            -parInstrumentWidth1 - parInsulationThickness
    Set ObjHoriBoxIns = PlaceBox(m_OutputColl, startPoint, endPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriBoxIns
    Set ObjHoriBoxIns = Nothing
    
' Insert your code for output 4(Vertical Cylinder Insulation)
    startPoint.Set 0, -parInstrumentHeight - parInsulationThickness, 0
    endPoint.Set 0, parInstrumentHeight1 - parInstrumentHeight + parInsulationThickness, 0
    Set ObjVerCylinIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, _
                                                            parInstrumentDiameter + 2 * parInsulationThickness, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerCylinIns
    Set ObjVerCylinIns = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
