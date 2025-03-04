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
'   Creation Date:  Tuesday, Sep 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IVM3,
'   PDMS Category Name=IFNE). Physical aspect consist of one horizontal cylinder, one top cylinder
'   along Y axis, a top box and two nozzles one facing -ive direction direction of X-axis and the other in +ive
'   direction of X-axis. Insulation aspect consist of a cylinders along X-axis and a box to cover both
'   top cylinder and top box
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
    Dim pipeDiam1        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim depth           As Double
    Dim pipeDiam2        As Double
    Dim flangeDiam2      As Double
    
    Dim ObjHoriCylinIns As Object
    Dim ObjTopBoxIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFace2toCenter = arrayOfInputs(3) 'D3
    parFlowDiameter = arrayOfInputs(4)  'D4
    parInstrumentHeight = arrayOfInputs(5)  'D5
    parInstrumentHeight1 = arrayOfInputs(6) 'D6
    parInstrumentWidth = arrayOfInputs(7)   'D7
    parInstrumentWidth1 = arrayOfInputs(8)  'D8
    parInstrumentLength = arrayOfInputs(9)  'D9
    parInstrumentLength1 = arrayOfInputs(10) 'D10
    parInsulationThickness = arrayOfInputs(11)

    iOutput = 0

 ' Insert your code for output 6(Horizontal Cylinder Insulation)
    Dim startPoint       As AutoMath.DPosition
    Dim endPoint         As AutoMath.DPosition
    Dim insulationDia As Double
    Dim MaxDia As Double
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
'   Check for max among Pipes, flanges and parFlowDiameter values
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    
    MaxDia = pipeDiam1
    If MaxDia < pipeDiam2 Then MaxDia = pipeDiam2
    If MaxDia < flangeDiam1 Then MaxDia = flangeDiam1
    If MaxDia < flangeDiam2 Then MaxDia = flangeDiam2
    If MaxDia < parFlowDiameter Then MaxDia = parFlowDiameter
    
    insulationDia = MaxDia + 2 * parInsulationThickness
    startPoint.Set -parFace1toCenter - parInsulationThickness, 0, 0
    endPoint.Set parFace2toCenter + parInsulationThickness, 0, 0
    Set ObjHoriCylinIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, _
                                                            insulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylinIns
    Set ObjHoriCylinIns = Nothing
    
' Insert your code for output 6(Top Cylinder and Top Box Insulation )
    startPoint.Set -parInstrumentLength - parInsulationThickness, _
                    parFlowDiameter / 2, _
                    -parInstrumentWidth1 - parInsulationThickness
    endPoint.Set parInstrumentLength1 + parInsulationThickness, _
                 parInstrumentHeight + parInsulationThickness, _
                 parInstrumentWidth + parInsulationThickness
    Set ObjTopBoxIns = PlaceBox(m_OutputColl, startPoint, endPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopBoxIns
    Set ObjTopBoxIns = Nothing
    
    Set endPoint = Nothing
    Set startPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
