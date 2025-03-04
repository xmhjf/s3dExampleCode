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
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IREG1,
'   PDMS Category Name=IVRF). Physical aspect consist of Two cones one converging other divrging
'   and a vertical Cylinder with varying diameter, two nozzles one facing -ive direction direction of
'   X-axis and the other in +ive direction of X-axis.
'   Insulation aspect consist of one Cylinder along X-axis and another cylinder along Y-axis
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
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam1      As Double
    Dim depth1           As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ObjConeIns As Object
    Dim ObjRegVerCylinIns As Object
    Dim parFace1toCenter As Double
    Dim parFaceToFace As Double
    Dim parCylHeight As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInsulationThickness As Double
    
    Dim iOutput     As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFaceToFace = arrayOfInputs(3)        'D3
    parCylHeight = arrayOfInputs(4)   'D4
    parInstrumentHeight = arrayOfInputs(5)      'D5
    parInstrumentHeight1 = arrayOfInputs(6)  'D6
    parInstrumentHeight2 = arrayOfInputs(7)      'D7
    parInstrumentDiameter = arrayOfInputs(8)     'D8
    parInstrumentDiameter1 = arrayOfInputs(9)        'D9
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0

 ' Insert your code for output 6(Horizontal Cylinder Insulation)
    Dim startPoint       As AutoMath.DPosition
    Dim endPoint         As AutoMath.DPosition
    Dim insulationDia As Double
    Dim MaxDia As Double
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset2, depth2

    MaxDia = pipeDiam1
    If MaxDia < pipeDiam2 Then MaxDia = pipeDiam2
    If MaxDia < flangeDiam1 Then MaxDia = flangeDiam1
    If MaxDia < flangeDiam2 Then MaxDia = flangeDiam2
        
    insulationDia = MaxDia + 2 * parInsulationThickness
    startPoint.Set -parFace1toCenter, 0, 0
    endPoint.Set parFaceToFace - parFace1toCenter, 0, 0
    Set ObjConeIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, _
                                                            insulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConeIns
    Set ObjConeIns = Nothing
    
' Insert your code for output 6(Vertical Cylinder )
   'Bottom Cylinder Diameter is taken as equal to Pipe 1 Diameter as per EDEN
   Dim CylDia As Double
    CylDia = pipeDiam1
    MaxDia = parInstrumentDiameter1
    If MaxDia < parInstrumentDiameter Then MaxDia = parInstrumentDiameter
    If MaxDia < CylDia Then MaxDia = CylDia
    startPoint.Set 0, -parCylHeight - parInsulationThickness, 0
    endPoint.Set 0, parInstrumentHeight + parInsulationThickness, 0
    insulationDia = MaxDia + 2 * parInsulationThickness
    Set ObjRegVerCylinIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, insulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRegVerCylinIns
    Set ObjRegVerCylinIns = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
