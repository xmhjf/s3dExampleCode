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
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          Sundar
'   Creation Date:  Wednesday, May 29 2002
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the catalog available at URL http://www.yardneyfilters.com/basket_specs.htm
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   09.Jul.2003     SymbolTeam(India)   Copyright Information, Header  is added.
'   22.Aug.2005     Sundar(svsmylav)    TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                       cylinder to avoid PlaceCylinder function failure.
'   27.Jan.2006     Sundar(svsmylav)    RI-28367: Deleted reference to 'S3AC or S3AZ' of PDS in the
'                                       revision history and added hyper link to Yardney's site.
'                                       Also, modified symbol input parameters to match that of the
'                                       manufacturer's skecth.
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
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput             As Double
    Dim ObjInsFilterBody   As Object
    Dim ObjInsPort1         As Object
    Dim ObjInsBody1         As Object
    Dim ObjInsPort2         As Object
    Dim ObjInsBody2         As Object

    Dim parFilterBodyHeight1 As Double
    Dim parNozzleOffset As Double
    Dim parOffset As Double
    Dim parFacetoFace As Double
    Dim parInsulationThickness  As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFilterBodyHeight1 = arrayOfInputs(2) 'A
    parNozzleOffset = arrayOfInputs(3)      'B
    parOffset = arrayOfInputs(4)            'C
    parFacetoFace = arrayOfInputs(5)        'D
    parInsulationThickness = arrayOfInputs(6)

    iOutput = 0
    
'   Assumption: Origin of the symbol is assumed to be at half of the parFilterBodyHeight1

' Insert your code for output 1(Insulation for Filetr  Body)
    Dim dInsulationDia  As Double
'   Assumption: Diameter of the body of the filter is taken 80% of parFacetoFace
    Dim dFilterBodyDiameter   As Double
    dFilterBodyDiameter = 0.8 * parFacetoFace

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    stPoint.Set 0, -parFilterBodyHeight1 / 2 - parInsulationThickness, 0
    enPoint.Set 0, parFilterBodyHeight1 / 2 + parInsulationThickness, 0
    dInsulationDia = dFilterBodyDiameter + 2 * parInsulationThickness
    Set ObjInsFilterBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsFilterBody
    Set ObjInsFilterBody = Nothing
    
' Insert your code for output 2(Insulation Port1 Side)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    
    stPoint.Set -parFacetoFace / 2, _
                (parOffset + parNozzleOffset) - parFilterBodyHeight1 / 2, 0
    enPoint.Set stPoint.x + flangeThick + parInsulationThickness, stPoint.y, stPoint.z
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDia = flangeDiam + 2 * parInsulationThickness
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing
    
' Insert your code for output 3(Insulation for body Port1 Side)
    stPoint.Set -parFacetoFace / 2 + flangeThick, _
                (parOffset + parNozzleOffset) - parFilterBodyHeight1 / 2, 0
    enPoint.Set 0, stPoint.y, stPoint.z
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjInsBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody1
    Set ObjInsBody1 = Nothing
    
' Insert your code for output 4(Insulation Port2 Side)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001

    stPoint.Set parFacetoFace / 2, _
                    parOffset - parFilterBodyHeight1 / 2, 0
    enPoint.Set stPoint.x - flangeThick - parInsulationThickness, stPoint.y, stPoint.z
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDia = flangeDiam + 2 * parInsulationThickness
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing
    
' Insert your code for output 5(Insulation for body Port2 Side)
    stPoint.Set parFacetoFace / 2 - flangeThick, _
                    parOffset - parFilterBodyHeight1 / 2, 0
    enPoint.Set 0, stPoint.y, stPoint.z
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjInsBody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody2
    Set ObjInsBody2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
