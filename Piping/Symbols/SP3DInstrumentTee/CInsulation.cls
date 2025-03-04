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
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Friday, Nov 15 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol constructed as per the documents available at URL http://www.resistoflex.com.
'   Physical aspect consist of Tee Body as Cylinder, Two Nozzles and One Nozzle with Length along
'   Y- Axis.
'   Insulation aspect consist of Three Cylinders, one for the Main Body seceond for the Branch Nozzle
'   and the third for the Branch Body
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   22.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeDiam As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjTeeBodyIns As Object
    Dim ObjTeeBranchIns As Object
    Dim ObjTeeBranNozIns As Object
    
    Dim parFacetoFace As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parFace3toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    Dim InstrumentDiameter As Double
    Dim raisedFaceOrSocketDiam As Double
    
    iOutput = 0

 ' Insert your code for output 3(Tee Body Insulation)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Dim insulationDia As Double
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0
    'The assuption is that for end preparations of 100 series Flange Diameter will be zero and we will get
    'a positive value for raised Face Diameter, Symbol created under that condition is Instrument Tee
    'For Lugged body Instrument the body diameter is equal to Flange Diameter, so when we get some positive
    'Value for Flange, raised face diameter will be zero.
    RetrieveParametersForThruBoltedEnds 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                    sptOffset, depth, raisedFaceOrSocketDiam
    If flangeDiam = 0 Then                  'Instrument Tee
        InstrumentDiameter = raisedFaceOrSocketDiam
    Else                                                    'Lugged Body Instrument
        InstrumentDiameter = flangeDiam
    End If
    
    If InstrumentDiameter < pipeDiam Then InstrumentDiameter = pipeDiam
    insulationDia = InstrumentDiameter + parInsulationThickness * 2
    Set ObjTeeBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, insulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTeeBodyIns
    Set ObjTeeBodyIns = Nothing
    
 ' Insert your code for output 4(Tee Branch Nozzle Insulation)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set 0, parFace3toCenter, 0
    enPoint.Set 0, parFace3toCenter - flangeThick - parInsulationThickness, 0
    insulationDia = flangeDiam + 2 * parInsulationThickness
    If pipeDiam > flangeDiam Then insulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjTeeBranNozIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, insulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTeeBranNozIns
    Set ObjTeeBranNozIns = Nothing

' Insert your code for output 4(Tee Branch Body Insulation)
    stPoint.Set 0, parFace3toCenter - flangeThick, 0
    enPoint.Set 0, InstrumentDiameter / 2, 0
    insulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjTeeBranchIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, insulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTeeBranchIns
    Set ObjTeeBranchIns = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
