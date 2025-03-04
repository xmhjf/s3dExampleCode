VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Monday, Sep 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ITM, PDMS Category Name=IFNF)
'   The symbol consists of seven outputs, five physical and two insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

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
    
    Dim iOutput     As Double
    Dim ObjTurbineMeterBody As Object
    Dim ObjTurbineVerCyl As Object
    Dim ObjTurbineMeterTopBody As Object
    Dim ObjTurbineMeterBodyIns As Object
    Dim ObjTurbineMeterTopBodyIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentWidth As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFace2toCenter = arrayOfInputs(3)     'D3
    parFlowDiameter = arrayOfInputs(4)      'D4
    parInstrumentHeight = arrayOfInputs(5)  'D5
    parInstrumentDiameter = arrayOfInputs(6)    'D6
    parInstrumentWidth = arrayOfInputs(7)       'D7
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0
'Origin is considered to be along the flow axis and at parFace1toCenter from nozzle 1
 ' Insert your code for output 1(TurbineMeterBody)
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFace2toCenter, 0, 0
    Set ObjTurbineMeterBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parFlowDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTurbineMeterBody
    Set ObjTurbineMeterBody = Nothing
    
 ' Insert your code for output 2(TurbineVerCyl)
    Dim meetHeight As Double
    Dim verCylDia As Double
'Diameter of Vertical cylinder is assumed to be 0.2 times the parInstrumentDiameter
    verCylDia = 0.2 * parInstrumentDiameter
    meetHeight = Sqr((parInstrumentDiameter / 2 * parInstrumentDiameter / 2) - (verCylDia / 2 * verCylDia / 2))
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parInstrumentHeight - meetHeight, 0
    Set ObjTurbineVerCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, verCylDia, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTurbineVerCyl
    Set ObjTurbineVerCyl = Nothing
    
 ' Insert your code for output 3(TurbineMeterTopBody)
    stPoint.Set 0, parInstrumentHeight, -parInstrumentWidth / 2
    enPoint.Set 0, parInstrumentHeight, parInstrumentWidth / 2
    Set ObjTurbineMeterTopBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTurbineMeterTopBody
    Set ObjTurbineMeterTopBody = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
