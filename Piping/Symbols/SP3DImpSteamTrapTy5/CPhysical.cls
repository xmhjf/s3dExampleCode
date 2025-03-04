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
'   Creation Date:  Wednesday, Jun 5 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No. is: S64BC, Page No. D-106 PDS Piping Component Data Reference Guide.
' SP3DImpSteamTrapTy5 Symbol is created with Eight Outputs - Five are Insulation aspect outputs.
' Other three are physical aspect outputs created as follows:
'   ObjTrapBody - Using 'PlaceCylinder' function and
'   two ObjNozzle objects  by using 'CreateNozzleWithLength' function.
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
    Dim ObjTrapBody  As Object
        
    Dim parNozz1toCentre As Double
    Dim parNozz2toCentre As Double
    Dim parOffset1 As Double
    Dim parOffset2 As Double
    Dim parTrapBodyDiameter As Double
    Dim parTrapHeightBottoCen As Double
    Dim parTrapHeightToptoCen As Double
    Dim parInsulationThickness As Double

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parNozz1toCentre = arrayOfInputs(2)  'P2
    parNozz2toCentre = arrayOfInputs(3)  'P3
    parOffset1 = arrayOfInputs(4)       'P4
    parOffset2 = arrayOfInputs(5)       'P5
    parTrapBodyDiameter = arrayOfInputs(6)      'P6
    parTrapHeightBottoCen = arrayOfInputs(7)     'P7
    parTrapHeightToptoCen = arrayOfInputs(8)     'P8
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0

' Insert your code for output 4(Body of Trap)
    stPoint.Set 0, 0, -parTrapHeightBottoCen
    enPoint.Set 0, 0, parTrapHeightToptoCen
    Set ObjTrapBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parTrapBodyDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTrapBody
    Set ObjTrapBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parNozz1toCentre - sptOffset + depth, 0, -parOffset1
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parNozz1toCentre)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set -parNozz2toCentre - sptOffset + depth, 0, parOffset2
    oDir.Set -1, 0, 0

    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, parNozz2toCentre)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

