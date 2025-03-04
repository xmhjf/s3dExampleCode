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
'   Author:          BG
'   Creation Date:  Wednesday, Jun 5 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol consist of Physical and Insulation Aspects. Physical aspect is
'   made up of Two Nozzles with Length, Horizontal Cylinder and Vertical Cylinder
'   Insulation aspect consist of Insulation for both Nozzles, Insulation for the pipe portion in
'   both Nozzles, Insultion for Hori and Verital Cylinder.
'   For more details about this Symbol please refer PDS Piping Component reference Guide
'   Symbol Number SN=S32BC or S32BZ
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
    Dim ObjVerticalCylin As Object
    Dim ObjHoriCylinder As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parCylinderDiameter As Double
    Dim parVertCylinHeightFromCen As Double
    Dim parHoriCylinLengthFromCen As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'P2
    parFace2toCenter = arrayOfInputs(3)     'P3
    parCylinderDiameter = arrayOfInputs(4)     'P4
    parVertCylinHeightFromCen = arrayOfInputs(5)      'P5
    parHoriCylinLengthFromCen = arrayOfInputs(6)      'P6
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0

 ' Insert your code for output 1(Vertical Cylinder)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set 0, 0, -parVertCylinHeightFromCen
    endPoint.Set 0, 0, parCylinderDiameter / 2
    Set ObjVerticalCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, parCylinderDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerticalCylin
    Set ObjVerticalCylin = Nothing
    
 ' Insert your code for output 2(Horizontal Cylinder)
    startPoint.Set 0, 0, 0
    endPoint.Set parHoriCylinLengthFromCen, 0, 0
    Set ObjHoriCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, parCylinderDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylinder
    Set ObjHoriCylinder = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim NozzleLength As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    NozzleLength = (parFace1toCenter - parVertCylinHeightFromCen)
    oPlacePoint.Set 0, 0, -parFace1toCenter - sptOffset + depth
    oDir.Set 0, 0, -1
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    NozzleLength = parFace2toCenter - parHoriCylinLengthFromCen
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
