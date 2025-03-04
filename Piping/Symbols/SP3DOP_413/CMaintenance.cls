VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:          MS
'   Creation Date:  Wednesday, September 22 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   08.22.2004      MS  Added Maintenace Envelope to the operator
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   22.Feb.2007         KKC           TR-112500 Maintenance envelopes for some valve operators do not use capped ends
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjMaintCyl1 As Object
    Dim ObjMaintCyl2 As Object
    Dim ObjMaintCone As Object
    
    Dim parOperatorHeight As Double
    Dim parStemOffset As Double
    Dim parActuatorHeight As Double
    Dim parActuatorDiameter As Double
    
'   Variables defined by svsmylav
    Dim StemStPt            As New AutoMath.DPosition
    Dim StemEndPt           As New AutoMath.DPosition
    Dim StemExtStPt         As New AutoMath.DPosition
    Dim StemExtEndPt        As New AutoMath.DPosition
    Dim ConeCenterBasePt    As New AutoMath.DPosition
    Dim ConeCenterTopPt     As New AutoMath.DPosition
    Dim StemDiameter        As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parStemOffset = arrayOfInputs(3)
    parActuatorHeight = arrayOfInputs(4)
    parActuatorDiameter = arrayOfInputs(5)
    
    iOutput = 0
'    StemDiameter = 0.05

     StemDiameter = parActuatorDiameter * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If StemDiameter < 0.015 Then StemDiameter = 0.015
   If StemDiameter > 0.05 Then StemDiameter = 0.05

 ' Insert your code for output 1(Stem)
    StemStPt.Set CenterPos.x, CenterPos.y - StemDiameter, CenterPos.z
    StemEndPt.Set CenterPos.x, CenterPos.y + parOperatorHeight + 1.5 * StemDiameter, _
                      CenterPos.z
                      
    'Dim ObjMainStem As IngrGeom3D.Projection3d
    Set ObjMaintCyl1 = PlaceCylinder(m_OutputColl, StemStPt, StemEndPt, StemDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl1
    Set ObjMaintCyl1 = Nothing
    
 ' Insert your code for output 2(Stem Extension)
    StemExtStPt.Set CenterPos.x, StemEndPt.y - StemDiameter / 2, CenterPos.z
    StemExtEndPt.Set CenterPos.x + parStemOffset + StemDiameter, _
                           StemExtStPt.y, CenterPos.z

    Set ObjMaintCyl2 = PlaceCylinder(m_OutputColl, StemExtStPt, StemExtEndPt, _
                                            StemDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl2
    Set ObjMaintCyl2 = Nothing
    
 ' Insert your code for output 3(Actuator (Cone))
    ConeCenterTopPt.Set CenterPos.x + parStemOffset, CenterPos.y + parOperatorHeight, _
                        CenterPos.z
    ConeCenterBasePt.Set ConeCenterTopPt.x, ConeCenterTopPt.y, _
                         ConeCenterTopPt.z + parActuatorHeight
                         

   ConeCenterBasePt.Set ConeCenterTopPt.x + 0, ConeCenterTopPt.y, _
                         ConeCenterTopPt.z + parActuatorHeight
    
    Set ObjMaintCone = PlaceCone(m_OutputColl, ConeCenterBasePt, ConeCenterTopPt, _
                             parActuatorDiameter, 0, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCone
    Set ObjMaintCone = Nothing
    Set CenterPos = Nothing
    Set StemStPt = Nothing
    Set StemEndPt = Nothing
    Set StemExtStPt = Nothing
    Set StemExtEndPt = Nothing
    Set ConeCenterBasePt = Nothing
    Set ConeCenterTopPt = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
