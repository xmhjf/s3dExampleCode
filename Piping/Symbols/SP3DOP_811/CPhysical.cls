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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Monday, Dec 24 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Page No. D-145 of PDS Piping Component Data Reference Guide.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   27.Aug.2003     MS/ACM TR47240:  Operator stem diameters are made proportionate so
'                                             that they donot appear for small valve sizes
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   22.Feb.2007         KKC           TR-112500 Maintenance envelopes for some valve operators do not use capped ends
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
'   Dim objCone As Object
'   Dim ObjLHSStem As Object
    Dim ObjLHSWt As Object
'   Dim ObjRHSUpStem As Object
    Dim ObjRHSUpWt As Object
'   Dim ObjRHSLoStem As Object
    Dim ObjRHSLoWt As Object
    
    Dim parLHSStemDIa As Double
    Dim parRHSUpStemDia As Double
    Dim parRHSLoStemDia As Double
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parCounterWtLeverLen1 As Double
    Dim parCounterWtBoxTk1 As Double
    Dim parOperatorHeight1 As Double
    Dim parCounterWtBoxHt1 As Double
    Dim parCounterWtBoxDia1 As Double
    Dim parCounterWtLeverLen2 As Double
    Dim parCounterWtBoxTk2 As Double
    Dim parOperatorHeight2 As Double
    Dim parCounterWtBoxHt2 As Double
    Dim parCounterWtBoxDia2 As Double
    Dim parOperatorHeight3 As Double
    Dim parCounterWtLeverLen3 As Double
    Dim parCounterWtBoxTk3 As Double
    Dim parCounterWtBoxHt3 As Double
    Dim parCounterWtBoxDia3 As Double
    
'   Variables defined by svsmylav
    Dim CenterPos           As New AutoMath.DPosition
    Dim ConeCenterBasePt    As New AutoMath.DPosition
    Dim ConeCenterTopPt     As New AutoMath.DPosition
    Dim LHSStemStPt         As New AutoMath.DPosition
    Dim LHSStemEndPt        As New AutoMath.DPosition
    Dim RHSUpStemStPt       As New AutoMath.DPosition
    Dim RHSUpStemEndPt      As New AutoMath.DPosition
    Dim RHSLoStemStPt       As New AutoMath.DPosition
    Dim RHSLoStemEndPt      As New AutoMath.DPosition
    Dim LHSWtStCorner       As New AutoMath.DPosition
    Dim LHSWtOppCorner      As New AutoMath.DPosition
    Dim RHSUpWtStCorner     As New AutoMath.DPosition
    Dim RHSUpWtOppCorner    As New AutoMath.DPosition
    Dim RHSLoWtStCorner     As New AutoMath.DPosition
    Dim RHSLoWtOppCorner    As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parCounterWtLeverLen1 = arrayOfInputs(4)
    parCounterWtBoxTk1 = arrayOfInputs(5)
    parOperatorHeight1 = arrayOfInputs(6)
    parCounterWtBoxHt1 = arrayOfInputs(7)
    parCounterWtBoxDia1 = arrayOfInputs(8)
    parCounterWtLeverLen2 = arrayOfInputs(9)
    parCounterWtBoxTk2 = arrayOfInputs(10)
    parOperatorHeight2 = arrayOfInputs(11)
    parCounterWtBoxHt2 = arrayOfInputs(12)
    parCounterWtBoxDia2 = arrayOfInputs(13)
    parOperatorHeight3 = arrayOfInputs(14)
    parCounterWtLeverLen3 = arrayOfInputs(15)
    parCounterWtBoxTk3 = arrayOfInputs(16)
    parCounterWtBoxHt3 = arrayOfInputs(17)
    parCounterWtBoxDia3 = arrayOfInputs(18)
    
'   Initialize
    iOutput = 0
    
'    parLHSStemDIa = 0.05
'    parRHSUpStemDia = 0.05
'    parRHSLoStemDia = 0.05

    parLHSStemDIa = parOperatorDiameter * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parLHSStemDIa < 0.015 Then parLHSStemDIa = 0.015
   If parLHSStemDIa > 0.05 Then parLHSStemDIa = 0.05
   
    parRHSUpStemDia = parLHSStemDIa
    parRHSLoStemDia = parLHSStemDIa

'   Prepare data points
    CenterPos.Set 0, 0, 0
    
    ConeCenterTopPt.Set CenterPos.x, CenterPos.y, CenterPos.z
    ConeCenterBasePt.Set CenterPos.x, _
                         CenterPos.y + parOperatorHeight, _
                         CenterPos.z
    
    LHSStemStPt.Set CenterPos.x, _
                    ConeCenterBasePt.y - parOperatorHeight1 + parLHSStemDIa / 2, _
                    CenterPos.z
    LHSStemEndPt.Set LHSStemStPt.x - parCounterWtLeverLen1, _
                     LHSStemStPt.y, _
                     LHSStemStPt.z
    
    RHSUpStemStPt.Set CenterPos.x, _
                      ConeCenterBasePt.y - parOperatorHeight2 + parRHSUpStemDia / 2, _
                      CenterPos.z
    RHSUpStemEndPt.Set CenterPos.x + parCounterWtLeverLen2, _
                       RHSUpStemStPt.y, _
                       RHSUpStemStPt.z
    
    RHSLoStemStPt.Set CenterPos.x, _
                      ConeCenterBasePt.y - parOperatorHeight3 - parRHSLoStemDia / 2, _
                      CenterPos.z
    RHSLoStemEndPt.Set CenterPos.x + parCounterWtLeverLen3, _
                       RHSLoStemStPt.y, _
                       CenterPos.z
    
    LHSWtStCorner.Set LHSStemEndPt.x, LHSStemEndPt.y - parLHSStemDIa / 2, _
                      LHSStemEndPt.z - parCounterWtBoxDia1 / 2
    LHSWtOppCorner.Set LHSWtStCorner.x - parCounterWtBoxTk1, _
                       LHSWtStCorner.y + parCounterWtBoxHt1, _
                       LHSWtStCorner.z + parCounterWtBoxDia1
    
    RHSUpWtStCorner.Set RHSUpStemEndPt.x, RHSUpStemEndPt.y - parRHSUpStemDia / 2, _
                        RHSUpStemEndPt.z - parCounterWtBoxDia2
    RHSUpWtOppCorner.Set RHSUpWtStCorner.x + parCounterWtBoxTk2, _
                         RHSUpWtStCorner.y + parCounterWtBoxHt2, _
                         RHSUpWtOppCorner.z + parCounterWtBoxDia2
    
    RHSLoWtStCorner.Set RHSLoStemEndPt.x, RHSLoStemEndPt.y + parRHSLoStemDia / 2, _
                        RHSLoStemEndPt.z - parCounterWtBoxDia3 / 2
    RHSLoWtOppCorner.Set RHSLoWtStCorner.x + parCounterWtBoxTk3, _
                         RHSLoWtStCorner.y - parCounterWtBoxHt3, _
                         RHSLoWtStCorner.z + parCounterWtBoxDia3

 ' Insert your code for output 1(Cone)
    Dim objCone     As IngrGeom3D.Cone3d
    Set objCone = PlaceCone(m_OutputColl, ConeCenterBasePt, ConeCenterTopPt, parOperatorDiameter, 0, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
    Set objCone = Nothing
    
 ' Insert your code for output 2(Left Hand Side Stem)
    Dim ObjLHSStem As IngrGeom3D.Projection3d
    Set ObjLHSStem = PlaceCylinder(m_OutputColl, LHSStemStPt, LHSStemEndPt, parLHSStemDIa, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLHSStem
    Set ObjLHSStem = Nothing
    
 ' Insert your code for output 3(Left Hand Side Weight)
    Set ObjLHSWt = PlaceBox(m_OutputColl, LHSWtStCorner, LHSWtOppCorner)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLHSWt
    Set ObjLHSWt = Nothing
    
 ' Insert your code for output 4(Right Hand Side Upper Stem)
    Dim ObjRHSUpStem As IngrGeom3D.Projection3d
    Set ObjRHSUpStem = PlaceCylinder(m_OutputColl, RHSUpStemStPt, RHSUpStemEndPt, parRHSUpStemDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRHSUpStem
    Set ObjRHSUpStem = Nothing
    
 ' Insert your code for output 5(RH Side Upper Weight)
    Set ObjRHSUpWt = PlaceBox(m_OutputColl, RHSUpWtStCorner, RHSUpWtOppCorner)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRHSUpWt
    Set ObjRHSUpWt = Nothing
    
 ' Insert your code for output 6(RH Side Lower Stem)
    Dim ObjRHSLoStem As IngrGeom3D.Projection3d
    Set ObjRHSLoStem = PlaceCylinder(m_OutputColl, RHSLoStemStPt, RHSLoStemEndPt, parRHSLoStemDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRHSLoStem
    Set ObjRHSLoStem = Nothing
    
 ' Insert your code for output 7(RH SIde Lower Weight)
    Set ObjRHSLoWt = PlaceBox(m_OutputColl, RHSLoWtStCorner, RHSLoWtOppCorner)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRHSLoWt
    Set ObjRHSLoWt = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
