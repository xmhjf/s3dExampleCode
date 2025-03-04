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
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:        ACM
'   Creation Date:  Thursday, Dec 19 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   18.Jun.2004     MS    (RI 32889)  Changed input parameter name from
'                                           HandWheelDia to HandWheelDiameter(as an existing interface can be used instead of creating a new interface for HandWheelDia)
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

''  Operator Floor stand symbol is prepared based on SP3D Concept design document
''  "Valve Accessories"(valveaccessoriescd.doc).
''  This symbol consists of 4 inputs and 4 outputs all are physical.

Option Explicit

Dim PI As Double

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
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
    
    Dim ObjHandWheelTorus As Object
    Dim ObjSpokeCyl1 As Object
    Dim ObjSpokeCyl2 As Object
    Dim ObjBodyRevolution As Object
    
    Dim parOperatorHeight As Double
    Dim parOperatorHeight1 As Double
    Dim parHandWheelDiameter As Double
    Dim parBasePlateDia As Double

'' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorHeight1 = arrayOfInputs(3)
    parHandWheelDiameter = arrayOfInputs(4)
    parBasePlateDia = arrayOfInputs(5)
 
    
    iOutput = 0
  
' Insert your code for output 4(Hand Wheel Body Torus)
   ''Assumed base Plate thickness is 20mm.
   ''Assumed Hand Wheel Cross-section Circuar diameter as 10mm(thickness)
   ''Assumed gap between handweel and top of operator cylinder as 20mm
   ''Assumed Operator Body Cylinder dia as half of the HandWheel Diameter.
   ''Assumed Connecting cylider of Hand wheel and Body, dia as  one forth of handwheel diameter.
   
   Dim stPoint As New AutoMath.DPosition
   Dim enPoint As New AutoMath.DPosition
   
   Dim dBasePlateThickness As Double
   Dim dHandWheelthickness  As Double
   Dim dHandWheelCylLen As Double
      
   dBasePlateThickness = 0.02
   dHandWheelthickness = 0.01
   dHandWheelCylLen = 0.02

   Dim dStemDia As Double
   dStemDia = 0.005
 
   Dim oTorusCentre As New AutoMath.DPosition
   Dim oTorusAxis As New AutoMath.DVector
   Dim dHWminorRad As Double
   
   dHWminorRad = 0.005
   
   oTorusCentre.Set 0, parOperatorHeight - dHandWheelthickness / 2, 0
   oTorusAxis.Set 0, 1, 0
   
   Set ObjHandWheelTorus = PlaceTorus(m_OutputColl, oTorusCentre, oTorusAxis, _
                                            parHandWheelDiameter / 2, dHWminorRad)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheelTorus
    Set ObjHandWheelTorus = Nothing

 ' Insert your code for output 2(Spoke Cylinder 1)
   '' Assumed that spoke diameter is 5mm
   Dim dSpokeDia As Double
   dSpokeDia = 0.005
   
   stPoint.Set 0, parOperatorHeight - dHandWheelthickness / 2, parHandWheelDiameter / 2
   enPoint.Set 0, parOperatorHeight - dHandWheelthickness / 2, -parHandWheelDiameter / 2
   
   Set ObjSpokeCyl1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dSpokeDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSpokeCyl1
    Set ObjSpokeCyl1 = Nothing
    
 ' Insert your code for output 3(Spoke Cylinder 2)
   stPoint.Set parHandWheelDiameter / 2, parOperatorHeight - dHandWheelthickness / 2, 0
   enPoint.Set -parHandWheelDiameter / 2, parOperatorHeight - dHandWheelthickness / 2, 0
   
   Set ObjSpokeCyl2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dSpokeDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSpokeCyl2
    Set ObjSpokeCyl2 = Nothing
    
' ' Insert your code for output 4(Body by Revolution of complex str)
   
   Dim dRevPts(0 To 29) As Double
   Dim oLine As IngrGeom3D.Line3d
   Dim oComplexStr As IngrGeom3D.ComplexString3d
   Dim oGeomFactory As New IngrGeom3D.GeometryFactory
   Dim oAxisVec As New AutoMath.DVector
   Dim oCollection As New Collection
   
   ''Body points.
   dRevPts(0) = 0
   dRevPts(1) = parOperatorHeight - dHandWheelthickness / 2 + dSpokeDia / 2
   dRevPts(2) = 0
   
   dRevPts(3) = -(0.25 * parHandWheelDiameter) / 2
   dRevPts(4) = parOperatorHeight - dHandWheelthickness / 2 + dSpokeDia / 2
   dRevPts(5) = 0
   
   dRevPts(6) = -(0.25 * parHandWheelDiameter) / 2
   dRevPts(7) = parOperatorHeight - dHandWheelthickness - dHandWheelCylLen
   dRevPts(8) = 0
   
   dRevPts(9) = -(0.5 * parHandWheelDiameter) / 2
   dRevPts(10) = parOperatorHeight - dHandWheelthickness - dHandWheelCylLen
   dRevPts(11) = 0
   
   dRevPts(12) = -(0.5 * parHandWheelDiameter) / 2
   dRevPts(13) = parOperatorHeight - parOperatorHeight1 + dBasePlateThickness
   dRevPts(14) = 0
   
   dRevPts(15) = -parBasePlateDia / 2
   dRevPts(16) = parOperatorHeight - parOperatorHeight1 + dBasePlateThickness
   dRevPts(17) = 0
   
   dRevPts(18) = -parBasePlateDia / 2
   dRevPts(19) = parOperatorHeight - parOperatorHeight1
   dRevPts(20) = 0
   
   dRevPts(21) = -dStemDia / 2
   dRevPts(22) = parOperatorHeight - parOperatorHeight1
   dRevPts(23) = 0
   
   dRevPts(24) = -dStemDia / 2
   dRevPts(25) = 0
   dRevPts(26) = 0
   
   dRevPts(27) = 0
   dRevPts(28) = 0
   dRevPts(29) = 0
   
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dRevPts(0), dRevPts(1), dRevPts(2), _
                                                        dRevPts(3), dRevPts(4), dRevPts(5))
   oCollection.Add oLine
   
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dRevPts(3), dRevPts(4), dRevPts(5), _
                                                        dRevPts(6), dRevPts(7), dRevPts(8))
   oCollection.Add oLine
   
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dRevPts(6), dRevPts(7), dRevPts(8), _
                                                        dRevPts(9), dRevPts(10), dRevPts(11))
   oCollection.Add oLine
      
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dRevPts(9), dRevPts(10), dRevPts(11), _
                                                        dRevPts(12), dRevPts(13), dRevPts(14))
   oCollection.Add oLine
      
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dRevPts(12), dRevPts(13), dRevPts(14), _
                                                        dRevPts(15), dRevPts(16), dRevPts(17))
   oCollection.Add oLine
   
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dRevPts(15), dRevPts(16), dRevPts(17), _
                                                        dRevPts(18), dRevPts(19), dRevPts(20))
   oCollection.Add oLine
   
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dRevPts(18), dRevPts(19), dRevPts(20), _
                                                        dRevPts(21), dRevPts(22), dRevPts(23))
   oCollection.Add oLine
   
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dRevPts(21), dRevPts(22), dRevPts(23), _
                                                        dRevPts(24), dRevPts(25), dRevPts(26))
   oCollection.Add oLine
   
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dRevPts(24), dRevPts(25), dRevPts(26), _
                                                        dRevPts(27), dRevPts(28), dRevPts(29))
   oCollection.Add oLine
   
   stPoint.Set dRevPts(0), dRevPts(1), dRevPts(2)
   Set oComplexStr = PlaceTrCString(stPoint, oCollection)
      
   oAxisVec.Set 0, 1, 0
   stPoint.Set 0, parOperatorHeight / 2, 0
 
   Set ObjBodyRevolution = PlaceRevolution(m_OutputColl, oComplexStr, oAxisVec, stPoint, 2 * PI, True)

'  Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyRevolution
    Set ObjBodyRevolution = Nothing
    
    Dim icount As Integer
    For icount = 1 To oCollection.Count
        oCollection.Remove 1
    Next icount
    
    Dim oComplexStrRem As IJDObject
    Set oComplexStrRem = oComplexStr
    oComplexStrRem.Remove
    Set oComplexStrRem = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oTorusAxis = Nothing
    Set oTorusCentre = Nothing
    Set oAxisVec = Nothing
    Set oLine = Nothing
    Set oCollection = Nothing
    Set oComplexStr = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

