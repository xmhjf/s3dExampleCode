VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:00:52 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          CYW
'   Creation Date:  Thursday, Mar 27 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjPyra     As Object
    Dim ObjBox      As Object
    Dim ObjCyli     As Object
    Dim ObjDish     As Object
    
    Dim strOrigin   As String
    Dim oriOrientation As Orientation
    Dim dblXBottom  As Double
    Dim dblYBottom  As Double
    Dim dblXTop     As Double
    Dim dblYTop     As Double
    Dim dblXOffset  As Double
    Dim dblYOffset  As Double
    Dim dblHeight   As Double
    Dim dblXLength  As Double
    Dim dblYLength  As Double
    Dim dblZLength  As Double
    Dim dblDiameter As Double
    Dim dblLength   As Double
    Dim strDirection As String
    Dim dblKnuckleRadius As Double
        
    Set oriOrientation = New Orientation
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    
    iOutput = 0
    
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutZ = 90
    oriOrientation.ApplyRotations
    
' Insert your code for output 1(Pyra)
    strOrigin = "E 0 S 0.0190 U 0.9398"
    dblXBottom = 3.6322
    dblYBottom = 1.9812
    dblXTop = 3.556
    dblYTop = 1.9812
    dblXOffset = 0.0381
    dblYOffset = 0
    dblHeight = 0.1524
    Set ObjPyra = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, _
        dblXOffset, dblYOffset, dblHeight)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPyra
    Set ObjPyra = Nothing
    
' Insert your code for output 2(Pyra)
    strOrigin = "W 0.7176 S 0.1841 U 1.3589"
    dblXBottom = 1.2192
    dblYBottom = 0.635
    dblXTop = 0.9144
    dblYTop = 0.4572
    dblXOffset = 0.0508
    dblYOffset = -0.0889
    dblHeight = 0.6865
    Set ObjPyra = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, _
        dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPyra
    Set ObjPyra = Nothing

' Insert your code for output 3(Pyra)
    strOrigin = "W 0.6795 N 0.4921 U 1.0668"
    dblXBottom = 0.1016
    dblYBottom = 0.635
    dblXTop = 0.1302
    dblYTop = 0.6096
    dblXOffset = 0.0143
    dblYOffset = -0.0127
    dblHeight = 0.1016
    Set ObjPyra = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, _
        dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPyra
    Set ObjPyra = Nothing

' Insert your code for output 4(Pyra)
    strOrigin = "W 0.9398 N 1.5621 U 0.6096"
    dblXBottom = 0.254
    dblYBottom = 0.1016
    dblXTop = 0.508
    dblYTop = 0.1016
    dblXOffset = 0
    dblYOffset = 0
    dblHeight = 0.508
    Set ObjPyra = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, _
        dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPyra
    Set ObjPyra = Nothing

' Insert your code for output 5(Pyra)
    strOrigin = "E 0 N 1.5621 U 0.4826"
    dblXBottom = 0.254
    dblYBottom = 1.778
    dblXTop = 0.381
    dblYTop = 1.778
    dblXOffset = 0
    dblYOffset = 0
    dblHeight = 0.254
    Set ObjPyra = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, _
        dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPyra
    Set ObjPyra = Nothing

' Insert your code for output 6(Box)
    strOrigin = "E 0 N 1.5621 U 0.7366"
    dblXLength = 0.381
    dblYLength = 1.778
    dblZLength = 0.254
    Set ObjBox = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox
    Set ObjBox = Nothing

' Insert your code for output 7(Pyra)
    strOrigin = "E 0.9398 N 1.5621 U .6096"
    dblXBottom = 0.254
    dblYBottom = 0.1016
    dblXTop = 0.508
    dblYTop = 0.1016
    dblXOffset = 0
    dblYOffset = 0
    dblHeight = 0.508
    Set ObjPyra = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, _
        dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPyra
    Set ObjPyra = Nothing

' Insert your code for output 8(Pyra)
    strOrigin = "E 0 S 0.0698 U 0.6096"
    dblXBottom = 0.9906
    dblYBottom = 1.9812
    dblXTop = 1.2954
    dblYTop = 1.9812
    dblXOffset = 0
    dblYOffset = 0
    dblHeight = 0.508
    Set ObjPyra = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, _
        dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPyra
    Set ObjPyra = Nothing

' Insert your code for output 9(Pyra)
    strOrigin = "E 0 S 1.6637 U 0.6096"
    dblXBottom = 0.2286
    dblYBottom = 1.9812
    dblXTop = 0.381
    dblYTop = 1.9812
    dblXOffset = 0.0762
    dblYOffset = 0
    dblHeight = 0.508
    Set ObjPyra = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, _
        dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPyra
    Set ObjPyra = Nothing

' Insert your code for output 10(Pyra)
    strOrigin = "E 0 N 1.7336 U 0.4318"
    dblXBottom = 0.1016
    dblYBottom = 0.0254
    dblXTop = 0
    dblYTop = 0.0254
    dblXOffset = 0.0127
    dblYOffset = 0
    dblHeight = 0.1524
    Set ObjPyra = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, _
        dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPyra
    Set ObjPyra = Nothing

' Insert your code for output 11(Cyli)
    strOrigin = "E 0 N 1.7589 U 0.3747"
    dblDiameter = 0.0254
    dblLength = 0.0317
    strDirection = "E 90 U 0"
    Set ObjCyli = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCyli
    Set ObjCyli = Nothing

' Insert your code for output 12(Cyli)
    strOrigin = "E 0.8636 S 1.0668 U 0.4191"
    dblDiameter = 0.8382
    dblLength = 0.2032
    strDirection = "E 90 U 0"
    Set ObjCyli = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCyli
    Set ObjCyli = Nothing

' Insert your code for output 13(Cyli)
    strOrigin = "E 0.8636 N 0.9271 U 0.4191"
    dblDiameter = 0.8382
    dblLength = 0.2032
    strDirection = "E 90 U 0"
    Set ObjCyli = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCyli
    Set ObjCyli = Nothing

' Insert your code for output 14(Cyli)
    strOrigin = "W 0.8636 N 0.9271 U .4191"
    dblDiameter = 0.8382
    dblLength = 0.2032
    strDirection = "E 90 U 0"
    Set ObjCyli = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCyli
    Set ObjCyli = Nothing

' Insert your code for output 15(Cyli)
    strOrigin = "W 0.8636 S 1.0668 U 0.4191"
    dblDiameter = 0.8382
    dblLength = 0.2032
    strDirection = "E 90 U 0"
    Set ObjCyli = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCyli
    Set ObjCyli = Nothing

' Insert your code for output 16(Dish)
    strOrigin = "E 0.720 N 1.7526 U 0.7620"
    strDirection = "N 0 U 0"
    dblDiameter = 0.1016
    dblHeight = 0.0254
    dblKnuckleRadius = 0
    Set ObjDish = CreateDish(m_outputColl, strOrigin, strDirection, dblDiameter, dblHeight, dblKnuckleRadius, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish
    Set ObjDish = Nothing

' Insert your code for output 17(Dish)
    strOrigin = "E 0.7620 N 1.7526 U 0.6604"
    strDirection = "N 0 U 0"
    dblDiameter = 0.0762
    dblHeight = 0.0254
    dblKnuckleRadius = 0
    Set ObjDish = CreateDish(m_outputColl, strOrigin, strDirection, dblDiameter, dblHeight, dblKnuckleRadius, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish
    Set ObjDish = Nothing

' Insert your code for output 18(Dish)
    strOrigin = "W 0.7620 N 1.7526 U 0.6604"
    strDirection = "N 0 U 0"
    dblDiameter = 0.0762
    dblHeight = 0.0254
    dblKnuckleRadius = 0
    Set ObjDish = CreateDish(m_outputColl, strOrigin, strDirection, dblDiameter, dblHeight, dblKnuckleRadius, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish
    Set ObjDish = Nothing

' Insert your code for output 19(Dish)
    strOrigin = "E 0.6096 N 1.7526 U 0.7620"
    strDirection = "N 0 U 0"
    dblDiameter = 0.1016
    dblHeight = 0.0254
    dblKnuckleRadius = 0
    Set ObjDish = CreateDish(m_outputColl, strOrigin, strDirection, dblDiameter, dblHeight, dblKnuckleRadius, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish
    Set ObjDish = Nothing

' Insert your code for output 20(Dish)
    strOrigin = "W 0.6096 N 1.7526 U 0.7620"
    strDirection = "N 0 U 0"
    dblDiameter = 0.1016
    dblHeight = 0.0254
    dblKnuckleRadius = 0
    Set ObjDish = CreateDish(m_outputColl, strOrigin, strDirection, dblDiameter, dblHeight, dblKnuckleRadius, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish
    Set ObjDish = Nothing

 ' Insert your code for output 21(Dish)
    strOrigin = "W 0.6096 N 1.7526 U 0.7620"
    strDirection = "S 180 U 0"
    dblDiameter = 0.1016
    dblHeight = 0.0254
    dblKnuckleRadius = 0
    Set ObjDish = CreateDish(m_outputColl, strOrigin, strDirection, dblDiameter, dblHeight, dblKnuckleRadius, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish
    Set ObjDish = Nothing

 ' Insert your code for output 22(Dish)
    strOrigin = "E 0.6096 S 1.8161 U 0.7620"
    strDirection = "S 180 U 0"
    dblDiameter = 0.1016
    dblHeight = 0.0254
    dblKnuckleRadius = 0
    Set ObjDish = CreateDish(m_outputColl, strOrigin, strDirection, dblDiameter, dblHeight, dblKnuckleRadius, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish
    Set ObjDish = Nothing

 ' Insert your code for output 23(Dish)
    strOrigin = "W 0.7620 N 1.7526 U 0.7620"
    strDirection = "N 0 U 0 "
    dblDiameter = 0.1016
    dblHeight = 0.0254
    dblKnuckleRadius = 0
    Set ObjDish = CreateDish(m_outputColl, strOrigin, strDirection, dblDiameter, dblHeight, dblKnuckleRadius, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish
    Set ObjDish = Nothing

 ' Insert your code for output 24(Box)
    strOrigin = "E 0 N 0.4826 U 2.0320"
    dblXLength = 3.683
    dblYLength = 0.2032
    dblZLength = 0.2794
    Set ObjBox = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox
    Set ObjBox = Nothing

 ' Insert your code for output 25(Box)
    strOrigin = "E 0 N 2.4003 U 2.0066"
    dblXLength = 0.1778
    dblYLength = 0.2286
    dblZLength = 0.381
    Set ObjBox = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox
    Set ObjBox = Nothing

 ' Insert your code for output 26(Box)
    strOrigin = "E 0 N 0.8382 U 2.0320"
    dblXLength = 2.8956
    dblYLength = 0.2286
    dblZLength = 0.3048
    Set ObjBox = CreateBox(m_outputColl, strOrigin, oriOrientation, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox
    Set ObjBox = Nothing

' Insert your code for output 27(Pyra)
    strOrigin = "E 0 S 0.9906 U 2.0320"
    dblXBottom = 0.8255
    dblYBottom = 0.9144
    dblXTop = 0.7366
    dblYTop = 0.9144
    dblXOffset = 0.0476
    dblYOffset = 0
    dblHeight = 0.3048
    Set ObjPyra = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, _
        dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPyra
    Set ObjPyra = Nothing

' Insert your code for output 28(Pyra)
    strOrigin = "E 0.1524 S 0.1778 U 1.6510"
    dblXBottom = 0.5842
    dblYBottom = 0.0254
    dblXTop = 0.5842
    dblYTop = 0.0254
    dblXOffset = 0.3556
    dblYOffset = 0
    dblHeight = 1.0668
    Set ObjPyra = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, _
        dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPyra
    Set ObjPyra = Nothing

' Insert your code for output 29(Pyra)
    strOrigin = "W 0.1524 S 0.1778 U 1.6510"
    dblXBottom = 0.5842
    dblYBottom = 0.0254
    dblXTop = 0.5842
    dblYTop = 0.0254
    dblXOffset = 0.3556
    dblYOffset = 0
    dblHeight = 1.0668
    Set ObjPyra = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, _
        dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPyra
    Set ObjPyra = Nothing


 ' Insert your code for output 30(Cyli)
    strOrigin = "E 0 N 0 U 1.0668"
    dblDiameter = 0.6858
    dblLength = 0.1016
    strDirection = "N 0 U 90"
    Set ObjCyli = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCyli
    Set ObjCyli = Nothing

    Set oriOrientation = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
