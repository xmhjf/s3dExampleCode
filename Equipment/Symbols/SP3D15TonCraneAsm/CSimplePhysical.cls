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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:00:22 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          HL
'   Creation Date:  Thursday, Dec 5 2002
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

    Dim CenterPos       As IJDPosition
    Dim startPos        As IJDPosition
    Dim endPos          As IJDPosition
        
    Set CenterPos = New DPosition
    Set startPos = New DPosition
    Set endPos = New DPosition
    
    CenterPos.Set 0, 0, 0
    
    Dim iOutput         As Double
    Dim ObjBody         As Object
    Dim iCount          As Integer
    
    Dim oriOrientation As Orientation
    Set oriOrientation = New Orientation

' Inputs
    Set oPartFclt = arrayOfInputs(1)

    iOutput = 0

' Insert your code for output 1(BodyBox)
    Set ObjBody = CreateBox(m_outputColl, "W 0.0254 N 0 u 1.09696", oriOrientation, 5.8166, 1.2192, 0.4572, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code for output 2(BodyFrontPyramid - Driver Cab)
    Set ObjBody = CreatePyramid(m_outputColl, "W 1.82245 N 0 U 1.78276", Nothing, 1.3716, 1.1176, 0.9906, 1.1176, 0.1905, 0, 0.9144)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

' Insert your code of output 3(BodyBackPyramid - Weight)
    Set ObjBody = CreatePyramid(m_outputColl, "E 2.4511 N 0 U 1.63036", Nothing, 1.1176, 1.1176, 1.1176, 1.1176, 0.2286, 0, 0.6096)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code for output 4(BodyCenterCylinder)
    startPos.Set CenterPos.x, CenterPos.y, CenterPos.z + 1.32556
    endPos.Set CenterPos.x, CenterPos.y, CenterPos.z + 1.42716
    Set ObjBody = PlaceCylinder(m_outputColl, startPos, endPos, 0.9144, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
        
' Insert your code of output 5 (BodySideCylinder 1)
    startPos.Set CenterPos.x - 0.5334, CenterPos.y - 0.92075, CenterPos.z + 1.09696
    endPos.Set CenterPos.x + 0.0762, CenterPos.y - 0.92075, CenterPos.z + 1.09696
    Set ObjBody = PlaceCylinder(m_outputColl, startPos, endPos, 0.4572, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 6 (BodySideCylinder 2)
    startPos.Set CenterPos.x - 0.5334, CenterPos.y + 0.89535, CenterPos.z + 1.09696
    endPos.Set CenterPos.x + 0.0762, CenterPos.y + 0.89535, CenterPos.z + 1.09696
    Set ObjBody = PlaceCylinder(m_outputColl, startPos, endPos, 0.4572, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 7 (BodyWheel 1)
    startPos.Set CenterPos.x - 1.7272, CenterPos.y - 1.3335, CenterPos.z + 0.635
    endPos.Set CenterPos.x - 1.7272, CenterPos.y - 0.8128, CenterPos.z + 0.635
    Set ObjBody = PlaceCylinder(m_outputColl, startPos, endPos, 1.27, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

' Insert your code of output 8 (BodyWheel 2)
    startPos.Set CenterPos.x + 1.3462, CenterPos.y - 1.3335, CenterPos.z + 0.635
    endPos.Set CenterPos.x + 1.3462, CenterPos.y - 0.8128, CenterPos.z + 0.635
    Set ObjBody = PlaceCylinder(m_outputColl, startPos, endPos, 1.27, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 9 (BodyWheel 3)
    startPos.Set CenterPos.x - 1.7272, CenterPos.y + 1.3335, CenterPos.z + 0.635
    endPos.Set CenterPos.x - 1.7272, CenterPos.y + 0.8128, CenterPos.z + 0.635
    Set ObjBody = PlaceCylinder(m_outputColl, startPos, endPos, 1.27, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 10 (BodyWheel 4)
    startPos.Set CenterPos.x + 1.3462, CenterPos.y + 1.3335, CenterPos.z + 0.635
    endPos.Set CenterPos.x + 1.3462, CenterPos.y + 0.8128, CenterPos.z + 0.635
    Set ObjBody = PlaceCylinder(m_outputColl, startPos, endPos, 1.27, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 11 (BodyFrontTriangle)

    Set ObjBody = CreatePyramid(m_outputColl, "W 2.7813 N 0 U 0.85884", Nothing, 0.0254, 2.667, 0.0254, 1.0668, 0, 0, 0.9144)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 12 (BodyFrontTriangle)
    Set ObjBody = CreatePyramid(m_outputColl, "W 2.9337 N 0 U 0.85884", Nothing, 0.0254, 2.667, 0.0254, 1.0668, 0, 0, 0.9144)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 13 (BodyFrontTriangle)
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = 90
    oriOrientation.RotationAboutZ = 180
    oriOrientation.ApplyRotations
    Set ObjBody = CreatePyramid(m_outputColl, "W 2.4892 N 0 U 0.76041", oriOrientation, 0, 1.2192, 0.4572, 1.2192, 0.2286, 0, 0.5588)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 14 (BodyBackTriangle)
    Set ObjBody = CreatePyramid(m_outputColl, "E 2.7305 N 0 U 0.85884", Nothing, 0.0254, 2.667, 0.0254, 1.0668, 0, 0, 0.9144)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 15 (BodyBackTriangle)
    Set ObjBody = CreatePyramid(m_outputColl, "E 2.8829 N 0 U 0.85884", Nothing, 0.0254, 2.667, 0.0254, 1.0668, 0, 0, 0.9144)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 16 (BodyBackTriangle)
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutZ = 180
    oriOrientation.ApplyRotations
    Set ObjBody = CreatePyramid(m_outputColl, "E 2.6035 N 0 U 0.71596", oriOrientation, 0, 1.2192, 0.4572, 0.6096, 0.2286, 0, 0.3048)
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 17 - 28 (BodyWheelWells)

'Front Left Wheel Well (17-19)
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = -60
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "W 2.30346 S 0.97155 U 1.11919", oriOrientation, 0.6096, 0.7239, 0.0762, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

    oriOrientation.ResetDefaultAxis
    Set ObjBody = CreateBox(m_outputColl, "W 1.7272 S 0.97155 U 1.36366", oriOrientation, 0.9144, 0.7239, 0.0762, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = -120
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "W 1.15094 S 0.97155 U 1.11919", oriOrientation, 0.6096, 0.7239, 0.0762, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

'Rear Left Wheel Well (20-22)
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = -60
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "E 0.76994 S 0.97155 U 1.11919", oriOrientation, 0.6096, 0.7239, 0.0762, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

    oriOrientation.ResetDefaultAxis
    Set ObjBody = CreateBox(m_outputColl, "E 1.3462 S 0.97155 U 1.36366", oriOrientation, 0.9144, 0.7239, 0.0762, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = -120
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "E 1.92246 S 0.97155 U 1.11919", oriOrientation, 0.6096, 0.7239, 0.0762, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

'Front Right Wheel Well (23-25)

    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = -60
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "W 2.30346 N 0.97155 U 1.11919", oriOrientation, 0.6096, 0.7239, 0.0762, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

    oriOrientation.ResetDefaultAxis
    Set ObjBody = CreateBox(m_outputColl, "W 1.7272 N 0.97155 U 1.36366", oriOrientation, 0.9144, 0.7239, 0.0762, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = -120
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "W 1.15094 N 0.97155 U 1.11919", oriOrientation, 0.6096, 0.7239, 0.0762, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

'Rear Right WheelWell (26 - 28)
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = -60
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "E 0.76994 N 0.97155 U 1.11919", oriOrientation, 0.6096, 0.7239, 0.0762, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

    oriOrientation.ResetDefaultAxis
    Set ObjBody = CreateBox(m_outputColl, "E 1.3462 N 0.97155 U 1.36366", oriOrientation, 0.9144, 0.7239, 0.0762, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = -120
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "E 1.92246 N 0.97155 U 1.11919", oriOrientation, 0.6096, 0.7239, 0.0762, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 29 - 32 (BodyDishes)
    Set ObjBody = CreateDish(m_outputColl, "E 0.0762 S 0.92075 U 1.09696", "E 90 U 0", 0.4572, 0.1016, 0, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

    Set ObjBody = CreateDish(m_outputColl, "E 0.0762 N 0.89535 U 1.09696", "E 90 U 0", 0.4572, 0.1016, 0, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    Set ObjBody = CreateDish(m_outputColl, "W 0.5334 S 0.92075 U 1.09696", "W 270 U 0", 0.4572, 0.1016, 0, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    Set ObjBody = CreateDish(m_outputColl, "W 0.5334 N 0.89535 U 1.09696", "W 270 U 0", 0.4572, 0.1016, 0, True)

    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' Insert your code of output 33 - 36 (BodyBoxes)
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutX = 120
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "W 2.8575 S 1.62084 U 0.20003", oriOrientation, 0.1524, 0.1524, 0.9144, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutX = -120
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "W 2.8575 N 1.57639 U 0.20003", oriOrientation, 0.1524, 0.1524, 0.9144, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutX = 120
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "E 2.8067 S 1.62084 U 0.20003", oriOrientation, 0.1524, 0.1524, 0.9144, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutX = -120
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "E 2.8067 N 1.57639 U 0.20003", oriOrientation, 0.1524, 0.1524, 0.9144, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

' Insert your code of output 37 (Boom Base Pyramid)
    Set ObjBody = CreatePyramid(m_outputColl, "E 2.3114 N 0 U 2.69716", Nothing, 1.524, 1.1176, 1.8288, 1.1176, 0.1524, 0, 0.4572)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

' Insert your code of output 38 (Boom Base Box)
    oriOrientation.ResetDefaultAxis
    Set ObjBody = CreateBox(m_outputColl, "E 0.8636 N 0.3429 U 2.69716", oriOrientation, 1.2192, 0.381, 0.4572, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
' Insert your code of output 39 (Boom Base Pyramid)
    Set ObjBody = CreatePyramid(m_outputColl, "E 0.4064 N 0 U 1.96056", Nothing, 1.1176, 1.1176, 0.6096, 1.1176, 0.8128, 0, 1.0668)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
 
' Insert your code of output 40 (Boom Box)
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = 30
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "W 1.1128 N 0 U 4.01638", oriOrientation, 6.096, 0.3048, 0.4572, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

' Insert your code of output 41 (Boom Box)
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = 30
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "W 6.39128 N 0 U 7.06438", oriOrientation, 6.2484, 0.3048, 0.381, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
' Insert your code of output 42 (Boom Box)
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = 30
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "W 11.7364 N 0 U 10.15047", oriOrientation, 6.2484, 0.3048, 0.3048, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
' Insert your code of output 43 (Boom Hook Box)
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = 30
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "W 14.678 N 0 U 11.76179", oriOrientation, 0.4572, 0.3048, 0.4572, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
' Insert your code of output 44 (Boom Hook Box)
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = 30
    oriOrientation.ApplyRotations
    Set ObjBody = CreateBox(m_outputColl, "W 14.9066 N 0 U 11.3649", oriOrientation, 0.4572, 0.3048, 0.4572, True)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

' Insert your code of output 45 (Boom Hook)
'    Set ObjBody = CreateCircularTorus(m_outputColl, "W 15.132 N 0 U 11.22998", 0.1016, 0.1524, , "W 15 D", "U 30 E", 45)
    Set ObjBody = CreateCircularTorus(m_outputColl, "W 14.78 N 0 U 11.17", 0.2, 0.25, , "E 90 U 90", "E 90 D -60", 90)
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing

' Insert your code of output 46 (Boom Hook)
'    Set ObjBody = CreateCircularTorus(m_outputColl, "W 15.06538 N 0 U 10.9855", 0.1016, 0.1524, , "D 30 W", "E 15 U", 45)
      Set ObjBody = CreateCircularTorus(m_outputColl, "W 15.1697 N 0 U 10.945", 0.2, 0.25, , "W 270 U 60", "N 0 D -90", 45)
  
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
' Insert your code of output 47 (Boom Hook)
'    Set ObjBody = CreateCircularTorus(m_outputColl, "W 15.06538 N 0 U 10.9855", 0.1016, 0.1524, , "U 30 E", "E 45 U", 45)
      Set ObjBody = CreateCircularTorus(m_outputColl, "W 15.1697 N 0 U 10.945", 0.2, 0.25, , "N 0 U 90", "W 270 U 75", 45)
  
    ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
    Set startPos = Nothing
    Set endPos = Nothing
        
    Exit Sub
    
ErrorLabel:
    logError (Err.Number & ": " & Err.description & " " & Err.Source)
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

