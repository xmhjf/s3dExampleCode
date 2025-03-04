VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:          Veena
'   Creation Date:  Friday, oct 04 2006
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This Symbol is a replicate of "Tank.TankServices" Symbol, except that the symbol is in line with
' Symbol created using wizard.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Operation:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim ObjSwitchGearBody As Object
    Dim ObjDoorSwing As Object
    
    Dim iOutput     As Double
    Dim parSwitchGearHeight As Double
    Dim parSwitchGearWidth As Double
    Dim parSwitchGearLength As Double
    Dim CenterPos As IJDPosition
    Dim pBodyCorner1   As IJDPosition
    Dim pBodyCorner2   As IJDPosition
    Set CenterPos = New DPosition
    Set pBodyCorner1 = New DPosition
    Set pBodyCorner2 = New DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parSwitchGearHeight = arrayOfInputs(2)
    parSwitchGearWidth = arrayOfInputs(3)
    parSwitchGearLength = arrayOfInputs(4)
    
    iOutput = 0
    
    'Create the SwitchGear Body
     CenterPos.Set 0, 0, 0
    pBodyCorner1.Set CenterPos.x + parSwitchGearLength, CenterPos.y, CenterPos.z
    pBodyCorner2.Set CenterPos.x + parSwitchGearLength + parSwitchGearHeight, CenterPos.y + parSwitchGearWidth, CenterPos.z + parSwitchGearHeight
    Set ObjSwitchGearBody = PlaceBox(m_outputColl, pBodyCorner1, pBodyCorner2)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjSwitchGearBody
    Set ObjSwitchGearBody = Nothing

    'Create 180deg SwitchGear Door swing
    Dim dSweepAngle As Double
    Dim PI As Double
    PI = 4 * Atn(1)
    dSweepAngle = PI
   
    Set ObjDoorSwing = PlaceDoorSwing(m_outputColl, dSweepAngle, parSwitchGearLength, parSwitchGearWidth, parSwitchGearHeight, False)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDoorSwing
    Set ObjDoorSwing = Nothing

    Set ObjDoorSwing = PlaceDoorOpen(m_outputColl, dSweepAngle, parSwitchGearLength, parSwitchGearWidth, parSwitchGearHeight, False)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDoorSwing
    Set ObjDoorSwing = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

'this function places a rectangular torus
Private Function PlaceDoorSwing(oOutputCol As Object, ByVal dSweepAngle As Double, ByVal dSwitchGearLength As Double, ByVal dSwitchGearWidth As Double, ByVal dSwitchGearHeight As Double, ByVal bIsCapped As Boolean) As Object
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory

    Const lNumberOfSides As Long = 4 'four sides for rectangle
    
    Dim dPoints(1 To (lNumberOfSides + 1) * 3) As Double
    
    dPoints(1) = 0
    dPoints(2) = dSwitchGearWidth
    dPoints(3) = dSwitchGearHeight
    
    dPoints(4) = 0
    dPoints(5) = 0.001
    dPoints(6) = dSwitchGearHeight
    
    dPoints(7) = 0
    dPoints(8) = 0.001
    dPoints(9) = 0
    
    dPoints(10) = 0
    dPoints(11) = dSwitchGearWidth
    dPoints(12) = 0
    
    dPoints(13) = dPoints(1)
    dPoints(14) = dPoints(2)
    dPoints(15) = dPoints(3)

    Dim oLS3d As IJLineString
    Set oLS3d = oGeometryFactory.LineStrings3d.CreateByPoints(Nothing, lNumberOfSides + 1, dPoints)
    
    Dim oRectangularTorus As IngrGeom3D.Revolution3d
    Set oRectangularTorus = oGeometryFactory.Revolutions3d.CreateByCurve(oOutputCol.ResourceManager, oLS3d, 0#, 0#, 1#, -dSwitchGearLength, -dSwitchGearWidth / 2, -dSwitchGearHeight / 2, dSweepAngle, bIsCapped)

    Set PlaceDoorSwing = oRectangularTorus
End Function

'this function places a rectangular torus
Private Function PlaceDoorOpen(oOutputCol As Object, ByVal dSweepAngle As Double, ByVal dSwitchGearLength As Double, ByVal dSwitchGearWidth As Double, ByVal dSwitchGearHeight As Double, ByVal bIsCapped As Boolean) As Object
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory

    Const lNumberOfSides As Long = 4 'four sides for rectangle
    
    Dim dPoints(1 To (lNumberOfSides + 1) * 3) As Double
    
    dPoints(1) = 0
    dPoints(2) = 0.001
    dPoints(3) = dSwitchGearHeight
    
    
    dPoints(4) = 0
    dPoints(5) = dSwitchGearWidth
    dPoints(6) = dSwitchGearHeight
    
       
    dPoints(7) = 0
    dPoints(8) = dSwitchGearWidth
    dPoints(9) = 0
    
    dPoints(10) = 0
    dPoints(11) = 0.001
    dPoints(12) = 0
    
    
    dPoints(13) = dPoints(1)
    dPoints(14) = dPoints(2)
    dPoints(15) = dPoints(3)

    Dim oLS3d As IJLineString
    Set oLS3d = oGeometryFactory.LineStrings3d.CreateByPoints(Nothing, lNumberOfSides + 1, dPoints)
    
    Dim oRectangularTorus As IngrGeom3D.Revolution3d
    Set oRectangularTorus = oGeometryFactory.Revolutions3d.CreateByCurve(oOutputCol.ResourceManager, oLS3d, 0#, 0#, 1#, -dSwitchGearLength, -dSwitchGearWidth / 2, -dSwitchGearHeight / 2, dSweepAngle, bIsCapped)

    Set PlaceDoorOpen = oRectangularTorus
End Function



