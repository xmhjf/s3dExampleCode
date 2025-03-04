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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   COperation.cls
'   Author: Veena
'   Creation Date:  Friday, Sep 15 2006
'
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************

Option Explicit
Private Const MODULE = "Operation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim ObjPanelboardBody As Object
    Dim ObjDoorSwing As Object
    
    Dim iOutput     As Double
    Dim parPanelboardHeight As Double
    Dim parPanelboardWidth As Double
    Dim parPanelboardLength As Double
    
    Dim pBodyCorner1   As IJDPosition
    Dim pBodyCorner2   As IJDPosition
    Set pBodyCorner1 = New DPosition
    Set pBodyCorner2 = New DPosition

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPanelboardHeight = arrayOfInputs(2)
    parPanelboardWidth = arrayOfInputs(3)
    parPanelboardLength = arrayOfInputs(4)
    
    iOutput = 0
    
    'Create the Panelboard Body
    pBodyCorner1.Set -parPanelboardLength, -parPanelboardWidth / 2, -parPanelboardHeight / 2
    pBodyCorner2.Set pBodyCorner1.x + parPanelboardLength, pBodyCorner1.y + parPanelboardWidth, pBodyCorner1.z + parPanelboardHeight
    Set ObjPanelboardBody = PlaceBox(m_outputColl, pBodyCorner1, pBodyCorner2)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPanelboardBody
    Set ObjPanelboardBody = Nothing

    'Create 270deg Panelboard Door swing
    Dim dSweepAngle As Double
    Dim PI As Double
    PI = 4 * Atn(1)
    dSweepAngle = (3 * PI) / 2
   
    Set ObjDoorSwing = PlaceDoorSwing(m_outputColl, dSweepAngle, parPanelboardLength, parPanelboardWidth, parPanelboardHeight, False)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDoorSwing
    Set ObjDoorSwing = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

'This function places a rectangular torus
Private Function PlaceDoorSwing(oOutputCol As Object, ByVal dSweepAngle As Double, ByVal dPanelboardLength As Double, ByVal dPanelboardWidth As Double, ByVal dPanelboardHeight As Double, ByVal bIsCapped As Boolean) As Object
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory

    Const lNumberOfSides As Long = 4 'four sides for rectangle
    
    Dim dPoints(1 To (lNumberOfSides + 1) * 3) As Double
    
    dPoints(1) = -dPanelboardLength
    dPoints(2) = dPanelboardWidth / 2
    dPoints(3) = dPanelboardHeight / 2
    
    dPoints(4) = -dPanelboardLength
    dPoints(5) = -(dPanelboardWidth / 2) + 0.001
    dPoints(6) = dPanelboardHeight / 2
    
    dPoints(7) = -dPanelboardLength
    dPoints(8) = -(dPanelboardWidth / 2) + 0.001
    dPoints(9) = -dPanelboardHeight / 2
    
    dPoints(10) = -dPanelboardLength
    dPoints(11) = dPanelboardWidth / 2
    dPoints(12) = -dPanelboardHeight / 2
    
    dPoints(13) = dPoints(1)
    dPoints(14) = dPoints(2)
    dPoints(15) = dPoints(3)

    Dim oLS3d As IJLineString
    Set oLS3d = oGeometryFactory.LineStrings3d.CreateByPoints(Nothing, lNumberOfSides + 1, dPoints)
    
    Dim oRectangularTorus As IngrGeom3D.Revolution3d
    Set oRectangularTorus = oGeometryFactory.Revolutions3d.CreateByCurve(oOutputCol.ResourceManager, oLS3d, 0#, 0#, 1#, -dPanelboardLength, -dPanelboardWidth / 2, -dPanelboardHeight / 2, dSweepAngle, bIsCapped)

    Set PlaceDoorSwing = oRectangularTorus
End Function
