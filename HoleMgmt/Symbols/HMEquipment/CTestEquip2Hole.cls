VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTestEquip2Hole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:44 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   CTestEquip2Hole.cls
'
'   Description: The symbol consists of a cylinder and port
'
'   Change History:
'   dd.mmm.yyyy         who          change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CTestEquip2Hole:" 'Used for error messages

Private Sub Class_Initialize()
''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim CenterPos As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim ObjHoleBody As Object
    
    Dim oCylinderDia As Double
    Dim oCylinderLen As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    oCylinderDia = arrayOfInputs(2)
    oCylinderLen = arrayOfInputs(3)
   
    Dim InLength    As Double, InWidth    As Double, InHeight    As Double
    
    Dim oGeom       As New IngrGeom3D.GeometryFactory
    Dim oElements   As IJElements
    Set oElements = New JObjectCollection
    Dim oComplex    As IngrGeom3D.ComplexString3d
    Dim pIJProjection As IngrGeom3D.Projection3d
    Dim oArc        As IngrGeom3D.Arc3d
    Dim oLine       As IngrGeom3D.Line3d
    Dim oCircle     As IngrGeom3D.Circle3d
    Dim P1          As New AutoMath.DPosition
    Dim P2          As New AutoMath.DPosition
    Dim P3          As New AutoMath.DPosition
    Dim P4          As New AutoMath.DPosition
    Dim P5          As New AutoMath.DPosition
    Dim P6          As New AutoMath.DPosition
    Dim Cen1        As New AutoMath.DPosition
    Dim Cen2        As New AutoMath.DPosition

    InHeight = oCylinderLen
    InLength = oCylinderDia * 3
    InWidth = oCylinderDia * 1.5

    '=====================================================
    ' Hole Construct
    '=====================================================
    P1.Set InWidth / 2, InLength / 2 - InWidth / 2, InHeight / 2
    P2.Set InWidth / 2, -InLength / 2 + InWidth / 2, InHeight / 2
    P3.Set 0, -InLength / 2, InHeight / 2
    P4.Set -InWidth / 2, -InLength / 2 + InWidth / 2, InHeight / 2
    P5.Set -InWidth / 2, InLength / 2 - InWidth / 2, InHeight / 2
    P6.Set 0, InLength / 2, InHeight / 2
    
    Cen1.Set 0, InLength / 2 - InWidth / 2, InHeight / 2
    Cen2.Set 0, -InLength / 2 + InWidth / 2, InHeight / 2
    
    Set oLine = oGeom.Lines3d.CreateBy2Points(Nothing, P1.x, P1.y, P1.z, P2.x, P2.y, P2.z)
    oElements.Add oLine
    Set oArc = oGeom.Arcs3d.CreateByCenterStartEnd(Nothing, Cen2.x, Cen2.y, Cen2.z, P2.x, P2.y, P2.z, P3.x, P3.y, P3.z)
    oElements.Add oArc
    Set oArc = oGeom.Arcs3d.CreateByCenterStartEnd(Nothing, Cen2.x, Cen2.y, Cen2.z, P3.x, P3.y, P3.z, P4.x, P4.y, P4.z)
    oElements.Add oArc
    
    Set oLine = oGeom.Lines3d.CreateBy2Points(Nothing, P4.x, P4.y, P4.z, P5.x, P5.y, P5.z)
    oElements.Add oLine
    Set oArc = oGeom.Arcs3d.CreateByCenterStartEnd(Nothing, Cen1.x, Cen1.y, Cen1.z, P5.x, P5.y, P5.z, P6.x, P6.y, P6.z)
    oElements.Add oArc
    Set oArc = oGeom.Arcs3d.CreateByCenterStartEnd(Nothing, Cen1.x, Cen1.y, Cen1.z, P6.x, P6.y, P6.z, P1.x, P1.y, P1.z)
    oElements.Add oArc
    
    Set oComplex = oGeom.ComplexStrings3d.CreateByCurves(Nothing, oElements)
    Set pIJProjection = oGeom.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, oComplex, 0, 0, -1, InHeight + 0.1, True)
    
    m_OutputColl.AddOutput "HoleShape", pIJProjection
    
    Set pIJProjection = Nothing
    Set oPartFclt = Nothing
    Set CenterPos = Nothing
    
    Set oGeom = Nothing
    Set oElements = Nothing
    Set oComplex = Nothing
    Set pIJProjection = Nothing
    Set oArc = Nothing
    Set oLine = Nothing
    Set oCircle = Nothing
    Set P1 = Nothing
    Set P2 = Nothing
    Set P3 = Nothing
    Set P4 = Nothing
    Set P5 = Nothing
    Set P6 = Nothing
    Set Cen1 = Nothing
    Set Cen2 = Nothing
    
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
            Err.HelpFile, Err.HelpContext
End Sub


 