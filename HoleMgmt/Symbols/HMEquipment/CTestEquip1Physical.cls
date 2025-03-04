VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTestEquip1Physical"
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
'   CTestEquip1Physical.cls
'
'   Description: The symbol consists of a cylinder and port
'
'   Change History:
'   dd.mmm.yyyy         who          change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CTestEquip1Physical:" 'Used for error messages

Private Sub Class_Initialize()
''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim CenterPos As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim ObjBody As Object
    Dim ObjInsulatedBody As Object
    Dim ObjPort As IJConduitPortOcc
    
    Dim oCylinderDia As Double
    Dim oCylinderLen As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    oCylinderDia = arrayOfInputs(2)
    oCylinderLen = arrayOfInputs(3)
    
    CenterPos.Set 0, 0, 0
    Dim StartPos As New AutoMath.DPosition
    Dim EndPos As New AutoMath.DPosition
    
    StartPos.Set CenterPos.x, CenterPos.y, CenterPos.z - oCylinderLen / 2
    EndPos.Set CenterPos.x, CenterPos.y, CenterPos.z + oCylinderLen / 2
    
    ' code for output 1 (ProjectedBody)
    Set ObjBody = PlaceCylinder(m_OutputColl, StartPos, EndPos, oCylinderDia, True)
    Set StartPos = Nothing
    Set EndPos = Nothing
    Set CenterPos = Nothing
    
    m_OutputColl.AddOutput "ProjectedBody", ObjBody
    Set ObjBody = Nothing
    
    ' code for output 2 (ConduitPort)
    Dim oDir As IJDVector
    Dim portPoint As AutoMath.DPosition
    Set portPoint = New AutoMath.DPosition
    Set oDir = New DVector
    oDir.Set 0, 0, 1
    
    portPoint.Set 0, 0, oCylinderLen / 2
    Set ObjPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 1)
    
    m_OutputColl.AddOutput "ConduitPort", ObjPort
    Set ObjPort = Nothing
    Set portPoint = Nothing
    Set oDir = Nothing
    
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
            Err.HelpFile, Err.HelpContext
End Sub
 