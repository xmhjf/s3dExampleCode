VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTestEquip1Hole"
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
'   CTestEquip1Hole.cls
'
'   Description: The symbol consists of a cylinder and port
'
'   Change History:
'   dd.mmm.yyyy         who          change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CTestEquip1Hole:" 'Used for error messages

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
   
    Dim StartPos As New AutoMath.DPosition
    Dim EndPos As New AutoMath.DPosition

    oCylinderDia = oCylinderDia * 1.5
    oCylinderLen = oCylinderLen * 1.5
    StartPos.Set CenterPos.x, CenterPos.y, CenterPos.z - oCylinderLen / 2
    EndPos.Set CenterPos.x, CenterPos.y, CenterPos.z + oCylinderLen / 2

    ' code for output 3 (HoleShape)
    Set ObjHoleBody = PlaceCylinder(m_OutputColl, StartPos, EndPos, oCylinderDia, True)
    m_OutputColl.AddOutput "HoleShape", ObjHoleBody
    
    Set ObjHoleBody = Nothing
    Set oPartFclt = Nothing
    Set CenterPos = Nothing
    Set StartPos = Nothing
    Set EndPos = Nothing
    
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
            Err.HelpFile, Err.HelpContext
End Sub


 