VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEquipmentHole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2011, Intergraph Corporation. All rights reserved.
'
'   CEquipmentHole.cls
'   Author: BhupatiRaju Neelima
'   Creation Date: 27 Sep 2011
'   Description: The symbol consists of a cylinder and port for the equipment holes
'   This class module is the place for user to implement graphical part of VBSymbol for the Equipmenthole aspect
'
'   Change History:
'   dd.mmm.yyyy         who          change description
'   -----------         -----        ------------------
'   27-Sep-2011         Neelima     Creating the file first time
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CEquipmentHole:" 'Used for error messages

Private Sub Class_Initialize()
''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim CenterPos As New AutoMath.DPosition
    Dim StartPos As New AutoMath.DPosition
    Dim EndPos As New AutoMath.DPosition
    Dim ObjHoleBody As Object
    Dim oCylinderDia As Double
    Dim oCylinderLen As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    oCylinderLen = arrayOfInputs(8) 'Shank Length
    oCylinderDia = arrayOfInputs(9) 'Shank Diameter

    ' Initializing the requied variables for creating the hole graphics
    CenterPos.Set 0, 0, 0
    oCylinderDia = oCylinderDia + 0.002
    
    StartPos.Set CenterPos.x, CenterPos.y, CenterPos.z - oCylinderLen
    EndPos.Set CenterPos.x, CenterPos.y, CenterPos.z

    ' Code for output 6 (HoleShape)
    ' Creating the cylinder graphics for cutting holes around the bolt tobe placed on the flanges
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

