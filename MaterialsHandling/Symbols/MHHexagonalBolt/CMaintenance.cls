VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2011, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author: BhupatiRaju Neelima
'   Creation Date: 28 Sep 2011
'   Description: The symbol consists of a cylinder for displaying the maintenance aspect.
'   This class module is the place for user to implement graphical part of VBSymbol for the maintenance aspect
'
'   Change History:
'   dd.mmm.yyyy         who          change description
'   -----------         -----        ------------------
'   28-Sep-2011         Neelima     Creating the file first time
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CMaintenance:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim CenterPos As New AutoMath.DPosition
    Dim StartPos As New AutoMath.DPosition
    Dim EndPos As New AutoMath.DPosition
    Dim ObjMaintenanceBody As Object
    Dim oCylinderDia As Double
    Dim oCylinderLen As Double
  
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    oCylinderLen = arrayOfInputs(3)  'Bolt Head thickness
    oCylinderDia = arrayOfInputs(9)  'Shank Diameter
    
    ' Initializing the variables
    CenterPos.Set 0, 0, 0
    oCylinderDia = (3# * oCylinderDia) + 0.01

    StartPos.Set CenterPos.x, CenterPos.y, CenterPos.z + oCylinderLen
    EndPos.Set CenterPos.x, CenterPos.y, CenterPos.z

    ' code for output 5 (HoleShape)
    'Creating a cylinderical volume for placing the maintenance aspect for bolts
    Set ObjMaintenanceBody = PlaceCylinder(m_OutputColl, StartPos, EndPos, oCylinderDia, True)
    m_OutputColl.AddOutput "WorkingSpace", ObjMaintenanceBody
    
    Set ObjMaintenanceBody = Nothing
    Set oPartFclt = Nothing
    Set CenterPos = Nothing
    Set StartPos = Nothing
    Set EndPos = Nothing
    
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub



