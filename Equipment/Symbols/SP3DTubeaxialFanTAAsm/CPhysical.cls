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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   File:           CPhysical.cls
'   Author:         CSA, Inc.
'   Creation Date:  Wednesday, April 18 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output: Body Cylinder.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper

Private Const E_FAIL = &H80004005

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oSymGeomHelper = New SymbolServices
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo Errx

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Integer
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl

    Dim oBody As Object
    Dim parFacetoFace As Double
    Dim parFlangeDiameter As Double
    Dim parFanInsideDiameter As Double
    Dim parBodyThickness As Double
    Dim parNozzleLength As Double
    Dim parFlangeThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parFlangeDiameter = arrayOfInputs(3)
    parFanInsideDiameter = arrayOfInputs(4)
    parBodyThickness = arrayOfInputs(5)
    parNozzleLength = arrayOfInputs(6)
    parFlangeThickness = arrayOfInputs(7)

    Dim oPos1 As IJDPosition
    Dim oPos2 As IJDPosition

    Set oPos1 = New DPosition
    Set oPos2 = New DPosition

    'Create the Body Cylinder (Output 1)
    oPos1.Set 0, 0, 0
    oPos2.Set -parFacetoFace, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oPos1, oPos2, parFanInsideDiameter + 2 * parBodyThickness
    Set oPos1 = Nothing
    Set oPos2 = Nothing

    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

