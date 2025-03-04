VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CInsulation.cls
'   Author:         RAH
'   Creation Date:  December 2005
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       This is a Generic Olet symbol. It takes inputs from the Generic Data tables.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private PI           As Double
Private RAD          As Double

Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices

    PI = 4 * Atn(1)
    RAD = 180 / PI
    
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
   
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjInsulatedBody As Object
    Dim parA As Double
    Dim parInsulationThickness As Double
    
    Dim CentertoInsu As Double
    Dim oInsuPoint         As New AutoMath.DPosition
    Dim oFacePoint         As New AutoMath.DPosition
    Dim parFlangelength As Double
    Dim parNippleLength As Double

   
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parA = arrayOfInputs(2)
    parFlangelength = arrayOfInputs(3)
    parNippleLength = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)

    m_oGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0

' Insert your code for output 2(Insulated Body)
    Dim dMajorBodyDia As Double
    
    'get the Header data
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth


    CentertoInsu = (0.5 * pipeDiam) * Cos(40# / RAD)
    oInsuPoint.Set 0, CentertoInsu, 0
    oFacePoint.Set 0, parA + (pipeDiam / 2) + parNippleLength + parFlangelength, 0

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
                        
    
    'Set the Output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oInsuPoint, oFacePoint, flangeDiam + parInsulationThickness * 2

    Set ObjInsulatedBody = Nothing
    Set oInsuPoint = Nothing
    Set oFacePoint = Nothing


    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

