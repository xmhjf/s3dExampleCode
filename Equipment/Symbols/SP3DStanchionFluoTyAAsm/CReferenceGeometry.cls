VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRefGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2009-10 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Monday, September 14 2009
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output:
'           Control Point is created at the origin
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   25.Mar.2010     RUK     CR-CP-177179  Enhance lighting symbols to provide control point for dimensioning purposes
'                           CR-CP-113328  "can be modified" flag should be True for all best practice equipment symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "RefGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt  As PartFacelets.IJDPart
    Dim iOutput As Integer
    
    Dim parPoleHeight As Double
    Dim parFixtureLength As Double
    Dim parPoleDiameter As Double
    Dim parAngle As Double
    Dim parwattage As Double
    Dim parFixtureWidth As Double
    Dim parFixtureDepth As Double
    Dim parCPx As Double
    Dim parCPy As Double
    Dim parCPz As Double
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPoleHeight = arrayOfInputs(2)    'A
    parPoleDiameter = arrayOfInputs(3)  'PD
    parAngle = arrayOfInputs(4)         'C
    parwattage = arrayOfInputs(5)
    parFixtureLength = arrayOfInputs(6)    'L
    parFixtureWidth = arrayOfInputs(7)  'W
    parFixtureDepth = arrayOfInputs(8)  'D
    
    If UBound(arrayOfInputs) > 8 Then
        parCPx = arrayOfInputs(9)
        parCPy = arrayOfInputs(10)
        parCPz = arrayOfInputs(11)
    End If
    
    iOutput = 0

    'Insert your code for Control Point (Output 11)
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    '==========================================
    'Construction of  Control Point
    '==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, _
                                                   parCPx, parCPy, parCPz, 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
