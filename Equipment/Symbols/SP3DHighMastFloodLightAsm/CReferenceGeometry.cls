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
    
    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Long
    
    Dim parPoleHeight As Double
    Dim parPoleDiameter1 As Double
    Dim parPoleDiameter2 As Double
    Dim parMountPlateLength As Double
    Dim parMountPlateWidth As Double
    Dim parMountPlateThickness As Double
    Dim parFixtureLength As Double
    Dim parFixtureWidth As Double
    Dim parFixtureHeight As Double
    Dim parBoxHeight As Double
    Dim parBoxWidth As Double
    Dim parBoxLength As Double
    Dim parwattage As Double
    Dim parCPx As Double
    Dim parCPy As Double
    Dim parCPz As Double
    
    'Inputs
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPoleHeight = arrayOfInputs(2)
    parPoleDiameter1 = arrayOfInputs(3)
    parPoleDiameter2 = arrayOfInputs(4)
    parMountPlateLength = arrayOfInputs(5)
    parMountPlateWidth = arrayOfInputs(6)
    parMountPlateThickness = arrayOfInputs(7)
    parFixtureHeight = arrayOfInputs(8)
    parFixtureWidth = arrayOfInputs(9)
    parFixtureLength = arrayOfInputs(10)
    parBoxHeight = arrayOfInputs(11)
    parBoxWidth = arrayOfInputs(12)
    parBoxLength = arrayOfInputs(13)
    parwattage = arrayOfInputs(14)
    
    If UBound(arrayOfInputs) > 14 Then
        parCPx = arrayOfInputs(15)
        parCPy = arrayOfInputs(16)
        parCPz = arrayOfInputs(17)
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
