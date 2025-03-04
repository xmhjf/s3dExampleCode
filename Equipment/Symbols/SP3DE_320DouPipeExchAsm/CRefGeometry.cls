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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:03:06 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CRefGeometry.cls
'   Author:         RUK
'   Creation Date:  Tuesday, Feb 27 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output:
'       ControlPoint at the origin of equipment
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     ---                 ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "RefGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput         As Double
    
    Dim m_oGBSFactory   As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
    Dim parBlockEndFromPP As Double         'P1
    Dim parPipeEndFromPP As Double          'P2
    Dim parPipeCentoCen As Double           'P3
    Dim parPipeDiameter As Double           'P4
    Dim parBlockLength As Double '          'P5
    Dim parBlockHeight As Double            'P6
    Dim parBlockWidth As Double             'P7
    Dim parBundlePullingLength As Double    'P8
    Dim parSupport1CenFromPP As Double      'P9
    Dim parSupportCentoCen As Double        'P10
    Dim parSupportHeight As Double          'P11
    Dim parSupportThickness As Double       'P12
    Dim parSupportWidth As Double           'P13

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBlockEndFromPP = arrayOfInputs(2)
    parPipeEndFromPP = arrayOfInputs(3)
    parPipeCentoCen = arrayOfInputs(4)
    parPipeDiameter = arrayOfInputs(5)
    parBlockLength = arrayOfInputs(6)
    parBlockHeight = arrayOfInputs(7)
    parBlockWidth = arrayOfInputs(8)
    parBundlePullingLength = arrayOfInputs(9)
    parSupport1CenFromPP = arrayOfInputs(10)
    parSupportCentoCen = arrayOfInputs(11)
    parSupportHeight = arrayOfInputs(12)
    parSupportThickness = arrayOfInputs(13)
    parSupportWidth = arrayOfInputs(14)
  
    '==========================================
    'Construction of  Control Point
    '==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
       
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
