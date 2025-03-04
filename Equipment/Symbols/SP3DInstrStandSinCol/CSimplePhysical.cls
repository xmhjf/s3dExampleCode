VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/13/2005-6:25:59 PM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         MS
'   Creation Date:  Monday 4th October, 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   13.Jan.2005     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "CSimplePhysical:" 'Used for error messages

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
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
        
    Dim iOutput     As Double
    Dim parInstrumentStandHeight  As Double
    Dim ObjBasePlate As Object

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentStandHeight = arrayOfInputs(2)
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
'Insert your code for output 1(Base Plate of Stand)
    'Assuming the base plate to be of 10in x 10in of 1/2in thickness.
    Dim dBasePlateWidth As Double
    Dim dBasePlateThk As Double
    dBasePlateWidth = 0.254
    dBasePlateThk = 0.0127
    
    oStPoint.Set -dBasePlateWidth / 2, -dBasePlateWidth / 2, 0
    oEndPoint.Set dBasePlateWidth / 2, dBasePlateWidth / 2, dBasePlateThk
    Set ObjBasePlate = PlaceBox(m_OutputColl, oStPoint, oEndPoint)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBasePlate
    Set ObjBasePlate = Nothing
    
'Insert your code for output 2(Column Body)
    'Assuming the column pipe to be of 3in SCH40 pipe. Hence outside diameter for it is 0.0889m
    Dim dColumnDiameter As Double
    dColumnDiameter = 0.0889
    oStPoint.Set 0, 0, parInstrumentStandHeight
    oEndPoint.Set 0, 0, dBasePlateThk
    ' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEndPoint, dColumnDiameter
        
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
' Insert your code for output 3( Fundation Port)
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory

      'Coordinate sysytem is at the bottom of BasePlate.
    '
    '               |-------|
    '   Y           |       |
    '   ^           |       |
    '   |           |       |
    '   |           |       |
    '   |           |       |
    '   -----> X    |       |
    '   Symbol CS   |       |
    '               |       |
    '               |   X   |
    '               |   ^   |
    '               |   |   |
    '               |   |   |
    '               |---|---|----> Y
    '                   Port CS

    'Assuming the base plate holes to be of 8in x 8in size
    Dim dBasePlateHoleCentoCen As Double
    dBasePlateHoleCentoCen = 0.2032
    
    Dim objEqpFoundationPort As IJEqpFoundationPort
    Set objEqpFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "EqpFoundationPort", _
                                                False, m_OutputColl.ResourceManager)
    Dim holes() As Variant
    Call objEqpFoundationPort.GetHoles(holes())
    holes(0, 1) = -dBasePlateHoleCentoCen / 2
    holes(0, 2) = -dBasePlateHoleCentoCen / 2
    holes(1, 1) = dBasePlateHoleCentoCen / 2
    holes(1, 2) = -dBasePlateHoleCentoCen / 2
    holes(2, 1) = dBasePlateHoleCentoCen / 2
    holes(2, 2) = dBasePlateHoleCentoCen / 2
    holes(3, 1) = -dBasePlateHoleCentoCen / 2
    holes(3, 2) = dBasePlateHoleCentoCen / 2


    Call objEqpFoundationPort.PutCS(0, 0, 0, 0, 1, 0, 0, 0, -1)
    Call objEqpFoundationPort.SetHoles(holes)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEqpFoundationPort
    Set objEqpFoundationPort = Nothing
        
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

