VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReferenceGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:32 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:          MS
'   Creation Date: Wednesday, Apr 7 2004
'   Description:
'    This is Exchanger Rear End type Q. Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E319 Symbol in Page no 310.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "ReferenceGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
    Dim iOutput     As Double
    
    Dim parFrontEndLength1 As Double
    Dim parExchangerLength As Double
    Dim parRearEndFlangeDia As Double
    Dim parRearEndFlangeTk1 As Double
    Dim parRearEndLength As Double
    Dim parRearEndDiameter As Double
    Dim parChannelDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFrontEndLength1 = arrayOfInputs(2)
    parExchangerLength = arrayOfInputs(3)
    parRearEndFlangeDia = arrayOfInputs(4) ' P40
    parRearEndFlangeTk1 = arrayOfInputs(5) ' P41
    parRearEndLength = arrayOfInputs(6) ' P42
    parRearEndDiameter = arrayOfInputs(7) ' P43
    parChannelDiameter = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    
    
    iOutput = 0
'   Origin is taken on axis of Exchanger at (parFrontEndLength1 + parExchangerLength) from the left face of the Rear end
    
' Insert your code for output 5(Control Point at origin )

'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, _
                                0, 0, 0, 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    
'   Create DefaultSurface - for output 6(a Circular shape is chosen)
    Dim ObjDefaultSurface As Object
    Dim objCircle As IngrGeom3D.Circle3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0, 0, 0, _
                            -1, 0, 0, _
                            0.1 * parChannelDiameter / 2)

'   create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set objCircle = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
