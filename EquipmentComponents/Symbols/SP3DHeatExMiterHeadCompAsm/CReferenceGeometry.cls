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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:11:13 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:          ACM
'   Creation Date:  Friday, May 16, 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   23.Aug.2003  SymbolTeam(India)       Updated Error Handling Code
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'   This is an Exchanger Miter Head Symbol.
'   This symbol is a Fully parametic symbol, consists of 23 inputs and 9 outputs.It has 2 aspects one is Simple Physical
'   another is ReferenceGeometry. 7 are physical and 2 is ReferenceGeometry.
'   This is a Dow Equipment E-161 exchanger Miter Head symbol

Option Explicit

Private Const MODULE = "ReferenceGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
   
    Dim parVesselDiameter As Double
    Dim parFaceDiameter As Double
    Dim parFaceThickness As Double
    Dim parCPtoFace As Double
    
        
    Dim iOutput     As Double
   
    Dim ObjMiterTop15degBody As Object
    Dim ObjMiterTop30degBody As Object
    Dim ObjMiterBottom30degBody As Object
    Dim ObjMiterBottom15degBody As Object

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parFaceDiameter = arrayOfInputs(3)
    parFaceThickness = arrayOfInputs(4)
    parCPtoFace = arrayOfInputs(5)
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New GeometryFactory

    iOutput = 0

' Insert your code for output 8(Reference plane on at Control Point of vessel)
    ' create non-persistent circle to use for creating default surface ---
    
    Dim ObjDefaultSurface As Object
    Dim objCircle As IngrGeom3D.Circle3d
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0, 0, 0, _
                            0, 0, 1, _
                            0.1 * parVesselDiameter / 2)

'   create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set objCircle = Nothing
    
' Insert your code for output 9( Point on Vessel axis )
    Dim ObjPoint As IngrGeom3D.Point3d

    Set ObjPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, 0, 0, 0)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPoint
    Set ObjPoint = Nothing

    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub



