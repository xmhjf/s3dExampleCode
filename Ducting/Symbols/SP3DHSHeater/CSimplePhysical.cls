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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         VRK
'   Creation Date:  Wednesday, May 23 2007
'   Description:
'       Slip-in Heaters
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   06.03.2008      VRK    CR-138091:Enhance slip-in heater symbol to support round duct and expose more dimensions
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
        
    '{<(InputParamDec)>}
    Dim parHeaterWidth As Double
    Dim parHeaterHeight As Double
    Dim parC As Double
    Dim parM As Double
    Dim parE As Double
    Dim parT As Double
    Dim parOR As Double
    Dim parL As Double
    Dim parHVACShape As Integer
    Dim parA As Double
    Dim parDuctDiameter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHeaterWidth = arrayOfInputs(2)     'W
    parHeaterHeight = arrayOfInputs(3)    'H
    parC = arrayOfInputs(4)
    parM = arrayOfInputs(5)
    parE = arrayOfInputs(6)
    parOR = arrayOfInputs(7)
    parT = arrayOfInputs(8)
    parL = arrayOfInputs(9)
    parHVACShape = arrayOfInputs(10)
    parA = arrayOfInputs(11)
    parDuctDiameter = arrayOfInputs(12)
    
    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    'Create Duct
    Dim oBox As Object
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Set oCenter = New AutoMath.DPosition
    
     If parHVACShape = Rectangular Then
        'Create Overhang Box
        oStPoint.Set parOR, 0.5 * parHeaterHeight + parT, 0
        oEnPoint.Set -(parM - parOR), -0.5 * parHeaterHeight - parL, parE
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
        
        oStPoint.Set 0, 0.5 * parHeaterHeight, 0
        oEnPoint.Set -parC, -0.5 * parHeaterHeight, -parHeaterWidth
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
        
    ElseIf parHVACShape = 4 Then 'Round
        'Create Adapter Plate  for round ducts
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Dim objCollection As Collection
        Set objCollection = New Collection
        Dim oLine As IngrGeom3D.Line3d
        Dim oComplexStr As IngrGeom3D.ComplexString3d
        Dim objSaddle As IngrGeom3D.Arc3d
        Dim dThickness As Double
        Dim iCount As Integer
        dThickness = 0.15 * parA
        Dim dY As Double
        dY = Sqr((parDuctDiameter / 2) ^ 2 - (parDuctDiameter / 2 - (parA - dThickness)) ^ 2)
        Dim dAngle As Double
        dAngle = Atn(dY / ((parDuctDiameter / 2) - (parA - dThickness)))

        'Create Duct
        oStPoint.Set 0, 0.5 * parHeaterHeight, dThickness
        oEnPoint.Set -parC, -0.5 * parHeaterHeight, -parHeaterWidth
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
        
        'Create Control Panel
        oStPoint.Set parOR, 0.5 * parHeaterHeight + parT, dThickness
        oEnPoint.Set -(parM - parOR), -0.5 * parHeaterHeight - parL, dThickness + parE
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing

        'Arc1
        oCenter.Set parOR, 0, -parDuctDiameter / 2
        oStPoint.Set parOR, dY, -(parA - dThickness)
        oEnPoint.Set parOR, -dY, -(parA - dThickness)

        Set objSaddle = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        oCenter.x, oCenter.y, oCenter.z, _
                                        oStPoint.x, oStPoint.y, oStPoint.z, _
                                        oEnPoint.x, oEnPoint.y, oEnPoint.z)
        objCollection.Add objSaddle
        Set objSaddle = Nothing
        
        'Line1
        oStPoint.Set parOR, -dY, -(parA - dThickness)
        oEnPoint.Set parOR, -(dY + dThickness * Sin(dAngle)), -((parA - dThickness) - dThickness * Cos(dAngle))
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        oStPoint.x, oStPoint.y, oStPoint.z, _
                                        oEnPoint.x, oEnPoint.y, oEnPoint.z)
        objCollection.Add oLine
        Set oLine = Nothing
        
        'Arc2
        oCenter.Set parOR, 0, -parDuctDiameter / 2
        oStPoint.Set parOR, -(dY + dThickness * Sin(dAngle)), -((parA - dThickness) - dThickness * Cos(dAngle))
        oEnPoint.Set oStPoint.x, -oStPoint.y, oStPoint.z
        Set objSaddle = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        oCenter.x, oCenter.y, oCenter.z, _
                                        oStPoint.x, oStPoint.y, oStPoint.z, _
                                        oEnPoint.x, oEnPoint.y, oEnPoint.z)
        objCollection.Add objSaddle
        Set objSaddle = Nothing
          
        'Line2
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set parOR, dY, -(parA - dThickness)
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        oStPoint.x, oStPoint.y, oStPoint.z, _
                                        oEnPoint.x, oEnPoint.y, oEnPoint.z)
        
        objCollection.Add oLine
        Set oLine = Nothing
        
        Dim objCurve As Object
        Dim oAxisVec As AutoMath.DVector
        Set oAxisVec = New DVector
        oAxisVec.Set -1, 0, 0
        oStPoint.Set parOR, dY, -(parA - dThickness)

        Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
        Set objCurve = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, parM, True)
        'Set the output
        m_OutputColl.AddOutput "ObjCurve", objCurve
        Set objCurve = Nothing
        Set oAxisVec = Nothing
        Set oComplexStr = Nothing
        Set objCollection = Nothing
    End If
        Set oStPoint = Nothing
        Set oEnPoint = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub
