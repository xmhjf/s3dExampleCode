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
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   File:           CSimplePhysical.cls
'   Author:         VRK
'   Creation Date:  Tuesday, Oct 23 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   22.Oct.2007     VRK        CR-127559:Created Wye symbol with options Upright,1/8 bend and Tapped wyes
'   06.Nov.2007     VRK        CR-127562:Created Reducing Wye with 1/8 th bend
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private PI As Double
Private Const MODULE = "SimplePhysical:" 'Used for error messages

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
   
    PI = Atn(1) * 4
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptoffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam3       As Double
    Dim cptoffset2      As Double
    Dim depth2          As Double
    Dim flangeThick2    As Double
    Dim cptoffset3      As Double
    Dim depth3          As Double
    Dim flangeThick3    As Double
    Dim iOutput         As Integer

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFacetoFace As Double
    Dim parPorttoPortOffset As Double
    Dim parTangentLength1 As Double
    Dim parTangentLength3 As Double
    Dim parFacetoBendCenter As Double
    Dim parCentertoPortOffset As Double
    Dim parBendRadius As Double
    Dim parSeattoSeat As Double
    Dim parSeat1toCenter As Double
    Dim parSeat3toCenter As Double
    Dim parInsulationThickness As Double

    
' Inputs
     Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)
'    parFace2toCenter = arrayOfInputs(3)
'    parFace3toCenter = arrayOfInputs(4)
'    parFacetoFace = arrayOfInputs(5)
'    parPorttoPortOffset = arrayOfInputs(6)
'    parTangentLength1 = arrayOfInputs(7)
'    parTangentLength3 = arrayOfInputs(8)
'    parFacetoBendCenter = arrayOfInputs(9)
'    parCentertoPortOffset = arrayOfInputs(10)
'    parSeattoSeat = arrayOfInputs(11)
'    parSeat1toCenter = arrayOfInputs(12)
'    parSeat3toCenter = arrayOfInputs(13)
    parInsulationThickness = arrayOfInputs(14)
    
    iOutput = 0

  'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
   
    'Declaration of Variables for Creation of bends for PDBs 471 and 472
    Dim Surfset As IngrGeom3D.IJElements
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oNormal As AutoMath.DVector
    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Dim oTraceStr As IngrGeom3D.ComplexString3d
    Dim oCircle As IngrGeom3D.Circle3d
    Dim iCount As Integer
    Dim oCollection As Collection
    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim stnorm() As Double
    Dim ednorm() As Double
    Dim oCenter As AutoMath.DPosition
    Dim ObjSurface As Object
    Dim dArcY, dLength As Double
    Dim dArcRadius As Double
    Dim dFacetoface As Double
    Dim dYlength As Double
    
    'Declaration of Variables for Creation of nozzles
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim Port1Length As Double
    Dim Port2Length As Double
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                            cptoffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam, _
                                                                            cptoffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam, _
                                                                            cptoffset3, depth3
    Select Case lPartDataBasis
        'Upright Wye
        Case 471

            parFace1toCenter = arrayOfInputs(2)     'G
            parFace3toCenter = arrayOfInputs(4)     'H
            parFacetoFace = arrayOfInputs(5)        'F
            parPorttoPortOffset = arrayOfInputs(6) 'E
            
        ' Place Nozzle 1
            oPlacePoint.Set -(parFace1toCenter - (depth - cptoffset)), 0, 0
            Port1Length = parFace1toCenter
            oDir.Set -1, 0, 0
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
          
        ' Place Nozzle 2
            oPlacePoint.Set parFacetoFace - parFace1toCenter + cptoffset2 - depth2, 0, 0
            oDir.Set 1, 0, 0
            Port2Length = parFacetoFace - parFace1toCenter
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port2Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
         
        ' Place Nozzle 3
            oPlacePoint.Set parFace3toCenter + cptoffset3 - depth3, -parPorttoPortOffset, 0
            oDir.Set 1, 0, 0
            Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    
            'Create bend with 45 Degree
            Set oCenter = New DPosition
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oTraceStr = New ComplexString3d
            Set oCollection = New Collection
            Set oLine = New Line3d
            Set oArc = New Arc3d
            Set oNormal = New DVector
        
            oCenter.Set 0, 0, 0
            oNormal.Set Cos(PI / 4), -Sin(PI / 4), 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                     oCenter.x, oCenter.y, oCenter.z, _
                                                    oNormal.x, oNormal.y, oNormal.z, pipeDiam3 / 2)
            
            'Assuming ArcRadius be H/10
            dArcRadius = parFace3toCenter / 10
            dLength = dArcRadius * Tan(PI / 8)

            oStPoint.Set oCenter.x, oCenter.y, oCenter.z
            oEnPoint.Set parPorttoPortOffset - (dLength * Cos(PI / 4)), _
                    -parPorttoPortOffset + (dLength * Sin(PI / 4)), 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
                    
            oStPoint.Set parPorttoPortOffset - (dLength * Cos(PI / 4)), _
                    -parPorttoPortOffset + (dLength * Sin(PI / 4)), 0
            dArcY = -parPorttoPortOffset + dArcRadius
            oCenter.Set parPorttoPortOffset + dLength, dArcY, 0
            oEnPoint.Set parPorttoPortOffset + dLength, -parPorttoPortOffset, 0
            
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                     oCenter.x, oCenter.y, oCenter.z, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
                                    
            oCollection.Add oArc
            
            oStPoint.Set parPorttoPortOffset + dLength, -parPorttoPortOffset, 0
            oEnPoint.Set parFace3toCenter, -parPorttoPortOffset, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine

            oStPoint.Set 0, 0, 0
            Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                      m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                    CircularCorner, 0, stnorm, ednorm, False)

            For Each ObjSurface In Surfset
                m_OutputColl.AddOutput "Elbow", ObjSurface
            Next ObjSurface

           'Remove References
            For iCount = 1 To oCollection.Count
                oCollection.Remove 1
            Next iCount
            For iCount = 1 To Surfset.Count
                Surfset.Remove 1
            Next iCount
             
        'Upright Wye
        Case 506
    
            parFace1toCenter = arrayOfInputs(2)     'G
            parFacetoFace = arrayOfInputs(5)        'F
            parPorttoPortOffset = arrayOfInputs(6) 'E
            parFace3toCenter = parFacetoFace - parFace1toCenter
            
        ' Place Nozzle 1
            oPlacePoint.Set -(parFace1toCenter - (depth - cptoffset)), 0, 0
            Port1Length = parFace1toCenter
            oDir.Set -1, 0, 0
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
          
        ' Place Nozzle 2
            oPlacePoint.Set (parFacetoFace - parFace1toCenter + cptoffset2 - depth2), 0, 0
            oDir.Set 1, 0, 0
            Port2Length = parFacetoFace - parFace1toCenter
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port2Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
         
        ' Place Nozzle 3
            oPlacePoint.Set parFace3toCenter + cptoffset3 - depth3, -parPorttoPortOffset, 0
            oDir.Set 1, 0, 0
            Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    
            'Create bend with 45 Degree
            Set oCenter = New DPosition
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oTraceStr = New ComplexString3d
            Set oCollection = New Collection
            Set oLine = New Line3d
            Set oArc = New Arc3d
            Set oNormal = New DVector
        
            oCenter.Set 0, 0, 0
            oNormal.Set Cos(PI / 4), -Sin(PI / 4), 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                     oCenter.x, oCenter.y, oCenter.z, _
                                                    oNormal.x, oNormal.y, oNormal.z, pipeDiam3 / 2)
            
            'Assuming ArcRadius be H/10
            dArcRadius = parFace3toCenter / 10
            dLength = dArcRadius * Tan(PI / 8)

            oStPoint.Set oCenter.x, oCenter.y, oCenter.z
            oEnPoint.Set parPorttoPortOffset - (dLength * Cos(PI / 4)), _
                    -parPorttoPortOffset + (dLength * Sin(PI / 4)), 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
                    
            oStPoint.Set parPorttoPortOffset - (dLength * Cos(PI / 4)), _
                    -parPorttoPortOffset + (dLength * Sin(PI / 4)), 0
            
            dArcY = -parPorttoPortOffset + dArcRadius
            oCenter.Set parPorttoPortOffset + dLength, dArcY, 0
            oEnPoint.Set parPorttoPortOffset + dLength, -parPorttoPortOffset, 0
            
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                     oCenter.x, oCenter.y, oCenter.z, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
                                    
            oCollection.Add oArc
            
            oStPoint.Set parPorttoPortOffset + dLength, -parPorttoPortOffset, 0
            oEnPoint.Set parFace3toCenter, -parPorttoPortOffset, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine

            oStPoint.Set 0, 0, 0
            Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                      m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                    CircularCorner, 0, stnorm, ednorm, False)

            For Each ObjSurface In Surfset
                m_OutputColl.AddOutput "Elbow", ObjSurface
            Next ObjSurface

           'Remove References
            For iCount = 1 To oCollection.Count
                oCollection.Remove 1
            Next iCount
            For iCount = 1 To Surfset.Count
                Surfset.Remove 1
            Next iCount
            
        '1/8 BendWye
        '516 is for reducing wye
        Case 472, 516
            
            parFace3toCenter = arrayOfInputs(4)     'H
            parFacetoFace = arrayOfInputs(5)        'F
            parPorttoPortOffset = arrayOfInputs(6) 'D
            parTangentLength1 = arrayOfInputs(7)    'G

          ' Place Nozzle 1
            oPlacePoint.Set -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 + depth - cptoffset), 0, 0
            oDir.Set -1, 0, 0
            Port1Length = 1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1
            If CmpDblLessThan(Port1Length, flangeThick) Then Port1Length = flangeThick
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            
            'Place Nozzle 2
            oPlacePoint.Set (9 / 10 * (parFacetoFace - parTangentLength1) - (depth2 - cptoffset2)), 0, 0
            oDir.Set 1, 0, 0
            Port2Length = 9 / 10 * (parFacetoFace - parTangentLength1)
            If CmpDblLessThan(Port2Length, flangeThick2) Then Port2Length = flangeThick2
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port2Length)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            
            'Place Nozzle 3
            oPlacePoint.Set (parPorttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), parFace3toCenter - (depth3 - cptoffset3), 0
            oDir.Set 0, 1, 0
            Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            
            'Create Bend wye
            Set oCenter = New DPosition
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oTraceStr = New ComplexString3d
            Set oCollection = New Collection
            Set oLine = New Line3d
            Set oArc = New Arc3d
            Set oNormal = New DVector

            oCenter.Set (parPorttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), parFace3toCenter, 0
            oNormal.Set 0, 1, 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                  oCenter.x, oCenter.y, oCenter.z, _
                                                 oNormal.x, oNormal.y, oNormal.z, pipeDiam3 / 2)


            oStPoint.Set (parPorttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), parFace3toCenter, 0
            oEnPoint.Set (parPorttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), 0.75 * parFace3toCenter, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine

            dYlength = 0.25 * parFace3toCenter * ((1 / Cos(PI / 4)) - 1)
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                         (parPorttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1) - (0.25 * parFace3toCenter / Cos(PI / 4))), 0.75 * parFace3toCenter, 0, _
                         (parPorttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), 0.75 * parFace3toCenter, 0, _
                         (parPorttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1) - dYlength), 0.5 * parFace3toCenter, 0)
            oCollection.Add oArc

            oStPoint.Set (parPorttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1) - dYlength), 0.5 * parFace3toCenter, 0
            oEnPoint.Set 0, 0, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine

            oStPoint.Set (parPorttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), parFace3toCenter, 0
            Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)

             For Each ObjSurface In Surfset
                 m_OutputColl.AddOutput "Elbow", ObjSurface
             Next ObjSurface

            'Remove References
             For iCount = 1 To oCollection.Count
                 oCollection.Remove 1
             Next iCount
             For iCount = 1 To Surfset.Count
                 Surfset.Remove 1
             Next iCount
             
        '1/8 BendWye
        Case 513
            
            parFace1toCenter = arrayOfInputs(2)     'G
            parFace3toCenter = arrayOfInputs(4)     'E
            parFacetoFace = arrayOfInputs(5)        'F
            parPorttoPortOffset = arrayOfInputs(6) 'D

         ' Place Nozzle 1
            oPlacePoint.Set -(parFace1toCenter - (depth - cptoffset)), 0, 0
            Port1Length = parFace1toCenter
            oDir.Set -1, 0, 0
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
          
        ' Place Nozzle 2
            oPlacePoint.Set (parFacetoFace - parFace1toCenter + cptoffset2 - depth2), 0, 0
            oDir.Set 1, 0, 0
            Port2Length = parFacetoFace - parFace1toCenter
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port2Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle

            'Place Nozzle 3
            oPlacePoint.Set (parPorttoPortOffset - parFace1toCenter), parFace3toCenter - (depth3 - cptoffset3), 0
            oDir.Set 0, 1, 0
            Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            
            'Create Bend wye
            Set oCenter = New DPosition
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oTraceStr = New ComplexString3d
            Set oCollection = New Collection
            Set oLine = New Line3d
            Set oArc = New Arc3d
            Set oNormal = New DVector
            
            oCenter.Set (parPorttoPortOffset - parFace1toCenter), parFace3toCenter, 0
            oNormal.Set 0, 1, 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                  oCenter.x, oCenter.y, oCenter.z, _
                                                 oNormal.x, oNormal.y, oNormal.z, pipeDiam3 / 2)

            oStPoint.Set (parPorttoPortOffset - parFace1toCenter), parFace3toCenter, 0
            oEnPoint.Set (parPorttoPortOffset - parFace1toCenter), 0.75 * parFace3toCenter, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            
            dYlength = 0.25 * parFace3toCenter * ((1 / Cos(PI / 4)) - 1)
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                         (parPorttoPortOffset - parFace1toCenter - (0.25 * parFace3toCenter / Cos(PI / 4))), 0.75 * parFace3toCenter, 0, _
                         (parPorttoPortOffset - parFace1toCenter), 0.75 * parFace3toCenter, 0, _
                         (parPorttoPortOffset - parFace1toCenter - dYlength), 0.5 * parFace3toCenter, 0)
            oCollection.Add oArc

            oStPoint.Set (parPorttoPortOffset - parFace1toCenter - dYlength), 0.5 * parFace3toCenter, 0
            oEnPoint.Set 0, 0, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine

            oStPoint.Set (parPorttoPortOffset - parFace1toCenter), parFace3toCenter, 0
            Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                    m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)

             For Each ObjSurface In Surfset
                 m_OutputColl.AddOutput "Elbow", ObjSurface
             Next ObjSurface

            'Remove References
             For iCount = 1 To oCollection.Count
                 oCollection.Remove 1
             Next iCount
             For iCount = 1 To Surfset.Count
                 Surfset.Remove 1
             Next iCount
             
        '1/8 BendWye
        Case 514
            
            parFace1toCenter = arrayOfInputs(2)     'B
            parPorttoPortOffset = arrayOfInputs(6) 'G
            parSeattoSeat = arrayOfInputs(11)       'X
            parSeat3toCenter = arrayOfInputs(13)    'E
            
         ' Place Nozzle 1
            oPlacePoint.Set -(parFace1toCenter - (depth - cptoffset)), 0, 0
            Port1Length = parFace1toCenter
            oDir.Set -1, 0, 0
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
          
        ' Place Nozzle 2
            dFacetoface = parSeattoSeat + (depth - cptoffset) + (depth2 - cptoffset2)
            oPlacePoint.Set dFacetoface - parFace1toCenter + cptoffset2 - depth2, 0, 0
            oDir.Set 1, 0, 0
            Port2Length = dFacetoface - parFace1toCenter
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port2Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle

        'Place Nozzle 3
            oPlacePoint.Set parPorttoPortOffset - parFace1toCenter, parSeat3toCenter, 0
            oDir.Set 0, 1, 0
            Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            
            'Create Bend wye
            Set oCenter = New DPosition
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oTraceStr = New ComplexString3d
            Set oCollection = New Collection
            Set oLine = New Line3d
            Set oArc = New Arc3d
            Set oNormal = New DVector
            
            oCenter.Set (parPorttoPortOffset - parFace1toCenter), parSeat3toCenter, 0
            oNormal.Set 0, 1, 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                  oCenter.x, oCenter.y, oCenter.z, _
                                                 oNormal.x, oNormal.y, oNormal.z, pipeDiam3 / 2)

            oStPoint.Set (parPorttoPortOffset - parFace1toCenter), parSeat3toCenter, 0
            oEnPoint.Set (parPorttoPortOffset - parFace1toCenter), 0.75 * parSeat3toCenter, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            
            dYlength = 0.25 * parSeat3toCenter * ((1 / Cos(PI / 4)) - 1)
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                         (parPorttoPortOffset - parFace1toCenter - (0.25 * parSeat3toCenter / Cos(PI / 4))), 0.75 * parSeat3toCenter, 0, _
                         (parPorttoPortOffset - parFace1toCenter), 0.75 * parSeat3toCenter, 0, _
                         (parPorttoPortOffset - parFace1toCenter - dYlength), 0.5 * parSeat3toCenter, 0)
            oCollection.Add oArc

            oStPoint.Set (parPorttoPortOffset - parFace1toCenter - dYlength), 0.5 * parSeat3toCenter, 0
            oEnPoint.Set 0, 0, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine

            oStPoint.Set (parPorttoPortOffset - parFace1toCenter), parSeat3toCenter, 0
            Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)
             For Each ObjSurface In Surfset
                 m_OutputColl.AddOutput "Elbow", ObjSurface
             Next ObjSurface
            
            'Remove References
             For iCount = 1 To oCollection.Count
                 oCollection.Remove 1
             Next iCount
             For iCount = 1 To Surfset.Count
                 Surfset.Remove 1
             Next iCount
             
        '1/8 BendWye
        Case 515
            
            parTangentLength1 = arrayOfInputs(7)        'B
            parTangentLength3 = arrayOfInputs(8)        'A'
            parFacetoBendCenter = arrayOfInputs(9)      'C'
            parCentertoPortOffset = arrayOfInputs(10)   'H
            parSeattoSeat = arrayOfInputs(11)           'X
            parSeat1toCenter = arrayOfInputs(12)        'G
            parSeat3toCenter = arrayOfInputs(13)        'X'
            parBendRadius = oPipeComponent.BendRadius   'R'
                        
            parFace1toCenter = parSeat1toCenter - cptoffset + depth
            parFacetoFace = parSeattoSeat + (depth - cptoffset) + (depth2 - cptoffset2)
            parFace2toCenter = parFacetoFace - parFace1toCenter
            parFace3toCenter = parSeat3toCenter - cptoffset3 + depth3
            
         ' Place Nozzle 1
            oPlacePoint.Set -(parFace1toCenter - (depth - cptoffset)), 0, 0
            Port1Length = parFace1toCenter
            oDir.Set -1, 0, 0
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
          
        ' Place Nozzle 2
            oPlacePoint.Set (parFacetoFace - parFace1toCenter + cptoffset2 - depth2), 0, 0
            oDir.Set 1, 0, 0
            Port2Length = parFacetoFace - parFace1toCenter
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port2Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle

        'Place Nozzle 3
            oPlacePoint.Set parCentertoPortOffset, parSeat3toCenter, 0
            oDir.Set 0, 1, 0
            Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            
            'Create Bend wye
            Set oCenter = New DPosition
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oTraceStr = New ComplexString3d
            Set oCollection = New Collection
            Set oLine = New Line3d
            Set oArc = New Arc3d
            Set oNormal = New DVector
            
            oCenter.Set parCentertoPortOffset, parSeat3toCenter, 0
            oNormal.Set 0, 1, 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                  oCenter.x, oCenter.y, oCenter.z, _
                                                 oNormal.x, oNormal.y, oNormal.z, pipeDiam3 / 2)
                                                 
            oStPoint.Set parCentertoPortOffset, parSeat3toCenter, 0
            oEnPoint.Set parCentertoPortOffset, parSeat3toCenter + depth3 - cptoffset3 - parTangentLength3, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
                        
            oStPoint.Set parCentertoPortOffset, parSeat3toCenter + depth3 - cptoffset3 - parTangentLength3, 0
            oCenter.Set parCentertoPortOffset - parBendRadius, parSeat3toCenter + depth3 - cptoffset3 - parTangentLength3, 0
            oEnPoint.Set parCentertoPortOffset - parBendRadius * (1 - Sin(PI / 4)), parSeat3toCenter + depth3 - cptoffset3 - parTangentLength3 - parBendRadius * Cos(PI / 4), 0
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                    oCenter.x, oCenter.y, oCenter.z, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            oCollection.Add oArc
            
            oStPoint.Set parCentertoPortOffset - parBendRadius * (1 - Sin(PI / 4)), parSeat3toCenter + depth3 - cptoffset3 - parTangentLength3 - parBendRadius * Cos(PI / 4), 0
            oEnPoint.Set 0, 0, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            
            oStPoint.Set parCentertoPortOffset, parSeat3toCenter, 0
            Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)
                                 
             For Each ObjSurface In Surfset
                 m_OutputColl.AddOutput "Elbow", ObjSurface
             Next ObjSurface
            
            'Remove References
             For iCount = 1 To oCollection.Count
                 oCollection.Remove 1
             Next iCount
             For iCount = 1 To Surfset.Count
                 Surfset.Remove 1
             Next iCount
             
        '1/8 BendWye
        Case 517
            
            parFace1toCenter = arrayOfInputs(2)         'G
            parFace2toCenter = arrayOfInputs(3)         'E
            parFace3toCenter = arrayOfInputs(4)         'E'
            parFacetoFace = arrayOfInputs(5)            'F
            parTangentLength1 = arrayOfInputs(7)        'B
            parTangentLength3 = arrayOfInputs(8)        'A'
            parFacetoBendCenter = arrayOfInputs(9)      'C'
            parCentertoPortOffset = arrayOfInputs(10)   'H
            parBendRadius = oPipeComponent.BendRadius   'R'
            
            parFace2toCenter = parFacetoFace - parFace1toCenter
            
         ' Place Nozzle 1
            oPlacePoint.Set -(parFace1toCenter - (depth - cptoffset)), 0, 0
            Port1Length = parFace1toCenter
            oDir.Set -1, 0, 0
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
          
        ' Place Nozzle 2
            oPlacePoint.Set (parFacetoFace - parFace1toCenter + cptoffset2 - depth2), 0, 0
            oDir.Set 1, 0, 0
            Port2Length = parFacetoFace - parFace1toCenter
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port2Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle

        'Place Nozzle 3
            oPlacePoint.Set parCentertoPortOffset, parFace3toCenter - (depth3 - cptoffset3), 0
            oDir.Set 0, 1, 0
            Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            
            'Create Bend wye
            Set oCenter = New DPosition
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oTraceStr = New ComplexString3d
            Set oCollection = New Collection
            Set oLine = New Line3d
            Set oArc = New Arc3d
            Set oNormal = New DVector
            
            oCenter.Set parCentertoPortOffset, parFace3toCenter - (depth3 - cptoffset3), 0
            oNormal.Set 0, 1, 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                  oCenter.x, oCenter.y, oCenter.z, _
                                                 oNormal.x, oNormal.y, oNormal.z, pipeDiam3 / 2)
                                                 
            oStPoint.Set parCentertoPortOffset, parFace3toCenter - (depth3 - cptoffset3), 0
            oEnPoint.Set parCentertoPortOffset, parFace3toCenter - parTangentLength3, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
                        
            oStPoint.Set parCentertoPortOffset, parFace3toCenter - parTangentLength3, 0
            oCenter.Set parCentertoPortOffset - parBendRadius, parFace3toCenter - parTangentLength3, 0
            oEnPoint.Set parCentertoPortOffset - parBendRadius * (1 - Sin(PI / 4)), parFace3toCenter - parTangentLength3 - parBendRadius * Cos(PI / 4), 0
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                    oCenter.x, oCenter.y, oCenter.z, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            oCollection.Add oArc
            
            oStPoint.Set parCentertoPortOffset - parBendRadius * (1 - Sin(PI / 4)), parFace3toCenter - parTangentLength3 - parBendRadius * Cos(PI / 4), 0
            oEnPoint.Set 0, 0, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            
            oStPoint.Set parCentertoPortOffset, parFace3toCenter - depth3 + cptoffset3, 0
            Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)
            For Each ObjSurface In Surfset
                 m_OutputColl.AddOutput "Elbow", ObjSurface
             Next ObjSurface
                                 
            'Remove References
             For iCount = 1 To oCollection.Count
                 oCollection.Remove 1
             Next iCount
             For iCount = 1 To Surfset.Count
                 Surfset.Remove 1
             Next iCount
             
        'Tapped Wye
        Case 473
            
            parFace1toCenter = arrayOfInputs(2) 'G
            parFace3toCenter = arrayOfInputs(4) 'D
            parFacetoFace = arrayOfInputs(5)    'F
            
         ' Place Nozzle 1
            oPlacePoint.Set -(parFace1toCenter - (depth - cptoffset)), 0, 0
            Port1Length = parFace1toCenter
            oDir.Set -1, 0, 0
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
          
        ' Place Nozzle 2
            oPlacePoint.Set (parFacetoFace - parFace1toCenter + cptoffset2 - depth2), 0, 0
            oDir.Set 1, 0, 0
            Port2Length = parFacetoFace - parFace1toCenter
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port2Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            
        ' Place Nozzle 3
            oPlacePoint.Set (parFace3toCenter + cptoffset3 - depth3) * Sqr(1 / 2), _
                (parFace3toCenter + cptoffset3 - depth3) * Sqr(1 / 2), 0
            oDir.Set 1, 1, 0
            Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFace3toCenter)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
                        
        'Tapped Wye
        Case 507
            
            parFace1toCenter = arrayOfInputs(2) 'G
            parFace3toCenter = arrayOfInputs(4) 'X'
            parSeattoSeat = arrayOfInputs(11)   'X
        
         ' Place Nozzle 1
            oPlacePoint.Set -(parFace1toCenter - (depth - cptoffset)), 0, 0
            Port1Length = parFace1toCenter
            oDir.Set -1, 0, 0
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            
        ' Place Nozzle 2
            dFacetoface = parSeattoSeat + (depth - cptoffset) + (depth2 - cptoffset2)
            oPlacePoint.Set dFacetoface - parFace1toCenter + cptoffset2 - depth2, 0, 0
            oDir.Set 1, 0, 0
            Port2Length = dFacetoface - parFace1toCenter
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port2Length)
        ' Set the output
            iOutput = iOutput + 1
             m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            
        ' Place Nozzle 3
            oPlacePoint.Set (parFace3toCenter + cptoffset3 - depth3) * Sqr(1 / 2), _
                (parFace3toCenter + cptoffset3 - depth3) * Sqr(1 / 2), 0
            oDir.Set 1, 1, 0
            Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFace3toCenter)
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
  
        Case Else
            GoTo ErrorLabel:
    End Select
            Set oArc = Nothing
            Set objNozzle = Nothing
            Set oCenter = Nothing
            Set oCircle = Nothing
            Set oCollection = Nothing
            Set oDir = Nothing
            Set oEnPoint = Nothing
            Set oGeomFactory = Nothing
            Set oLine = Nothing
            Set oNormal = Nothing
            Set oPipeComponent = Nothing
            Set oPlacePoint = Nothing
            Set oStPoint = Nothing
            Set oTraceStr = Nothing
            Set Surfset = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
