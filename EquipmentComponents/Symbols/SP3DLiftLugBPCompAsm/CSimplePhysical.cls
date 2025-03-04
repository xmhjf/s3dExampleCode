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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:10:24 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, Feb 25 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim objBox1     As Object
    Dim objBox2     As Object
    Dim objBox3      As Object
    Dim objBox4      As Object
    Dim objBox5      As Object
    Dim objBox6      As Object
    Dim objBox7      As Object
    Dim objBox8      As Object
    
    Dim objPyra1     As Object
    Dim objPyra2     As Object
    Dim objPyra3     As Object
    Dim objPyra4     As Object
    Dim objPyra5     As Object
    Dim objPyra6     As Object
    Dim objPyra7     As Object
    Dim objPyra8     As Object
    
    Dim objPoint    As Object
    Dim objPlane    As Object
        
    Dim strPosition     As String
    Dim dblXLength      As Double
    Dim dblYLength      As Double
    Dim dblZLength      As Double
    Dim dblXBottom      As Double
    Dim dblYBottom      As Double
    Dim dblXTop         As Double
    Dim dblYTop         As Double
    Dim dblXOffset      As Double
    Dim dblYOffset      As Double
    Dim dblHeight       As Double
    
    Dim parVesselLugOffset   As Double
    Dim parVesselDiameter           As Double
    Dim parLugBasePlateWidth       As Double
    Dim parLugBasePlateLength      As Double
    Dim parLugBasePlateThickness   As Double
    Dim parLugGussetHeight         As Double
    Dim parLugGussetWidth          As Double
    Dim parLugGussetThickness      As Double
    Dim parLugGussetSpacing        As Double
    Dim parLugBoltSlotEccentricity As Double
    Dim parLugBoltDiameter         As Double
    
    Dim posPlane As IJDPosition
    Set posPlane = New DPosition
     
    Dim posOrigin As IJDPosition
    Set posOrigin = New DPosition
    
    Dim tmxMatrix As IJDT4x4
    Set tmxMatrix = New DT4x4
    tmxMatrix.LoadIdentity
    
    Dim myOri As Orientation
    Set myOri = New Orientation

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselLugOffset = arrayOfInputs(3)
    parLugBasePlateWidth = arrayOfInputs(4)
    parLugBasePlateLength = arrayOfInputs(5)
    parLugBasePlateThickness = arrayOfInputs(6)
    parLugGussetHeight = arrayOfInputs(7)
    parLugGussetWidth = arrayOfInputs(8)
    parLugGussetThickness = arrayOfInputs(9)
    parLugGussetSpacing = arrayOfInputs(10)
    parLugBoltSlotEccentricity = arrayOfInputs(11)
    parLugBoltDiameter = arrayOfInputs(12)
    
    iOutput = 0
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'****************************************************************************
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'****************************************************************************
 ' Insert your code for output Dynamic Plane_1(Support Lug Base 1 & Bolt Hole 1)
        
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim iCount          As Integer
    Dim Surfset         As IngrGeom3D.IJElements
    Dim WCaps           As Long
    Dim oElements       As IJElements
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Dim oSubShape       As Object
    Dim ObjTopSupport   As Object
    Dim arrPoints(11)    As IJDPosition
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oElements = New JObjectCollection
    
    For iCount = 0 To 5
        Set arrPoints(iCount) = New DPosition
    Next iCount
    posPlane.Set (parVesselDiameter / 2) + (parLugBasePlateWidth / 2), 0, parVesselLugOffset
    posOrigin.Set (parVesselDiameter / 2) + (parLugBoltSlotEccentricity), 0, parVesselLugOffset
    dblXLength = parLugBasePlateWidth
    dblYLength = parLugBasePlateLength
    dblZLength = parLugBasePlateThickness
    
    
' Set the radian input as false to take angle in degrees
    myOri.RadianInput = False
    myOri.ResetDefaultAxis
    myOri.RotationAboutZ = 45
    myOri.ApplyRotations
    loadOriIntoTransformationMatrix tmxMatrix, myOri
    
'Make boundry by points for projected hole
    arrPoints(0).Set posOrigin.x - (parLugBoltDiameter / 2 + 0.003), posOrigin.y - (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(1).Set posOrigin.x + (parLugBoltDiameter / 2 + 0.003), posOrigin.y - (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(2).Set arrPoints(1).x + parLugBoltDiameter / 2, posOrigin.y, posOrigin.z
    arrPoints(3).Set posOrigin.x + (parLugBoltDiameter / 2 + 0.003), posOrigin.y + (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(4).Set posOrigin.x - (parLugBoltDiameter / 2 + 0.003), posOrigin.y + (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(5).Set arrPoints(4).x - parLugBoltDiameter / 2, posOrigin.y, posOrigin.z

'Transform origin for projected hole
    For iCount = 0 To 5
        Set arrPoints(iCount) = getTransformedDPos(arrPoints(iCount), tmxMatrix)
    Next iCount
    
'Transform origin for plane
    Set posPlane = getTransformedDPos(posPlane, tmxMatrix)
'Build Plane
    Set objPlane = CreateDefaultPlane(Nothing, "E " & posPlane.x & " N " & posPlane.y & " U " & posPlane.z, dblXLength, dblYLength, "U", myOri)

'Build boundry from points and add to DynElements (2 lines, 2 Arcs)
    Set oSubShape = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0).x, arrPoints(0).y, arrPoints(0).z, _
        arrPoints(1).x, arrPoints(1).y, arrPoints(1).z)
    oElements.Add oSubShape
    
    Set oSubShape = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, arrPoints(1).x, arrPoints(1).y, arrPoints(1).z, _
        arrPoints(2).x, arrPoints(2).y, arrPoints(2).z, arrPoints(3).x, arrPoints(3).y, arrPoints(3).z)
    oElements.Add oSubShape

    Set oSubShape = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3).x, arrPoints(3).y, arrPoints(3).z, _
        arrPoints(4).x, arrPoints(4).y, arrPoints(4).z)
    oElements.Add oSubShape

    Set oSubShape = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, arrPoints(4).x, arrPoints(4).y, arrPoints(4).z, _
        arrPoints(5).x, arrPoints(5).y, arrPoints(5).z, arrPoints(0).x, arrPoints(0).y, arrPoints(0).z)
    oElements.Add oSubShape

'Create object from DynElements
    Set objCStr = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
    objPlane.AddHole objCStr
    oElements.Clear
    objCStr.RemoveCurve True

'Project the shape
    WCaps = 1
    Set Surfset = oGeomFactory.GeometryServices.CreateByProjection(m_outputColl.ResourceManager, _
       objPlane, 0, 0, 1, parLugBasePlateThickness, WCaps)

'Add the output
    For Each ObjTopSupport In Surfset
        iOutput = iOutput + 1
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjTopSupport
    Next ObjTopSupport

' Set the output
    Set ObjTopSupport = Nothing
    Set objPlane = Nothing
    Set Surfset = Nothing
    Set oGeomFactory = Nothing
    Set oElements = Nothing
    Set objCStr = Nothing
    Set oSubShape = Nothing


' Insert your code for output 1(Support Lug Gusset 1)
    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2), parLugGussetSpacing / 2, parVesselLugOffset + parLugBasePlateThickness / 2 + parLugBasePlateThickness
    dblXLength = parLugGussetWidth
    dblYLength = parLugGussetThickness
    dblZLength = parLugBasePlateThickness
    
    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objBox1 = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox1
    Set objBox1 = Nothing

' Insert your code for output 2(Support Lug Gusset 2)

    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2), -parLugGussetSpacing / 2, parVesselLugOffset + parLugBasePlateThickness / 2 + parLugBasePlateThickness
    dblXLength = parLugGussetWidth
    dblYLength = parLugGussetThickness
    dblZLength = parLugBasePlateThickness

    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objBox2 = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox2
    Set objBox2 = Nothing

' Insert your code for output 3(Support Lug Gusset 1)

    dblXBottom = parLugGussetWidth
    dblYBottom = parLugGussetThickness
    dblXTop = parLugGussetWidth * 0.22
    dblYTop = parLugGussetThickness
    dblXOffset = (dblXBottom - dblXTop) / -2
    dblYOffset = 0
    dblHeight = parLugGussetHeight - parLugBasePlateThickness

    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2) + dblXOffset / 2, parLugGussetSpacing / 2, parVesselLugOffset + parLugGussetHeight / 2 + parLugBasePlateThickness + parLugBasePlateThickness / 2
    
    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objPyra1 = CreatePyramid(m_outputColl, strPosition, myOri, dblXBottom, dblYBottom, dblXTop, dblYTop, dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPyra1
    Set objPyra1 = Nothing
    
' Insert your code for output 4(Support Lug Gusset 1)

    dblXBottom = parLugGussetWidth
    dblYBottom = parLugGussetThickness
    dblXTop = parLugGussetWidth * 0.22
    dblYTop = parLugGussetThickness
    dblXOffset = (dblXBottom - dblXTop) / -2
    dblYOffset = 0
    dblHeight = parLugGussetHeight - parLugBasePlateThickness

    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2) + dblXOffset / 2, -parLugGussetSpacing / 2, parVesselLugOffset + parLugGussetHeight / 2 + parLugBasePlateThickness + parLugBasePlateThickness / 2
    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objPyra2 = CreatePyramid(m_outputColl, strPosition, myOri, dblXBottom, dblYBottom, dblXTop, dblYTop, dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPyra2
    Set objPyra2 = Nothing
    
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'****************************************************************************
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'****************************************************************************
' Insert your code for output Dynamic Plane_1(Support Lug Base 2 & Bolt Hole 2)

    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oElements = New JObjectCollection
    
    For iCount = 0 To 5
        Set arrPoints(iCount) = New DPosition
    Next iCount
    posPlane.Set (parVesselDiameter / 2) + (parLugBasePlateWidth / 2), 0, parVesselLugOffset
    posOrigin.Set (parVesselDiameter / 2) + (parLugBoltSlotEccentricity), 0, parVesselLugOffset
    dblXLength = parLugBasePlateWidth
    dblYLength = parLugBasePlateLength
    dblZLength = parLugBasePlateThickness
    
    myOri.ResetDefaultAxis
    myOri.RotationAboutZ = 135
    myOri.ApplyRotations
    loadOriIntoTransformationMatrix tmxMatrix, myOri
    
'Make boundry by points for projected hole
    arrPoints(0).Set posOrigin.x - (parLugBoltDiameter / 2 + 0.003), posOrigin.y - (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(1).Set posOrigin.x + (parLugBoltDiameter / 2 + 0.003), posOrigin.y - (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(2).Set arrPoints(1).x + parLugBoltDiameter / 2, posOrigin.y, posOrigin.z
    arrPoints(3).Set posOrigin.x + (parLugBoltDiameter / 2 + 0.003), posOrigin.y + (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(4).Set posOrigin.x - (parLugBoltDiameter / 2 + 0.003), posOrigin.y + (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(5).Set arrPoints(4).x - parLugBoltDiameter / 2, posOrigin.y, posOrigin.z

'Transform origin for projected hole
    For iCount = 0 To 5
        Set arrPoints(iCount) = getTransformedDPos(arrPoints(iCount), tmxMatrix)
    Next iCount
    
'Transform origin for plane
    Set posPlane = getTransformedDPos(posPlane, tmxMatrix)
'Build Plane
    Set objPlane = CreateDefaultPlane(Nothing, "E " & posPlane.x & " N " & posPlane.y & " U " & posPlane.z, dblXLength, dblYLength, "U", myOri)

'Build boundry from points and add to DynElements (2 lines, 2 Arcs)
    Set oSubShape = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0).x, arrPoints(0).y, arrPoints(0).z, _
        arrPoints(1).x, arrPoints(1).y, arrPoints(1).z)
    oElements.Add oSubShape
    
    Set oSubShape = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, arrPoints(1).x, arrPoints(1).y, arrPoints(1).z, _
        arrPoints(2).x, arrPoints(2).y, arrPoints(2).z, arrPoints(3).x, arrPoints(3).y, arrPoints(3).z)
    oElements.Add oSubShape

    Set oSubShape = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3).x, arrPoints(3).y, arrPoints(3).z, _
        arrPoints(4).x, arrPoints(4).y, arrPoints(4).z)
    oElements.Add oSubShape

    Set oSubShape = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, arrPoints(4).x, arrPoints(4).y, arrPoints(4).z, _
        arrPoints(5).x, arrPoints(5).y, arrPoints(5).z, arrPoints(0).x, arrPoints(0).y, arrPoints(0).z)
    oElements.Add oSubShape

'Create object from DynElements
    Set objCStr = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
    objPlane.AddHole objCStr
    oElements.Clear
    objCStr.RemoveCurve True

'Project the shape
    WCaps = 1
    Set Surfset = oGeomFactory.GeometryServices.CreateByProjection(m_outputColl.ResourceManager, _
       objPlane, 0, 0, 1, parLugBasePlateThickness, WCaps)

'Add the output
    For Each ObjTopSupport In Surfset
        iOutput = iOutput + 1
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjTopSupport
    Next ObjTopSupport

' Set the output
    Set ObjTopSupport = Nothing
    Set objPlane = Nothing
    Set Surfset = Nothing
    Set oGeomFactory = Nothing
    Set oElements = Nothing
    Set objCStr = Nothing
    Set oSubShape = Nothing

' Insert your code for output 5(Support Lug Gusset 1)

    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2), parLugGussetSpacing / 2, parVesselLugOffset + parLugBasePlateThickness / 2 + parLugBasePlateThickness
    dblXLength = parLugGussetWidth
    dblYLength = parLugGussetThickness
    dblZLength = parLugBasePlateThickness
    
    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objBox3 = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox3
    Set objBox3 = Nothing

' Insert your code for output 6(Support Lug Gusset 2)

    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2), -parLugGussetSpacing / 2, parVesselLugOffset + parLugBasePlateThickness / 2 + parLugBasePlateThickness
    dblXLength = parLugGussetWidth
    dblYLength = parLugGussetThickness
    dblZLength = parLugBasePlateThickness

    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objBox4 = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox4
    Set objBox4 = Nothing

' Insert your code for output 7(Support Lug Gusset 1)
    dblXBottom = parLugGussetWidth
    dblYBottom = parLugGussetThickness
    dblXTop = parLugGussetWidth * 0.22
    dblYTop = parLugGussetThickness
    dblXOffset = (dblXBottom - dblXTop) / -2
    dblYOffset = 0
    dblHeight = parLugGussetHeight - parLugBasePlateThickness

    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2) + dblXOffset / 2, parLugGussetSpacing / 2, parVesselLugOffset + parLugGussetHeight / 2 + parLugBasePlateThickness + parLugBasePlateThickness / 2
    
    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objPyra3 = CreatePyramid(m_outputColl, strPosition, myOri, dblXBottom, dblYBottom, dblXTop, dblYTop, dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPyra3
    Set objPyra3 = Nothing
    
' Insert your code for output 8(Support Lug Gusset 1)

    dblXBottom = parLugGussetWidth
    dblYBottom = parLugGussetThickness
    dblXTop = parLugGussetWidth * 0.22
    dblYTop = parLugGussetThickness
    dblXOffset = (dblXBottom - dblXTop) / -2
    dblYOffset = 0
    dblHeight = parLugGussetHeight - parLugBasePlateThickness

    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2) + dblXOffset / 2, -parLugGussetSpacing / 2, parVesselLugOffset + parLugGussetHeight / 2 + parLugBasePlateThickness + parLugBasePlateThickness / 2
    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objPyra4 = CreatePyramid(m_outputColl, strPosition, myOri, dblXBottom, dblYBottom, dblXTop, dblYTop, dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPyra4
    Set objPyra4 = Nothing

'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'****************************************************************************
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'****************************************************************************
' Insert your code for output Dynamic Plane_1(Support Lug Base 3 & Bolt Hole 3)
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oElements = New JObjectCollection
    
    For iCount = 0 To 5
        Set arrPoints(iCount) = New DPosition
    Next iCount
    posPlane.Set (parVesselDiameter / 2) + (parLugBasePlateWidth / 2), 0, parVesselLugOffset
    posOrigin.Set (parVesselDiameter / 2) + (parLugBoltSlotEccentricity), 0, parVesselLugOffset
    dblXLength = parLugBasePlateWidth
    dblYLength = parLugBasePlateLength
    dblZLength = parLugBasePlateThickness
    
    myOri.ResetDefaultAxis
    myOri.RotationAboutZ = 225
    myOri.ApplyRotations
    loadOriIntoTransformationMatrix tmxMatrix, myOri
    
'Make boundry by points for projected hole
    arrPoints(0).Set posOrigin.x - (parLugBoltDiameter / 2 + 0.003), posOrigin.y - (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(1).Set posOrigin.x + (parLugBoltDiameter / 2 + 0.003), posOrigin.y - (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(2).Set arrPoints(1).x + parLugBoltDiameter / 2, posOrigin.y, posOrigin.z
    arrPoints(3).Set posOrigin.x + (parLugBoltDiameter / 2 + 0.003), posOrigin.y + (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(4).Set posOrigin.x - (parLugBoltDiameter / 2 + 0.003), posOrigin.y + (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(5).Set arrPoints(4).x - parLugBoltDiameter / 2, posOrigin.y, posOrigin.z

'Transform origin for projected hole
    For iCount = 0 To 5
        Set arrPoints(iCount) = getTransformedDPos(arrPoints(iCount), tmxMatrix)
    Next iCount
    
'Transform origin for plane
    Set posPlane = getTransformedDPos(posPlane, tmxMatrix)
'Build Plane
    Set objPlane = CreateDefaultPlane(Nothing, "E " & posPlane.x & " N " & posPlane.y & " U " & posPlane.z, dblXLength, dblYLength, "U", myOri)

'Build boundry from points and add to DynElements (2 lines, 2 Arcs)
    Set oSubShape = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0).x, arrPoints(0).y, arrPoints(0).z, _
        arrPoints(1).x, arrPoints(1).y, arrPoints(1).z)
    oElements.Add oSubShape
    
    Set oSubShape = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, arrPoints(1).x, arrPoints(1).y, arrPoints(1).z, _
        arrPoints(2).x, arrPoints(2).y, arrPoints(2).z, arrPoints(3).x, arrPoints(3).y, arrPoints(3).z)
    oElements.Add oSubShape

    Set oSubShape = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3).x, arrPoints(3).y, arrPoints(3).z, _
        arrPoints(4).x, arrPoints(4).y, arrPoints(4).z)
    oElements.Add oSubShape

    Set oSubShape = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, arrPoints(4).x, arrPoints(4).y, arrPoints(4).z, _
        arrPoints(5).x, arrPoints(5).y, arrPoints(5).z, arrPoints(0).x, arrPoints(0).y, arrPoints(0).z)
    oElements.Add oSubShape

'Create object from DynElements
    Set objCStr = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
    objPlane.AddHole objCStr
    oElements.Clear
    objCStr.RemoveCurve True

'Project the shape
    WCaps = 1
    Set Surfset = oGeomFactory.GeometryServices.CreateByProjection(m_outputColl.ResourceManager, _
       objPlane, 0, 0, 1, parLugBasePlateThickness, WCaps)

'Add the output
    For Each ObjTopSupport In Surfset
        iOutput = iOutput + 1
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjTopSupport
    Next ObjTopSupport

' Set the output
    Set ObjTopSupport = Nothing
    Set objPlane = Nothing
    Set Surfset = Nothing
    Set oGeomFactory = Nothing
    Set oElements = Nothing
    Set objCStr = Nothing
    Set oSubShape = Nothing
    
' Insert your code for output 9(Support Lug Gusset 1)
    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2), parLugGussetSpacing / 2, parVesselLugOffset + parLugBasePlateThickness / 2 + parLugBasePlateThickness
    dblXLength = parLugGussetWidth
    dblYLength = parLugGussetThickness
    dblZLength = parLugBasePlateThickness
    
    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objBox5 = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox5
    Set objBox5 = Nothing

' Insert your code for output 10(Support Lug Gusset 2)
    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2), -parLugGussetSpacing / 2, parVesselLugOffset + parLugBasePlateThickness / 2 + parLugBasePlateThickness
    dblXLength = parLugGussetWidth
    dblYLength = parLugGussetThickness
    dblZLength = parLugBasePlateThickness

    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objBox6 = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox6
    Set objBox6 = Nothing

' Insert your code for output 11(Support Lug Gusset 1)
    dblXBottom = parLugGussetWidth
    dblYBottom = parLugGussetThickness
    dblXTop = parLugGussetWidth * 0.22
    dblYTop = parLugGussetThickness
    dblXOffset = (dblXBottom - dblXTop) / -2
    dblYOffset = 0
    dblHeight = parLugGussetHeight - parLugBasePlateThickness

    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2) + dblXOffset / 2, parLugGussetSpacing / 2, parVesselLugOffset + parLugGussetHeight / 2 + parLugBasePlateThickness + parLugBasePlateThickness / 2
    
    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objPyra5 = CreatePyramid(m_outputColl, strPosition, myOri, dblXBottom, dblYBottom, dblXTop, dblYTop, dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPyra5
    Set objPyra5 = Nothing
    
' Insert your code for output 12(Support Lug Gusset 1)
    dblXBottom = parLugGussetWidth
    dblYBottom = parLugGussetThickness
    dblXTop = parLugGussetWidth * 0.22
    dblYTop = parLugGussetThickness
    dblXOffset = (dblXBottom - dblXTop) / -2
    dblYOffset = 0
    dblHeight = parLugGussetHeight - parLugBasePlateThickness

    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2) + dblXOffset / 2, -parLugGussetSpacing / 2, parVesselLugOffset + parLugGussetHeight / 2 + parLugBasePlateThickness + parLugBasePlateThickness / 2
    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objPyra6 = CreatePyramid(m_outputColl, strPosition, myOri, dblXBottom, dblYBottom, dblXTop, dblYTop, dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPyra6
    Set objPyra6 = Nothing

'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'****************************************************************************
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'****************************************************************************
' Insert your code for output Dynamic Plane_1(Support Lug Base 4 & Bolt Hole 4)
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oElements = New JObjectCollection
    
    For iCount = 0 To 5
        Set arrPoints(iCount) = New DPosition
    Next iCount
    posPlane.Set (parVesselDiameter / 2) + (parLugBasePlateWidth / 2), 0, parVesselLugOffset
    posOrigin.Set (parVesselDiameter / 2) + (parLugBoltSlotEccentricity), 0, parVesselLugOffset
    dblXLength = parLugBasePlateWidth
    dblYLength = parLugBasePlateLength
    dblZLength = parLugBasePlateThickness
    
    myOri.ResetDefaultAxis
    myOri.RotationAboutZ = 315
    myOri.ApplyRotations
    loadOriIntoTransformationMatrix tmxMatrix, myOri
    
'Make boundry by points for projected hole
    arrPoints(0).Set posOrigin.x - (parLugBoltDiameter / 2 + 0.003), posOrigin.y - (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(1).Set posOrigin.x + (parLugBoltDiameter / 2 + 0.003), posOrigin.y - (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(2).Set arrPoints(1).x + parLugBoltDiameter / 2, posOrigin.y, posOrigin.z
    arrPoints(3).Set posOrigin.x + (parLugBoltDiameter / 2 + 0.003), posOrigin.y + (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(4).Set posOrigin.x - (parLugBoltDiameter / 2 + 0.003), posOrigin.y + (parLugBoltDiameter / 2 + 0.003), posOrigin.z
    arrPoints(5).Set arrPoints(4).x - parLugBoltDiameter / 2, posOrigin.y, posOrigin.z

'Transform origin for projected hole
    For iCount = 0 To 5
        Set arrPoints(iCount) = getTransformedDPos(arrPoints(iCount), tmxMatrix)
    Next iCount
    
'Transform origin for plane
    Set posPlane = getTransformedDPos(posPlane, tmxMatrix)
'Build Plane
    Set objPlane = CreateDefaultPlane(Nothing, "E " & posPlane.x & " N " & posPlane.y & " U " & posPlane.z, dblXLength, dblYLength, "U", myOri)

'Build boundry from points and add to DynElements (2 lines, 2 Arcs)
    Set oSubShape = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0).x, arrPoints(0).y, arrPoints(0).z, _
        arrPoints(1).x, arrPoints(1).y, arrPoints(1).z)
    oElements.Add oSubShape
    
    Set oSubShape = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, arrPoints(1).x, arrPoints(1).y, arrPoints(1).z, _
        arrPoints(2).x, arrPoints(2).y, arrPoints(2).z, arrPoints(3).x, arrPoints(3).y, arrPoints(3).z)
    oElements.Add oSubShape

    Set oSubShape = oGeomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3).x, arrPoints(3).y, arrPoints(3).z, _
        arrPoints(4).x, arrPoints(4).y, arrPoints(4).z)
    oElements.Add oSubShape

    Set oSubShape = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, arrPoints(4).x, arrPoints(4).y, arrPoints(4).z, _
        arrPoints(5).x, arrPoints(5).y, arrPoints(5).z, arrPoints(0).x, arrPoints(0).y, arrPoints(0).z)
    oElements.Add oSubShape

'Create object from DynElements
    Set objCStr = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
    objPlane.AddHole objCStr
    oElements.Clear
    objCStr.RemoveCurve True

'Project the shape
    WCaps = 1
    Set Surfset = oGeomFactory.GeometryServices.CreateByProjection(m_outputColl.ResourceManager, _
       objPlane, 0, 0, 1, parLugBasePlateThickness, WCaps)

'Add the output
    For Each ObjTopSupport In Surfset
        iOutput = iOutput + 1
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjTopSupport
    Next ObjTopSupport

' Set the output
    Set ObjTopSupport = Nothing
    Set objPlane = Nothing
    Set Surfset = Nothing
    Set oGeomFactory = Nothing
    Set oElements = Nothing
    Set objCStr = Nothing
    Set oSubShape = Nothing
    Set posPlane = Nothing
    
' Insert your code for output 13(Support Lug Gusset 1)
    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2), parLugGussetSpacing / 2, parVesselLugOffset + parLugBasePlateThickness / 2 + parLugBasePlateThickness
    dblXLength = parLugGussetWidth
    dblYLength = parLugGussetThickness
    dblZLength = parLugBasePlateThickness
    
    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objBox7 = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox7
    Set objBox7 = Nothing

' Insert your code for output 14(Support Lug Gusset 2)
    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2), -parLugGussetSpacing / 2, parVesselLugOffset + parLugBasePlateThickness / 2 + parLugBasePlateThickness
    dblXLength = parLugGussetWidth
    dblYLength = parLugGussetThickness
    dblZLength = parLugBasePlateThickness

    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objBox8 = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox8
    Set objBox8 = Nothing

' Insert your code for output 15(Support Lug Gusset 1)
    dblXBottom = parLugGussetWidth
    dblYBottom = parLugGussetThickness
    dblXTop = parLugGussetWidth * 0.22
    dblYTop = parLugGussetThickness
    dblXOffset = (dblXBottom - dblXTop) / -2
    dblYOffset = 0
    dblHeight = parLugGussetHeight - parLugBasePlateThickness

    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2) + dblXOffset / 2, parLugGussetSpacing / 2, parVesselLugOffset + parLugGussetHeight / 2 + parLugBasePlateThickness + parLugBasePlateThickness / 2
    
    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objPyra7 = CreatePyramid(m_outputColl, strPosition, myOri, dblXBottom, dblYBottom, dblXTop, dblYTop, dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPyra7
    Set objPyra7 = Nothing
    
' Insert your code for output 16(Support Lug Gusset 1)
    dblXBottom = parLugGussetWidth
    dblYBottom = parLugGussetThickness
    dblXTop = parLugGussetWidth * 0.22
    dblYTop = parLugGussetThickness
    dblXOffset = (dblXBottom - dblXTop) / -2
    dblYOffset = 0
    dblHeight = parLugGussetHeight - parLugBasePlateThickness

    posOrigin.Set (parVesselDiameter / 2) + (parLugGussetWidth / 2) + dblXOffset / 2, -parLugGussetSpacing / 2, parVesselLugOffset + parLugGussetHeight / 2 + parLugBasePlateThickness + parLugBasePlateThickness / 2
    Set posOrigin = getTransformedDPos(posOrigin, tmxMatrix)
    strPosition = "E " & posOrigin.x & " N " & posOrigin.y & " U " & posOrigin.z
    Set objPyra8 = CreatePyramid(m_outputColl, strPosition, myOri, dblXBottom, dblYBottom, dblXTop, dblYTop, dblXOffset, dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPyra8
    Set objPyra8 = Nothing
    Set myOri = Nothing
    Set posOrigin = Nothing
    Set tmxMatrix = Nothing

'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'****************************************************************************

' Insert your code for output 17(Point)
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    Set CenterPos = Nothing

'Insert your code for output Dynamic Plane_1(Default Plane)
    Dim oPlane As IngrGeom3D.Plane3d
    Dim oCircle As IngrGeom3D.Circle3d
    Dim circleCenter As IJDPosition
    Dim circleNormal As DVector
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set circleCenter = New DPosition
    Set circleNormal = New DVector
    
    circleNormal.Set 0, 0, -1
    circleCenter.Set 0, 0, 0

    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            circleCenter.x, circleCenter.y, circleCenter.z, _
                            circleNormal.x, circleNormal.y, circleNormal.z, _
                            0.1 * parVesselDiameter / 2)
    Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(m_outputColl.ResourceManager, oCircle)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPlane

    Set oPlane = Nothing
    Set oCircle = Nothing
    Set circleNormal = Nothing
    Set circleCenter = Nothing
    Set oGeomFactory = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
