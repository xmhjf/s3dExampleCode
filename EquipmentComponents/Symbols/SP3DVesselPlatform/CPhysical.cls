VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:19:19 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: RAH
'   Creation Date:  July 2004
'   Description:
'   Construct a vessel platform with single section or double section. The sections can be a circular or triangular types.
'   Also, this symbol includes the facility for creating up to 3 holes for each section.
'   This symbols will built a vessel platform with the following conditions:
'   Input values limit values:
'   - Angle1, Angle2 and Angle3 should be between 0 and 360
'   - Projection distance fot the location of the access ladder  must be between 0 to 750 mm
'   - Angle1 should be less than (Angle2 - Start Sweep Angle)
'   - Angle2 should be greater than the Start Sweep Angle
'   - Angle2 should be less than 180 and Angle1 should be less than 180 for rectangular section
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const E_FAIL = -2147467259
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    
    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double

    Dim parVesselDiameter As Double
    
    Dim parPlatformWidth1 As Double
    Dim parPlatformHeight1 As Double
    
    Dim parPlatformWidth2 As Double
    Dim parPlatformHeight2 As Double
    
    Dim parAngle1 As Double
    Dim parAngle2 As Double
    Dim parAngle3 As Double
    
    Dim parOrientation As Long
    Dim parNumPlatform As Long
    Dim parPlatformType As Long
    Dim parLadderDistance As Double
    
    Dim NumberofHoles1 As Long
    Dim NumberofHoles2 As Long
    Dim HoleRadius(1 To 6) As Double
    
    Dim HoleAngle1 As Double
    Dim HolePosition1 As Double
    
    Dim HoleAngle2 As Double
    Dim HolePosition2 As Double
    
    Dim HoleAngle3 As Double
    Dim HolePosition3 As Double
     
    Dim HoleAngle4 As Double
    Dim HolePosition4 As Double
     
    Dim HoleAngle5 As Double
    Dim HolePosition5 As Double
    
    Dim HoleAngle6 As Double
    Dim HolePosition6 As Double
      
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parPlatformWidth1 = arrayOfInputs(3)
    parPlatformHeight1 = arrayOfInputs(4)
    parAngle1 = arrayOfInputs(5)
    parAngle2 = arrayOfInputs(6)
    
    parLadderDistance = arrayOfInputs(7)
    parOrientation = arrayOfInputs(8)
    parPlatformType = arrayOfInputs(9)
    parNumPlatform = arrayOfInputs(10)
       
    parAngle3 = arrayOfInputs(11)
    parPlatformWidth2 = arrayOfInputs(12)
    parPlatformHeight2 = arrayOfInputs(13)
                
    NumberofHoles1 = arrayOfInputs(14)
    
    HoleAngle1 = arrayOfInputs(15)
    HolePosition1 = arrayOfInputs(16)
    HoleRadius(1) = arrayOfInputs(17)
    
    HoleAngle2 = arrayOfInputs(18)
    HolePosition2 = arrayOfInputs(19)
    HoleRadius(2) = arrayOfInputs(20)
    
    HoleAngle3 = arrayOfInputs(21)
    HolePosition3 = arrayOfInputs(22)
    HoleRadius(3) = arrayOfInputs(23)
    
    NumberofHoles2 = arrayOfInputs(24)
    
    HoleAngle4 = arrayOfInputs(25)
    HolePosition4 = arrayOfInputs(26)
    HoleRadius(4) = arrayOfInputs(27)
    
    HoleAngle5 = arrayOfInputs(28)
    HolePosition5 = arrayOfInputs(29)
    HoleRadius(5) = arrayOfInputs(30)
    
    HoleAngle6 = arrayOfInputs(31)
    HolePosition6 = arrayOfInputs(32)
    HoleRadius(6) = arrayOfInputs(33)
    
    Dim dOrigintoPlatform As Double
    dOrigintoPlatform = parVesselDiameter / 2
    
    Dim p2angle_in As Double
    Dim p2angle_out As Double
    
    p2angle_in = Arcsin(Abs(parLadderDistance) / Abs(dOrigintoPlatform))
    p2angle_out = Arcsin(Abs(parLadderDistance) / Abs(dOrigintoPlatform + parPlatformWidth1))
    
    '-----------------------------------------------------------------------------------------------
    ' Check the inputs
    '-----------------------------------------------------------------------------------------------

    If (parLadderDistance > 0.75) Then
        oErrors.Add E_FAIL, "SP3DPlatform.CPlatform", "Ladder Distance must be less than 750 mm", ""
        GoTo ErrorLabel
    End If
    
    If (parAngle1 = 2 * PI) Then
        parAngle1 = 6.2831
    End If
    
    If (parAngle2 = 2 * PI) Then
        parAngle2 = 6.2831
    End If
 
    If (parAngle3 = 2 * PI) Then
        parAngle3 = 6.2831
    End If
 
    If (parAngle1 > 2 * PI Or parAngle1 < 0) Then
        oErrors.Add E_FAIL, "SP3DPlatform.CPlatform", "Angle1 should be between 0 and 360", ""
        GoTo ErrorLabel
    End If
    
    If (parAngle2 > 2 * PI Or parAngle2 < 0) Then
        oErrors.Add E_FAIL, "SP3DPlatform.CPlatform", "Angle2 should be between 0 and 360", ""
        GoTo ErrorLabel
    End If
            
    If parAngle1 > (parAngle2 - p2angle_in) Then
        oErrors.Add E_FAIL, "SP3DPlatform.CPlatform", "Angle1 should be less than Angle2 - Start Sweep Angle", ""
        GoTo ErrorLabel
    End If

    If (parAngle2 < p2angle_in) Then
        oErrors.Add E_FAIL, "SP3DPlatform.CPlatform", "Angle2 should be greater than the Start Sweep Angle", ""
        GoTo ErrorLabel
    End If

    If parPlatformType = 2 Then
        If parAngle1 >= PI Or parAngle2 >= PI Then
            oErrors.Add E_FAIL, "SP3DPlatform.CPlatform", "Angle2 < 180 and Angle1 < 180 for rectangular platform", ""
            GoTo ErrorLabel
        End If
    End If

    If (parAngle3 < parAngle2) And (parNumPlatform = 2) Then
        oErrors.Add E_FAIL, "SP3DPlatform.CPlatform", "Angle3 should be greater than Angle2", ""
        GoTo ErrorLabel
    End If

    '-----Section 1 ------------------------------------------------------------------------------------------
    
    Dim x11 As Double
    Dim x12 As Double
    Dim H(1 To 6) As Double
        
    Dim end_angle As Double
        
    If parOrientation = 1 Then 'counterclockwise
        x11 = parAngle1 + p2angle_in
        x12 = parAngle1 + p2angle_out
        H(1) = HoleAngle1
        H(2) = HoleAngle2
        H(3) = HoleAngle3
    Else
        x11 = -parAngle1 - p2angle_in
        x12 = -parAngle1 - p2angle_out
        H(1) = -HoleAngle1
        H(2) = -HoleAngle2
        H(3) = -HoleAngle3
    End If
        
    If parOrientation = 1 Then
        end_angle = parAngle2
    Else
        end_angle = -parAngle2
    End If
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
          
'Creation of Platform
    'Create the lines and arcs of the top surface . use this lines and arcs to create the complexstring.
    'Use the complex string to project the thickness.
    
    Dim oTLine1 As Object
    Dim oTArc2 As Object
    Dim oTLine2 As Object
    Dim oTArc1 As Object
    
    Dim ObjPlatformBody As Object
       
    Dim oCenterPoint As New AutoMath.DPosition
    Dim arrPt(1 To 7) As New AutoMath.DPosition
    
    Dim oNormalVector   As IJDVector
    Set oNormalVector = New AutoMath.DVector
    
    Dim oCollection  As Collection
    Set oCollection = New Collection
    
    If parOrientation = 1 Then
        oNormalVector.Set 0, 0, 1
    Else
        oNormalVector.Set 0, 0, -1
    End If
    
    oCenterPoint.Set 0, 0, 0
    
    Dim t1 As Double
    t1 = 90# * PI / 180#
    
    Dim x_rad As Double
    
    If parPlatformType = 1 Then
        x_rad = Abs(dOrigintoPlatform + parPlatformWidth1)
    Else
        x_rad = (dOrigintoPlatform + parPlatformWidth1) / Cos((end_angle - x11) / 2)
        x_rad = Abs(x_rad)
    End If
    
    arrPt(1).Set (dOrigintoPlatform) * Cos(end_angle - t1), -(dOrigintoPlatform) * Sin(end_angle - t1), 0
    arrPt(2).Set x_rad * Cos(end_angle - t1), -x_rad * Sin(end_angle - t1), 0
    arrPt(3).Set x_rad * Cos(x12 - t1), -x_rad * Sin(x12 - t1), 0
    arrPt(4).Set (dOrigintoPlatform) * Cos(x11 - t1), -(dOrigintoPlatform) * Sin(x11 - t1), 0
    
'Holes for first sections
    arrPt(5).Set (HolePosition1) * Cos(H(1) - t1), -(HolePosition1) * Sin(H(1) - t1), 0
    arrPt(6).Set (HolePosition2) * Cos(H(2) - t1), -(HolePosition2) * Sin(H(2) - t1), 0
    arrPt(7).Set (HolePosition3) * Cos(H(3) - t1), -(HolePosition3) * Sin(H(3) - t1), 0
    
    'Top Line 1
    iOutput = 0
    Set oTLine1 = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, arrPt(1).X, arrPt(1).y, arrPt(1).z, arrPt(2).X, arrPt(2).y, arrPt(2).z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTLine1
    oCollection.Add oTLine1
    Set oTLine1 = Nothing
      
    If parPlatformType = 1 Then
'    Top Arc 2
        Set oTArc2 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(m_OutputColl.ResourceManager, _
                                                          oCenterPoint.X, oCenterPoint.y, oCenterPoint.z, _
                                                         oNormalVector.X, oNormalVector.y, oNormalVector.z, _
                                                         arrPt(2).X, arrPt(2).y, arrPt(2).z, _
                                                        arrPt(3).X, arrPt(3).y, arrPt(3).z)
        m_OutputColl.AddOutput "line_", oTArc2
        oCollection.Add oTArc2
        Set oTArc2 = Nothing
    Else
        Set oTArc2 = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, arrPt(2).X, arrPt(2).y, arrPt(2).z, arrPt(3).X, arrPt(3).y, arrPt(3).z)
        m_OutputColl.AddOutput "line_", oTArc2
        oCollection.Add oTArc2
        Set oTArc2 = Nothing
    End If
    
'    Top Line 2
    Set oTLine2 = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, arrPt(3).X, arrPt(3).y, arrPt(3).z, arrPt(4).X, arrPt(4).y, arrPt(4).z)
    m_OutputColl.AddOutput "line_", oTLine2
    oCollection.Add oTLine2
    Set oTLine2 = Nothing
    
'    Top Arc 1
    If parOrientation = 1 Then
        oNormalVector.Set 0, 0, -1
    Else
        oNormalVector.Set 0, 0, 1
    End If
    Set oTArc1 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(m_OutputColl.ResourceManager, _
                                                      oCenterPoint.X, oCenterPoint.y, oCenterPoint.z, _
                                                     oNormalVector.X, oNormalVector.y, oNormalVector.z, _
                                                     arrPt(4).X, arrPt(4).y, arrPt(4).z, _
                                                    arrPt(1).X, arrPt(1).y, arrPt(1).z)

    m_OutputColl.AddOutput "line_", oTArc1
    oCollection.Add oTArc1
    Set oTArc1 = Nothing
'
    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Set oComplexStr = PlaceTrCString(arrPt(1), oCollection)

'   Project the complexstring to create the body of the platform
    oNormalVector.Set 0, 0, -1
    Set ObjPlatformBody = PlaceProjection(m_OutputColl, oComplexStr, oNormalVector, parPlatformHeight1, False)
    m_OutputColl.AddOutput "line_", ObjPlatformBody
    Set ObjPlatformBody = Nothing

' plane3d
    Dim ObjTopSurface As IngrGeom3D.Plane3d
    Set ObjTopSurface = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, arrPt(1).X, arrPt(1).y, arrPt(1).z, 0, 0, 1)
    ObjTopSurface.AddBoundary oComplexStr

    Dim nholes As Integer
    Dim objCircle As IngrGeom3D.Circle3d
    Dim circleNormal As New AutoMath.DVector
        
    Dim oElements As IJElements
    Dim objCStr   As IngrGeom3D.ComplexString3d
    Set oElements = New JObjectCollection
       
 'Make the holes
    For nholes = 1 To NumberofHoles1
            circleNormal.Set 0, 0, 1
            Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, _
                        arrPt(4 + nholes).X, arrPt(4 + nholes).y, arrPt(4 + nholes).z, _
                        circleNormal.X, circleNormal.y, circleNormal.z, _
                        HoleRadius(nholes))
                        m_OutputColl.AddOutput "line_", objCircle
            oElements.Add objCircle
            Set objCStr = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
            ObjTopSurface.AddHole objCStr
            oElements.Clear

            oNormalVector.Set 0, 0, -1
            Set ObjPlatformBody = PlaceProjection(m_OutputColl, objCStr, oNormalVector, parPlatformHeight1, False)
            m_OutputColl.AddOutput "line_", ObjPlatformBody
            objCStr.RemoveCurve True
    Next nholes
        
    Set ObjPlatformBody = Nothing
    Set objCircle = Nothing
    
    m_OutputColl.AddOutput "line_", ObjTopSurface
    Set ObjTopSurface = Nothing

    Dim Count As Integer
    For Count = 1 To oCollection.Count
        oCollection.Remove 1
    Next Count
'
    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing

' ---- bottom surface --------------------------------------------------------
'
    For Count = 1 To 7
        arrPt(Count).z = -parPlatformHeight1
    Next Count
    oCenterPoint.Set 0, 0, -parPlatformHeight1
'
    If parOrientation = 1 Then
    oNormalVector.Set 0, 0, 1
    Else
    oNormalVector.Set 0, 0, -1
    End If

'    Create the bottom surface plane
    Dim oComplexString1 As New IngrGeom3D.ComplexString3d
    Dim obotCollection As New Collection
    Dim ObjBottSurface As Object
'
'    Bottom Line 1
    Set oTLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPt(1).X, arrPt(1).y, arrPt(1).z, arrPt(2).X, arrPt(2).y, arrPt(2).z)
    obotCollection.Add oTLine1
    Set oTLine1 = Nothing
'
'    Bottom Arc 2
    If parPlatformType = 1 Then
'    Top Arc 2
    Set oTArc2 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                          oCenterPoint.X, oCenterPoint.y, oCenterPoint.z, _
                                                         oNormalVector.X, oNormalVector.y, oNormalVector.z, _
                                                         arrPt(2).X, arrPt(2).y, arrPt(2).z, _
                                                        arrPt(3).X, arrPt(3).y, arrPt(3).z)
    obotCollection.Add oTArc2
    Set oTArc2 = Nothing
'
    Else
'
    Set oTArc2 = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPt(2).X, arrPt(2).y, arrPt(2).z, arrPt(3).X, arrPt(3).y, arrPt(3).z)
    obotCollection.Add oTArc2
    Set oTArc2 = Nothing
    End If
'
'
'    Top Line 2
    Set oTLine2 = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPt(3).X, arrPt(3).y, arrPt(3).z, arrPt(4).X, arrPt(4).y, arrPt(4).z)
    obotCollection.Add oTLine2
    Set oTLine2 = Nothing

    If parOrientation = 1 Then
    oNormalVector.Set 0, 0, -1
    Else
    oNormalVector.Set 0, 0, 1
    End If
'
    Set oTArc1 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                      oCenterPoint.X, oCenterPoint.y, oCenterPoint.z, _
                                                     oNormalVector.X, oNormalVector.y, oNormalVector.z, _
                                                     arrPt(4).X, arrPt(4).y, arrPt(4).z, _
                                                    arrPt(1).X, arrPt(1).y, arrPt(1).z)

    obotCollection.Add oTArc1
    Set oTArc1 = Nothing

    Set oComplexString1 = PlaceTrCString(arrPt(1), obotCollection)
'
'    Create an infinite plane and appy the complex string as boundary
    Set ObjBottSurface = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, arrPt(1).X, arrPt(1).y, arrPt(1).z, 0, 0, -1)
    ObjBottSurface.AddBoundary oComplexString1

'-----------
    For nholes = 1 To NumberofHoles1
        circleNormal.Set 0, 0, -1
        Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        arrPt(4 + nholes).X, arrPt(4 + nholes).y, arrPt(4 + nholes).z, _
                        circleNormal.X, circleNormal.y, circleNormal.z, _
                        HoleRadius(nholes))
        oElements.Add objCircle
        Set objCStr = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
        ObjBottSurface.AddHole objCStr

        oElements.Clear
        objCStr.RemoveCurve True
    Next nholes
     
    Set objCircle = Nothing
    m_OutputColl.AddOutput "line_", ObjBottSurface
    Set ObjBottSurface = Nothing

    For Count = 1 To obotCollection.Count
        obotCollection.Remove 1
    Next Count
    oComplexString1.RemoveCurve True
'
' ---------------Section 2 ----------------------------------------------------------------

If parNumPlatform = 2 Then

    If parOrientation = 1 Then
        x11 = parAngle2
        x12 = parAngle2
        H(4) = HoleAngle4
        H(5) = HoleAngle5
        H(6) = HoleAngle6
    Else
        x11 = -parAngle2
        x12 = -parAngle2
        H(4) = -HoleAngle4
        H(5) = -HoleAngle5
        H(6) = -HoleAngle6
    
    End If
    
    If parOrientation = 1 Then
        end_angle = parAngle3
    Else
        end_angle = -parAngle3
    End If
    
    If parOrientation = 1 Then
        oNormalVector.Set 0, 0, 1
    Else
        oNormalVector.Set 0, 0, -1
    End If

    oCenterPoint.Set 0, 0, 0
    If parPlatformType = 1 Then
        x_rad = Abs(dOrigintoPlatform + parPlatformWidth2)
    Else
        x_rad = (dOrigintoPlatform + parPlatformWidth2) / Cos((end_angle - x11) / 2)
        x_rad = Abs(x_rad)
    End If
    
    arrPt(1).Set (dOrigintoPlatform) * Cos(end_angle - t1), -(dOrigintoPlatform) * Sin(end_angle - t1), 0
    arrPt(2).Set x_rad * Cos(end_angle - t1), -x_rad * Sin(end_angle - t1), 0

    arrPt(3).Set x_rad * Cos(x12 - t1), -x_rad * Sin(x12 - t1), 0
    arrPt(4).Set (dOrigintoPlatform) * Cos(x11 - t1), -(dOrigintoPlatform) * Sin(x11 - t1), 0
    
    'Holes for second sections
    arrPt(5).Set (HolePosition4) * Cos(H(4) - t1), -(HolePosition4) * Sin(H(4) - t1), 0
    arrPt(6).Set (HolePosition5) * Cos(H(5) - t1), -(HolePosition2) * Sin(H(5) - t1), 0
    arrPt(7).Set (HolePosition6) * Cos(H(6) - t1), -(HolePosition3) * Sin(H(6) - t1), 0

    
    'Top Line 1
    Set oTLine1 = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, arrPt(1).X, arrPt(1).y, arrPt(1).z, arrPt(2).X, arrPt(2).y, arrPt(2).z)
    m_OutputColl.AddOutput "line_", oTLine1
       oCollection.Add oTLine1
    Set oTLine1 = Nothing
          
    If parPlatformType = 1 Then
    'Top Arc 2
        Set oTArc2 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(m_OutputColl.ResourceManager, _
                                                          oCenterPoint.X, oCenterPoint.y, oCenterPoint.z, _
                                                         oNormalVector.X, oNormalVector.y, oNormalVector.z, _
                                                         arrPt(2).X, arrPt(2).y, arrPt(2).z, _
                                                        arrPt(3).X, arrPt(3).y, arrPt(3).z)
        m_OutputColl.AddOutput "line_", oTArc2
        oCollection.Add oTArc2
        Set oTArc2 = Nothing
    Else
        Set oTArc2 = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, arrPt(2).X, arrPt(2).y, arrPt(2).z, arrPt(3).X, arrPt(3).y, arrPt(3).z)
        m_OutputColl.AddOutput "line_", oTArc2
        oCollection.Add oTArc2
        Set oTArc2 = Nothing
    End If
'   Top Line 2
    Set oTLine2 = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, arrPt(3).X, arrPt(3).y, arrPt(3).z, arrPt(4).X, arrPt(4).y, arrPt(4).z)
    m_OutputColl.AddOutput "line_", oTLine2
    oCollection.Add oTLine2
    Set oTLine2 = Nothing
'    'Top Arc 1
    If parOrientation = 1 Then
        oNormalVector.Set 0, 0, -1
    Else
        oNormalVector.Set 0, 0, 1
    End If
    Set oTArc1 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(m_OutputColl.ResourceManager, _
                                                      oCenterPoint.X, oCenterPoint.y, oCenterPoint.z, _
                                                     oNormalVector.X, oNormalVector.y, oNormalVector.z, _
                                                     arrPt(4).X, arrPt(4).y, arrPt(4).z, _
                                                    arrPt(1).X, arrPt(1).y, arrPt(1).z)

    m_OutputColl.AddOutput "line_", oTArc1
    oCollection.Add oTArc1
    Set oTArc1 = Nothing
    Set oComplexStr = PlaceTrCString(arrPt(1), oCollection)

'   Project the complexstring to create the body of the platform
    oNormalVector.Set 0, 0, -1
    Set ObjPlatformBody = PlaceProjection(m_OutputColl, oComplexStr, oNormalVector, parPlatformHeight2, False)
    m_OutputColl.AddOutput "line_", ObjPlatformBody
    Set ObjPlatformBody = Nothing

' plane3d
    Set ObjTopSurface = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, arrPt(1).X, arrPt(1).y, arrPt(1).z, 0, 0, 1)
    ObjTopSurface.AddBoundary oComplexStr

 'Make the holes
    For nholes = 1 To NumberofHoles2
            circleNormal.Set 0, 0, 1
            Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, _
                        arrPt(4 + nholes).X, arrPt(4 + nholes).y, arrPt(4 + nholes).z, _
                        circleNormal.X, circleNormal.y, circleNormal.z, _
                        HoleRadius(3 + nholes))
                        m_OutputColl.AddOutput "line_", objCircle
            oElements.Add objCircle
            Set objCStr = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
            ObjTopSurface.AddHole objCStr
            oElements.Clear

            oNormalVector.Set 0, 0, -1
            Set ObjPlatformBody = PlaceProjection(m_OutputColl, objCStr, oNormalVector, parPlatformHeight2, False)
            m_OutputColl.AddOutput "line_", ObjPlatformBody
            objCStr.RemoveCurve True
    Next nholes

    Set ObjPlatformBody = Nothing
    Set objCircle = Nothing
   
    m_OutputColl.AddOutput "line_", ObjTopSurface
    Set ObjTopSurface = Nothing
    
    For Count = 1 To oCollection.Count
        oCollection.Remove 1
    Next Count

    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing
'
' '---- bottom surface --------------------------------------------------------
'
    For Count = 1 To 4
        arrPt(Count).z = -parPlatformHeight2
    Next Count
    oCenterPoint.Set 0, 0, -parPlatformHeight2
'
    If parOrientation = 1 Then
        oNormalVector.Set 0, 0, 1
    Else
        oNormalVector.Set 0, 0, -1
    End If
'    'Create the bottom surface plane
'    'Bottom Line 1
    Set oTLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPt(1).X, arrPt(1).y, arrPt(1).z, arrPt(2).X, arrPt(2).y, arrPt(2).z)
    obotCollection.Add oTLine1
    Set oTLine1 = Nothing
'
'   Bottom Arc 2
    If parPlatformType = 1 Then
    'Top Arc 2
        Set oTArc2 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                          oCenterPoint.X, oCenterPoint.y, oCenterPoint.z, _
                                                         oNormalVector.X, oNormalVector.y, oNormalVector.z, _
                                                         arrPt(2).X, arrPt(2).y, arrPt(2).z, _
                                                        arrPt(3).X, arrPt(3).y, arrPt(3).z)
        obotCollection.Add oTArc2
        Set oTArc2 = Nothing
    Else
        Set oTArc2 = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPt(2).X, arrPt(2).y, arrPt(2).z, arrPt(3).X, arrPt(3).y, arrPt(3).z)
        obotCollection.Add oTArc2
        Set oTArc2 = Nothing
    End If
    'Top Line 2
    Set oTLine2 = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPt(3).X, arrPt(3).y, arrPt(3).z, arrPt(4).X, arrPt(4).y, arrPt(4).z)
    obotCollection.Add oTLine2
    Set oTLine2 = Nothing

    If parOrientation = 1 Then
        oNormalVector.Set 0, 0, -1
    Else
        oNormalVector.Set 0, 0, 1
    End If

    Set oTArc1 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                      oCenterPoint.X, oCenterPoint.y, oCenterPoint.z, _
                                                     oNormalVector.X, oNormalVector.y, oNormalVector.z, _
                                                     arrPt(4).X, arrPt(4).y, arrPt(4).z, _
                                                    arrPt(1).X, arrPt(1).y, arrPt(1).z)

    obotCollection.Add oTArc1
    Set oTArc1 = Nothing

    Set oComplexString1 = PlaceTrCString(arrPt(1), obotCollection)

    'Create an infinite plane and appy the complex string as boundary
    Set ObjBottSurface = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, arrPt(1).X, arrPt(1).y, arrPt(1).z, 0, 0, -1)
    ObjBottSurface.AddBoundary oComplexString1
    
        For nholes = 1 To NumberofHoles2
        circleNormal.Set 0, 0, -1
        Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        arrPt(4 + nholes).X, arrPt(4 + nholes).y, arrPt(4 + nholes).z, _
                        circleNormal.X, circleNormal.y, circleNormal.z, _
                        HoleRadius(3 + nholes))
        oElements.Add objCircle
        Set objCStr = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
        ObjBottSurface.AddHole objCStr

        oElements.Clear
        objCStr.RemoveCurve True
    Next nholes

    m_OutputColl.AddOutput "line_", ObjBottSurface
    Set objCircle = Nothing
    Set ObjBottSurface = Nothing
    For Count = 1 To obotCollection.Count
        obotCollection.Remove 1
    Next Count

    oComplexString1.RemoveCurve True
    
End If
    Set oComplexString1 = Nothing
    Set obotCollection = Nothing

    Set oCollection = Nothing
    Set oCenterPoint = Nothing
    Set circleNormal = Nothing
    Set oNormalVector = Nothing
    Set oElements = Nothing
    Set objCStr = Nothing
    Set oTLine1 = Nothing
    Set oTArc2 = Nothing
    Set oTLine2 = Nothing
    Set oTArc1 = Nothing
    For Count = 1 To 7
        Set arrPt(Count) = Nothing
    Next Count
    Set geomFactory = Nothing
    Exit Sub

ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
End Sub

Private Function Arcsin(ByVal X As Double) As Double
    'Inverse Sine
    If Abs(X) = 1 Then
        Arcsin = Atn(1) * Sgn(X) * 2
    Else
        Arcsin = Atn(X / Sqr(-X * X + 1))
    End If
End Function

Private Function Arccos(ByVal X As Double) As Double
    ' Inverse Cosine
    If Abs(X) = 1 Then
        Arccos = Atn(1) * (1 - Sgn(X)) * 4
    Else
        Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    End If
End Function

