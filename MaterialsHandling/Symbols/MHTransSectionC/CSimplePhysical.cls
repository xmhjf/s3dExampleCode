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
'   Copyright (c) 2010, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author: BhupatiRaju
'   Creation Date: 16 Nov 2010
'   Description: Simple physical file for creating the Transition section 'A' of
'   the chute symbol to be used in the Layout desgn command of MHE.
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy         who         change description
'   -----------         -----       ------------------
'   16-Nov-2010         Neelima     Creating the file first time
'   07-Oct-2011         Nagesh      CR-CP-199489  As part of this CR symbol origin is moved to center of top face
'   15-Oct-2012         Shireesha   CR-CP-218458  Create the transition symbol for 2D & 3D symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CSimplePhysical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI As Double

Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
     Set m_oSymGeomHelper = New SymbolServices
     PI = 4 * Atn(1)
     
     Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext

End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt                       As PartFacelets.IJDPart
    Dim oCenterPos                      As IJDPosition
    Dim iOutput                         As Double
    Dim DblTransSecLength               As Double
    Dim DblTransSecWidth                As Double
    Dim DblTransSecDepth                As Double
    Dim DblTransSecConvergAngle         As Double
    Dim DblTransSecDivergAngle          As Double
    Dim DblTransSecOpeningLength        As Double
    Dim DblTransSecOpeningWidth         As Double
    Dim DblTempX                        As Double
    Dim DblTempY                        As Double
    Dim oTopPts(0 To 3)                 As IJDPosition
    Dim oBotPts(0 To 3)                 As IJDPosition
    Dim iCount                          As Integer
    Dim ObjTransSecColl                 As Collection
    Dim ObjPointColl                    As Collection
    Dim oGeomFactory                    As IngrGeom3D.GeometryFactory
    Dim oCurrentPosTransform            As IJDT4x4
    Dim pVector                         As IJDVector
    Dim oRotBotPts(0 To 3)              As IJDPosition
    Dim DblRotationAngle                As Double
    Dim oNewPos                         As IJDPosition
    Dim oPos                            As IJDPosition
    
    Set oNewPos = New DPosition
    Set oPos = New DPosition
    Set oCenterPos = New DPosition
    oCenterPos.Set 0, 0, 0
    
    Set pVector = New DVector
    Set oCurrentPosTransform = New DT4x4
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    DblTransSecLength = arrayOfInputs(2)         '1.25
    DblTransSecWidth = arrayOfInputs(3)          '1.7
    DblTransSecDepth = arrayOfInputs(4)          '1.2
    DblTransSecConvergAngle = arrayOfInputs(5)        '60
    DblTransSecDivergAngle = arrayOfInputs(6)         '70
    DblTransSecOpeningLength = arrayOfInputs(7)        '1.25
    DblTransSecOpeningWidth = arrayOfInputs(8)         '1.85
    DblRotationAngle = arrayOfInputs(9)         '0
    
    DblTempX = DblTransSecDepth / Round((Tan(DblTransSecConvergAngle)), 2)
    DblTempY = DblTransSecDepth / Round((Tan(DblTransSecDivergAngle)), 2)
    
    iOutput = 0
     
    m_oSymGeomHelper.outputCollection = m_outputColl
 
    ' Constructing the Points for front face of head section of chute.
    For iCount = 0 To 3
        Set oTopPts(iCount) = New DPosition
        Set oBotPts(iCount) = New DPosition
        Set oRotBotPts(iCount) = New DPosition
    Next iCount
    
'Start : CR-CP-199489

    oTopPts(0).Set -DblTransSecLength / 2, -DblTransSecWidth / 2, oCenterPos.z
    oTopPts(1).Set DblTransSecLength / 2, -DblTransSecWidth / 2, oCenterPos.z
    oTopPts(2).Set DblTransSecLength / 2, DblTransSecWidth / 2, oCenterPos.z
    oTopPts(3).Set -DblTransSecLength / 2, DblTransSecWidth / 2, oCenterPos.z
    
    Dim dblConvAngleX As Double
    Dim dblDivAngleY As Double
    
    dblConvAngleX = oTopPts(1).x - DblTempX
    dblDivAngleY = oTopPts(3).y - DblTempY
    
    oBotPts(0).Set dblConvAngleX - DblTransSecOpeningLength, dblDivAngleY - DblTransSecOpeningWidth, -DblTransSecDepth
    oBotPts(1).Set dblConvAngleX, oBotPts(0).y, -DblTransSecDepth
    oBotPts(2).Set dblConvAngleX, dblDivAngleY, -DblTransSecDepth
    oBotPts(3).Set oBotPts(0).x, dblDivAngleY, -DblTransSecDepth
        
    oNewPos.Set (oBotPts(0).x + oBotPts(2).x) / 2, (oBotPts(0).y + oBotPts(2).y) / 2, (oBotPts(0).y + oBotPts(2).y) / 2  'Get the centerPosition of Bottom Plane
        
    
'End : CR-CP-199489

'Start : Original Points

'    oTopPts(0).Set oCenterPos.x, oCenterPos.y, oCenterPos.z
'    oTopPts(1).Set oCenterPos.x + DblTransSecLength, oCenterPos.y, oCenterPos.z
'    oTopPts(2).Set oCenterPos.x + DblTransSecLength, oCenterPos.y + DblTransSecWidth, oCenterPos.z
'    oTopPts(3).Set oCenterPos.x, oCenterPos.y + DblTransSecWidth, oCenterPos.z
'
'    oBotPts(0).Set oTopPts(0).x - DblTempX, oTopPts(0).y - DblTempY, -(oTopPts(0).z + DblTransSecDepth)
'    oBotPts(1).Set oBotPts(0).x + DblTransSecOpeningLength, oBotPts(0).y, oBotPts(0).z
'    oBotPts(2).Set oBotPts(1).x, oBotPts(1).y + DblTransSecOpeningWidth, oBotPts(0).z
'    oBotPts(3).Set oBotPts(0).x, oBotPts(2).y, oBotPts(0).z

'End
    'Start CR-CP-218458
    If DblRotationAngle < 0 Then             'For Negative Angles
        DblRotationAngle = (2 * PI) + (DblRotationAngle)
    End If
    If DblRotationAngle >= PI And DblRotationAngle < (2 * PI) Then      'For 3 and 4th quadrant angles
        DblRotationAngle = DblRotationAngle - PI
    End If

    If DblRotationAngle > 0 Then
        For iCount = 0 To 3
            pVector.x = 0
            pVector.y = 0
            pVector.z = 1
            oCurrentPosTransform.LoadIdentity
            oCurrentPosTransform.Rotate DblRotationAngle, pVector
            ''''''''''''''''''''''''''''''''''
            'Rotate w.r.t Particular Positon(here oNewPos) instead of origin
            Set oPos = oCurrentPosTransform.TransformPosition(oNewPos)
            oCurrentPosTransform.IndexValue(12) = oNewPos.x - oPos.x
            oCurrentPosTransform.IndexValue(13) = oNewPos.y - oPos.y
            oCurrentPosTransform.IndexValue(14) = oNewPos.z - oPos.z
            Set oRotBotPts(iCount) = oCurrentPosTransform.TransformPosition(oBotPts(iCount))
            oBotPts(iCount).x = oRotBotPts(iCount).x
            oBotPts(iCount).y = oRotBotPts(iCount).y
            oBotPts(iCount).z = oRotBotPts(iCount).z
            ''''''''''''''''''''''''''''''''''''
        Next
        If DblRotationAngle > (PI / 2) And DblRotationAngle < PI Then
            For iCount = 0 To 2
              oBotPts(iCount + 1).x = oRotBotPts(iCount).x
              oBotPts(iCount + 1).y = oRotBotPts(iCount).y
              oBotPts(iCount + 1).z = oRotBotPts(iCount).z
            Next
            oBotPts(0).x = oRotBotPts(3).x
            oBotPts(0).y = oRotBotPts(3).y
            oBotPts(0).z = oRotBotPts(3).z
        End If
        
    End If
    
    If DblRotationAngle = 0 Or DblRotationAngle = (2 * PI) Then
      Set ObjTransSecColl = PlaceTrapezoidWithPlanes(m_outputColl, oTopPts, oBotPts)
    Else
      Set ObjTransSecColl = PlaceTrapezoidWithTriangles(m_outputColl, oTopPts, oBotPts)
    End If
    'End CR-CP-218458
    
   'Create the Head Section (Output 1)
    ''Set ObjTransSecColl = PlaceTrapezoidWithPlanes(m_outputColl, oTopPts, oBotPts)
    'Set the Output
    iOutput = iOutput + 1
    For iCount = 1 To ObjTransSecColl.Count
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjTransSecColl(iCount)
    Next iCount
    Set ObjTransSecColl = Nothing
    
    '============================================================================
    'Create the Reference Points on the bottoms face of head section (Output 3)
    '============================================================================
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set ObjPointColl = New Collection
    
    'Creating the points on all the four vertices of bottom face
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oTopPts(0).x, oTopPts(0).y, oTopPts(0).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oTopPts(1).x, oTopPts(1).y, oTopPts(1).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oTopPts(2).x, oTopPts(2).y, oTopPts(2).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oTopPts(3).x, oTopPts(3).y, oTopPts(3).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oBotPts(0).x, oBotPts(0).y, oBotPts(0).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oBotPts(1).x, oBotPts(1).y, oBotPts(1).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oBotPts(2).x, oBotPts(2).y, oBotPts(2).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oBotPts(3).x, oBotPts(3).y, oBotPts(3).z)
    
    'Set the output
    iOutput = iOutput + 1
    For iCount = 1 To ObjPointColl.Count
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPointColl(iCount)
    Next iCount
    
    For iCount = 0 To 3
        Set oTopPts(iCount) = Nothing
        Set oBotPts(iCount) = Nothing
        Set oRotBotPts(iCount) = Nothing
    Next
    Set oGeomFactory = Nothing
    Set ObjPointColl = Nothing
    Set oCurrentPosTransform = Nothing
    Set pVector = Nothing
    Set oNewPos = Nothing
    Set oPos = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

