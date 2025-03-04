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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006-2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RRK
'   Creation Date:  12 July, 2006
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      Source: Warren Controls, http://www.warrencontrols.com/html/pdf/1800ProductSpec.pdf
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   11.FEB.2008     PK      CR-135567       Enhanced the existing 3 way diverter valve to support dimensional basis for
'                                               (i)Face-to-center dimension basis
'                                               (ii)Face-to-Face dimension and Face-to-Center dimension for port 3
'                                               (iii)Asymmetrical Face-to-Center dimension basis
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As IJDOutputCollection, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim PipeDia1        As Double
    Dim PipeDia2        As Double
    Dim PipeDia3        As Double
    Dim FlangeThick1    As Double
    Dim FlangeThick2    As Double
    Dim FlangeThick3    As Double
    Dim FlangeDia1      As Double
    Dim FlangeDia2      As Double
    Dim FlangeDia3      As Double
    Dim Depth           As Double
    Dim CptOffset       As Double
       
    Dim parFacetoFace       As Double
    Dim parFacetoCenter     As Double
    Dim parInstrumentHeight As Double
    
    Dim parFace1toCenter    As Double
    Dim parFace2toCenter    As Double
    Dim parFace3toCenter    As Double
    Dim parAngle            As Double
    
    Dim dFace1toCenter      As Double
    Dim dFace2toCenter      As Double
    Dim dFace3toCenter      As Double
    Dim dAngle              As Double
    Dim dBodyRadius         As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFacetoFace = arrayOfInputs(2)
    'parFacetoCenter = arrayOfInputs(3)
    'parInstrumentHeight = arrayOfInputs(4)
    'parFace1toCenter = arrayOfInputs(5)
    'parFace2toCenter = arrayOfInputs(6)
    'parFace3toCenter = arrayOfInputs(7)
    'parAngle=arrayOfInputs(8)
    
    Dim oPipeComponent      As IJDPipeComponent
    Dim lPartdatabasis      As Long
    
    Set oPipeComponent = oPartFclt
    lPartdatabasis = oPipeComponent.PartDataBasis
    
    Select Case lPartdatabasis
    
        Case Is <= 1                            'default case
            parFacetoFace = arrayOfInputs(2)
            parFacetoCenter = arrayOfInputs(3)
            parInstrumentHeight = arrayOfInputs(4)
            dFace1toCenter = parFacetoFace / 2
            dFace2toCenter = parFacetoFace / 2
            dFace3toCenter = parFacetoCenter
                      
        Case 11                                 'Face-to-Face dimension and Face-to-Center dimension for port 3
            parFacetoFace = arrayOfInputs(2)
            parFace3toCenter = arrayOfInputs(7)
            parAngle = arrayOfInputs(8)
            dFace1toCenter = parFacetoFace / 2
            dFace2toCenter = parFacetoFace / 2
            dFace3toCenter = parFace3toCenter
            dAngle = parAngle
            
        Case 10                                 'Face-to-center dimension basis
            parFacetoCenter = arrayOfInputs(3)
            parAngle = arrayOfInputs(8)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dFace3toCenter = parFacetoCenter
            dAngle = parAngle
            
        Case 13                                 'Asymmetrical Face-to-Center dimension basis
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            parFace3toCenter = arrayOfInputs(7)
            parAngle = arrayOfInputs(8)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dFace3toCenter = parFace3toCenter
            dAngle = parAngle
            If dFace1toCenter = 0 Then
                dFace1toCenter = dFace2toCenter
            ElseIf dFace2toCenter = 0 Then
                dFace2toCenter = dFace1toCenter
            End If
            
        Case Else
            GoTo ErrorLabel:
            
    End Select
    
    'Place body of valve (Output 1)
    '-------------------------------------------------------------------------------------
    If lPartdatabasis <= 1 Then
        Dim objSphere As Object
        Dim CenterPos As AutoMath.DPosition
        Set CenterPos = New AutoMath.DPosition
        CenterPos.Set 0, 0, 0
                
        'Assumption: Spherical body diameter is taken to be 80% of parFacetoFace
        
        dBodyRadius = (0.8 * parFacetoFace) / 2
        Set objSphere = PlaceSphere(m_OutputColl, CenterPos, dBodyRadius)
        
        'Set the output
        m_OutputColl.AddOutput "Body", objSphere
        Set objSphere = Nothing
        Set CenterPos = Nothing
    ElseIf lPartdatabasis = 10 Or 11 Or 13 Then
        
        RetrieveParameters 1, oPartFclt, m_OutputColl, _
                       PipeDia1, FlangeThick1, FlangeDia1, CptOffset, Depth
        RetrieveParameters 2, oPartFclt, m_OutputColl, _
                       PipeDia2, FlangeThick2, FlangeDia2, CptOffset, Depth
                       
        Dim oCylinder   As Object
        Dim dCylLength  As Double
        Dim oCylStPoint As AutoMath.DPosition
        Dim oCylEnPoint As AutoMath.DPosition
        Set oCylStPoint = New AutoMath.DPosition
        Set oCylEnPoint = New AutoMath.DPosition
        
        'Assumption: Cylinder body radius is taken to be 40% of the minimum of Face 1 to Center and
        'Face 2 to Center and length of cylinder is assumed to be 2 times the maximum of pipe
        'diameter  1 and pipe diameter2
        If CmpDblLessThan(dFace1toCenter, dFace2toCenter) Then
            dBodyRadius = 0.4 * dFace1toCenter
        Else
            dBodyRadius = 0.4 * dFace2toCenter
        End If
        
        If CmpDblGreaterthan(PipeDia1, PipeDia2) Then
            dCylLength = 2 * PipeDia1
        Else
            dCylLength = 2 * PipeDia2
        End If
        
        oCylStPoint.Set 0, 0, dCylLength / 2
        oCylEnPoint.Set 0, 0, -dCylLength / 2
            
        Set oCylinder = PlaceCylinder(m_OutputColl, oCylStPoint, oCylEnPoint, 2 * dBodyRadius, True)
        
        'Set the output
        m_OutputColl.AddOutput "Body", oCylinder
        Set oCylinder = Nothing
        Set oCylStPoint = Nothing
        Set oCylEnPoint = Nothing
    End If
     
    If lPartdatabasis <= 1 Then
    
        'Place Cylinder 1 (Output 2)
        '---------------------------------------------------------------------------------------
        RetrieveParameters 3, oPartFclt, m_OutputColl, _
                           PipeDia3, FlangeThick3, FlangeDia3, CptOffset, Depth
        Dim stPoint As New AutoMath.DPosition
        Dim enPoint As New AutoMath.DPosition
        Dim dCylinderDia As Double
        Dim dMeetDistance As Double
    
        dCylinderDia = 1.2 * PipeDia3
        
        dMeetDistance = Sqr(Abs(dBodyRadius ^ 2 - (dCylinderDia / 2) ^ 2))
        
        'Assumption: Cylinder 1 height is taken as 60% of the parInstrumentHeight
        Dim dCyl1Height As Double
        
        dCyl1Height = 0.6 * parInstrumentHeight
        stPoint.Set 0, dMeetDistance, 0
        enPoint.Set 0, dCyl1Height, 0
        
        Dim objCylinder1 As IngrGeom3D.Projection3d
        Set objCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dCylinderDia, True)
        
        'Set the output
        m_OutputColl.AddOutput "Cylinder1", objCylinder1
        Set objCylinder1 = Nothing
    
        'Place Cylinder 2 (Output 3)
        ' ------------------------------------------------------------------------------------------
        'Assumption: Cylinder 2 height is taken to be 50% of the parFace1toCenter. However if this value
        'is greater than the Cylinder 1 height, then Cylinder 1 height is taken as Cylinder 2 height (Cylinder 1 and Cylinder 2 look symmetric)
        Dim dCyl2Height As Double

        stPoint.Set 0, -dMeetDistance, 0
        If CmpDblGreaterthan(0.5 * parFacetoCenter, dCyl1Height) Then
           dCyl2Height = dCyl1Height
        Else
           dCyl2Height = 0.5 * parFacetoCenter
        End If
        enPoint.Set 0, -dCyl2Height, 0
       
        Dim objCylinder2 As IngrGeom3D.Projection3d
        Set objCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dCylinderDia, True)
       
        'Set the output
        m_OutputColl.AddOutput "Cylinder2", objCylinder2
        Set objCylinder2 = Nothing
    
        'Place Flange 1 (Output 4)
        '------------------------------------------------------------------------------------------
        'Assumption: Two bolted flanges in the instrument are represented with Flange 1, Flange 1 thickness
        'is taken as 20 % of parInstrumentHeight. Flange diameter is taken to be 2 times the PipeDia3
        Dim dFlange1Thk As Double
    
        dFlange1Thk = 0.2 * parInstrumentHeight
        stPoint.Set 0, dCyl1Height, 0
        enPoint.Set 0, dCyl1Height + dFlange1Thk, 0
    
        Dim objFlange As IngrGeom3D.Projection3d
        Set objFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, 1.5 * PipeDia3, True)
    
        'Set the output
        m_OutputColl.AddOutput "Flange1", objFlange

        'Place Flange 2 (Output 5)
        '------------------------------------------------------------------------------------------
        stPoint.Set 0, -dCyl2Height, 0
        enPoint.Set 0, stPoint.y - dFlange1Thk, 0
        
        Set objFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, 1.5 * PipeDia3, True)
    
        'Set the output
        m_OutputColl.AddOutput "Flange2", objFlange
        Set objFlange = Nothing
        
        'Place Conical Part (Output 6)
        '------------------------------------------------------------------------------------------
        stPoint.Set 0, -dCyl2Height - dFlange1Thk, 0
        enPoint.Set 0, -parFacetoCenter, 0
    
        Dim objCone As Object
        'Assumption: Base radius of cone is taken to be 60% of PipeDia3
        Set objCone = PlaceCone(m_OutputColl, stPoint, enPoint, 0.6 * PipeDia3, PipeDia3 / 2, True)
    
        'Set the output
        m_OutputColl.AddOutput "Cone", objCone
        Set objCone = Nothing
        Set stPoint = Nothing
        Set enPoint = Nothing
     
    End If
    'Place Nozzle 1 (Output 1)
    '------------------------------------------------------------------------------------------
    RetrieveParameters 1, oPartFclt, m_OutputColl, _
                       PipeDia1, FlangeThick1, FlangeDia1, CptOffset, Depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -dFace1toCenter - CptOffset + Depth, 0, 0
    oDir.Set -1, 0, 0
    
    Dim dNozzleLength As Double
    dNozzleLength = dFace1toCenter - Sqr(Abs(dBodyRadius ^ 2 - (PipeDia1 / 2) ^ 2))
    If CmpDblLessThan(dNozzleLength, FlangeThick1) Then dNozzleLength = FlangeThick1
    Set objNozzle = CreateNozzleWithLength( _
                        1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
                        
    'Set the output
    m_OutputColl.AddOutput "Port1", objNozzle
    Set objNozzle = Nothing

    'Place Nozzle 2 (Output 2)
    '------------------------------------------------------------------------------------------
    RetrieveParameters 2, oPartFclt, m_OutputColl, _
                       PipeDia2, FlangeThick2, FlangeDia2, CptOffset, Depth

    oPlacePoint.Set dFace2toCenter + CptOffset - Depth, 0, 0
    oDir.Set 1, 0, 0

    dNozzleLength = dFace2toCenter - Sqr(Abs(dBodyRadius ^ 2 - (PipeDia2 / 2) ^ 2))
    If CmpDblLessThan(dNozzleLength, FlangeThick2) Then dNozzleLength = FlangeThick2
    Set objNozzle = CreateNozzleWithLength( _
                        2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
    
    'Set the output
    m_OutputColl.AddOutput "Port2", objNozzle
    Set objNozzle = Nothing

    'Place Nozzle 3 (Output 3)
    '------------------------------------------------------------------------------------------
    RetrieveParameters 3, oPartFclt, m_OutputColl, _
                       PipeDia3, FlangeThick3, FlangeDia3, CptOffset, Depth

    Dim dPort3Pos   As Double
    dPort3Pos = dFace3toCenter + CptOffset - Depth
    oPlacePoint.Set dPort3Pos * Sin(dAngle), -dPort3Pos * Cos(dAngle), 0
    oDir.Set Sin(dAngle), -Cos(dAngle), 0
    
    If lPartdatabasis <= 1 Then
        Set objNozzle = CreateNozzle( _
                        3, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint)
    Else
        dNozzleLength = dFace3toCenter
        Set objNozzle = CreateNozzleWithLength( _
                         3, oPartFclt, m_OutputColl, oDir, oPlacePoint, _
                         dNozzleLength)
    End If
                      
    'Set the output
    m_OutputColl.AddOutput "Port3", objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    'Place Valve Operator (Output 10)
    '------------------------------------------------------------------------------------------
    Dim oSymbolHelper As IJSymbolGeometryHelper
    Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
    oSymbolHelper.OutputCollection = m_OutputColl
    
    On Error Resume Next
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector
    
    oDirX.Set 1, 0, 0
    oDirY.Set 0, 1, 0
    oDirZ.Set 0, 0, 1
    
    Set oPipeComponent = oPartFclt
    On Error GoTo ErrorLabel
    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc   As IJPartOcc
    If Not oPipeComponent Is Nothing Then
        Set oOperatorPart = oPipeComponent.GetValveOperatorPart
        If Not oOperatorPart Is Nothing Then
            Dim OpOrigin As IJDPosition
            Set OpOrigin = New DPosition
            OpOrigin.Set 0, 0, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
            Set OpOrigin = Nothing
        End If
    End If
    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing
    
 Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
