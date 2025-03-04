VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   ProgID:         SP3DSteamTrapAssembly.SteamTA
'   Author:         MP
'   Creation Date:  Wednesday, Oct 15 2008
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'    15.Oct.2008    MP     CR-151135  Provide steam trap fitting unit symbols per Yarway catalog
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double
Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim PipeDiam1 As Double
    Dim PipeDiam2 As Double
    Dim PipeDiam3 As Double
    Dim FlangeThick1 As Double
    Dim FlangeThick2 As Double
    Dim FlangeThick3 As Double
    Dim CptOffset1 As Double
    Dim CptOffset2 As Double
    Dim CptOffset3 As Double
    Dim FlangeDiam1 As Double
    Dim FlangeDiam2 As Double
    Dim FlangeDiam3 As Double
    Dim Depth1 As Double
    Dim Depth2 As Double
    Dim Depth3 As Double
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnpoint As New AutoMath.DPosition
    
    Dim lPartDataBasis As Long
    Dim oPipeComponent As IJDPipeComponent

    Dim iOutput As Long
    
    Dim ObjInsulatedBody1 As Object
    Dim ObjInsulatedBody2 As Object
    Dim objInsPort1 As Object
    Dim objInsport2 As Object
    Dim objInsPort3 As Object
    
    Dim parL As Double 'FacetoFace
    Dim parH As Double  'OperatorHeight
    Dim parB As Double  'TrapLength1
    Dim parW As Double  'OperatorDiameter
    Dim parInsulationThickness As Double 'InsulationThickness
    Dim parK As Double  'TrapLength2
    
    Dim oDirvec As AutoMath.DVector
    Set oDirvec = New AutoMath.DVector
    
    Dim dValveDia As Double
    Dim dBody2Dia As Double
    Dim dInsDia1 As Double
    Dim dInsDia2 As Double
    Dim dInsDia As Double
    Dim dInsDia3 As Double
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parL = arrayOfInputs(2) 'FacetoFace
    parH = arrayOfInputs(3) 'OperatorHeight
    parB = arrayOfInputs(4) 'TrapLength1
    parW = arrayOfInputs(5) 'OperatorDiameter
    parInsulationThickness = arrayOfInputs(6)
    
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    If lPartDataBasis = STP_Cock_Bypass Then
        parK = arrayOfInputs(7) 'Traplength2
    End If
    
    iOutput = 0
    Dim dAngle As Double
    dAngle = PI / 4
    
    If lPartDataBasis <= 1 Or lPartDataBasis = STA_DEFAULT Or lPartDataBasis = STA_Cock Or _
                                   lPartDataBasis = STP_Cock_Bypass Then
        'insert code for insulation body1
        oStPoint.Set -parL / 2, 0, 0
        oEnpoint.Set parL / 2, 0, 0
  
        RetrieveParameters 1, oPartFclt, m_OutputColl, PipeDiam1, FlangeThick1, FlangeDiam1, CptOffset1, Depth1
        RetrieveParameters 2, oPartFclt, m_OutputColl, PipeDiam2, FlangeThick2, FlangeDiam2, CptOffset2, Depth2
        

        dBody2Dia = 0.8 * PipeDiam1

        dValveDia = 1.25 * PipeDiam1
        
        dInsDia = dValveDia + 2 * parInsulationThickness
        
        Set ObjInsulatedBody1 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dInsDia, True)
        
        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody1
        Set ObjInsulatedBody1 = Nothing
        Set oStPoint = Nothing
        Set oEnpoint = Nothing
        
        'insert code for insulated body2
        If lPartDataBasis = STA_Cock Or lPartDataBasis = STP_Cock_Bypass Then
            oStPoint.Set -0.5 * parB, 0, 0
            oEnpoint.Set -(0.5 * parB - Cos(dAngle) * parB), -Cos(dAngle) * parB, 0
        Else
            oStPoint.Set -0.5 * parB, 0, 0
            oEnpoint.Set -0.3 * parB, -0.2 * parB, 0
        End If
        
        Set ObjInsulatedBody2 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dInsDia, True)
        
        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody2
        Set ObjInsulatedBody2 = Nothing
        Set oStPoint = Nothing
        Set oEnpoint = Nothing
        
        'insert code for insulation on bypass
        If lPartDataBasis = STP_Cock_Bypass Then
            Dim oInsBypass1 As Object
            Dim oInsBypass2 As Object
            Dim oInsBypass3 As Object
            Dim oInsBypassElbow1 As Object
            Dim oInsBypassElbow2 As Object
            
            'code for insbypass1
            oStPoint.Set -parL / 3, 0, 0
            oEnpoint.Set -parL / 3, 0, -parK * 2 / 3
            dInsDia = dValveDia + 2 * parInsulationThickness
            Set oInsBypass1 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dInsDia, True)
            
            'set the output
            m_OutputColl.AddOutput "InsBypass1", oInsBypass1
            Set oInsBypass1 = Nothing
            Set oStPoint = Nothing
            Set oEnpoint = Nothing
            
            'code for insbyass elbow1
            Dim oCircle As IngrGeom3D.Circle3d
            Dim oCenterPoint As New AutoMath.DPosition
            Dim oGeomFactory    As IngrGeom3D.GeometryFactory
                
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            oCenterPoint.Set -parL / 3 + parK / 3, 0, -2 * parK / 3
            oDirvec.Set 0, -1, 0
            Dim dInsDiaMax As Double
            dInsDiaMax = PipeDiam1 / 2 + (parK / 3)
            If CmpDblGreaterthan(dInsDia, dInsDiaMax) Then
                dInsDia = dInsDiaMax - LINEAR_TOLERANCE
            End If
            
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, -parL / 3, 0, -parK * 2 / 3, 0, 0, 1, dInsDia / 2)
                
            Set oInsBypassElbow1 = PlaceRevolution(m_OutputColl, oCircle, oDirvec, oCenterPoint, PI / 2, True)
                
            'set the output
            m_OutputColl.AddOutput "InsBypassElbow1", oInsBypassElbow1
            Set oInsBypassElbow1 = Nothing
            Set oCenterPoint = Nothing
            Set oDirvec = Nothing
            Set oCircle = Nothing
            
            'code for insbypass2
            oStPoint.Set -parL / 3 + parK / 3, 0, -parK
            oEnpoint.Set parL / 3 - parK / 3, 0, -parK
            dInsDia = dValveDia + 2 * parInsulationThickness
            
            Set oInsBypass2 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dInsDia, True)
            
            'set the output
            m_OutputColl.AddOutput "InsBypass2", oInsBypass2
            Set oInsBypass2 = Nothing
            Set oStPoint = Nothing
            Set oEnpoint = Nothing
            
            'code for InsBypass elbow2
            oCenterPoint.Set parL / 3 - parK / 3, 0, -2 * parK / 3
            Set oDirvec = New AutoMath.DVector
            oDirvec.Set 0, 1, 0
            If CmpDblGreaterthan(dInsDia, dInsDiaMax) Then
                dInsDia = dInsDiaMax - LINEAR_TOLERANCE
            End If
            
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parL / 3, 0, -parK * 2 / 3, 0, 0, 1, dInsDia / 2)
                
            Set oInsBypassElbow2 = PlaceRevolution(m_OutputColl, oCircle, oDirvec, oCenterPoint, PI / 2, True)
                
            'set the output
            m_OutputColl.AddOutput "InsBypassElbow2", oInsBypassElbow2
            Set oInsBypassElbow2 = Nothing
            Set oCenterPoint = Nothing
            Set oDirvec = Nothing
            Set oCircle = Nothing
                            
            
            'code for InsBypass3
            oStPoint.Set parL / 3, 0, 0
            oEnpoint.Set parL / 3, 0, -parK * 2 / 3
            dInsDia = dValveDia + 2 * parInsulationThickness
            
            Set oInsBypass3 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dInsDia, True)
            
            m_OutputColl.AddOutput "InsBypass3", oInsBypass3
            Set oInsBypass3 = Nothing
            Set oStPoint = Nothing
            Set oEnpoint = Nothing
            
        End If
        
        'insert code for insulated port1
        If CmpDblGreaterthan(FlangeDiam1, FlangeDiam2) Then
            dInsDia1 = FlangeDiam1 + parInsulationThickness * 2
        Else
            dInsDia1 = FlangeDiam2 + parInsulationThickness * 2
        End If
        
        oStPoint.Set -parL / 2 + FlangeThick1 + parInsulationThickness, 0, 0
        oEnpoint.Set -parL / 2, 0, 0
        
        Set objInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dInsDia1, True)
        
        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort1
        Set objInsPort1 = Nothing
        Set oStPoint = Nothing
        Set oEnpoint = Nothing
        
        'insert code for insport2
         If CmpDblGreaterthan(FlangeDiam2, PipeDiam2) Then
            dInsDia2 = FlangeDiam2 + parInsulationThickness * 2
        Else
            dInsDia2 = PipeDiam2 + parInsulationThickness * 2
        End If
        
        oStPoint.Set parL / 2 - FlangeThick2 - parInsulationThickness, 0, 0
        oEnpoint.Set parL / 2, 0, 0
        
        Set objInsport2 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dInsDia2, True)
        
        'set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsport2
        Set objInsport2 = Nothing
        Set oStPoint = Nothing
        Set oEnpoint = Nothing
        
        
        'insert code for insport3
        RetrieveParameters 3, oPartFclt, m_OutputColl, PipeDiam3, FlangeThick3, FlangeDiam3, CptOffset3, Depth3
        Dim dInsLength As Double
        
        If CmpDblGreaterthan(FlangeThick3, LINEAR_TOLERANCE) Then
            dInsLength = FlangeThick3 + parInsulationThickness
        Else
            dInsLength = 0.001
        End If
        If CmpDblGreaterthan(FlangeDiam3, PipeDiam3) Then
            dInsDia3 = FlangeDiam3 + parInsulationThickness * 2
        Else
            dInsDia3 = PipeDiam3 + parInsulationThickness * 2
        End If
        If lPartDataBasis = STA_Cock Or lPartDataBasis = STP_Cock_Bypass Then
            oStPoint.Set -(0.5 * parB - (parB - dInsLength) * Cos(dAngle)), (-parB + dInsLength) * Cos(dAngle), 0
            oEnpoint.Set -(0.5 * parB - Cos(dAngle) * parB), -Cos(dAngle) * parB, 0
        Else
            oStPoint.Set -0.3 * parB - dInsLength * Cos(dAngle), -0.2 * parB + dInsLength * Cos(dAngle), 0
            oEnpoint.Set -0.3 * parB, -0.2 * parB, 0
        End If
               
        Set objInsPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, dInsDia3, True)
        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort3
        Set objInsPort3 = Nothing
        Set oStPoint = Nothing
        Set oEnpoint = Nothing
    End If
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
