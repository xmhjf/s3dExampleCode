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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  30 March, 2000
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy           who                    change description
'   -----------          -----                    ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As IJDOutputCollection, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim PipeDia As Double
    Dim FlangeThick As Double
    Dim CptOffset As Double
    Dim FlangeDia As Double
    Dim Depth As Double
    
    Dim iOutput     As Double
    
    Dim FacetoCenter As Double
    Dim OperatorHeight As Double
    Dim BodyDia As Double
    Dim OperatorLength As Double
    
    Dim OperatorDia As Double
    Dim HandleDia As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    FacetoCenter = arrayOfInputs(2)
    OperatorHeight = arrayOfInputs(3)
    BodyDia = arrayOfInputs(4)
    OperatorLength = arrayOfInputs(5)

    RetrieveParameters 1, oPartFclt, m_OutputColl, _
                       PipeDia, FlangeThick, FlangeDia, CptOffset, Depth

' For Test -------------------------------------------------------------------
'''    FacetoCenter = arrayOfInputs(1)
'''    OperatorHeight = arrayOfInputs(2)
'''    BodyDia = arrayOfInputs(3)
'''    OperatorLength = arrayOfInputs(4)
'''
'''    PipeDia = 0.1
'''    FlangeThick = 0.018
'''    FlangeDia = 0.21
'''    CptOffset = 0.001
' ----------------------------------------------------------------------------
    
    OperatorDia = BodyDia / 4
    HandleDia = OperatorDia / 2
    
    iOutput = 0

    Dim CenterPos As New AutoMath.DPosition
    CenterPos.Set 0, 0, 0
    
 ' Place body of valve
 ' -------------------------------------------------------------------------------------

    Dim objBodyV As IngrGeom3D.Sphere3d
    Set objBodyV = PlaceSphere(m_OutputColl, CenterPos, BodyDia / 2)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBodyV
    
    Set objBodyV = Nothing
    
''' ' Place Inlet, Outlet
''' ' --------------------------------------------------------------------------------
'''    Dim StartPos As New AutoMath.DPosition
'''    Dim EndPos As New AutoMath.DPosition
'''
'''    StartPos.Set CenterPos.x - FacetoCenter  + FlangeThick, CenterPos.y, CenterPos.z
'''    EndPos.Set StartPos.x + FacetoCenter/2 - FlangeThick * 2, StartPos.y, StartPos.z
'''
'''    Dim objInOutlet As IngrGeom3D.Projection3d
'''    Set objInOutlet = PlaceCylinder(m_OutputColl, StartPos, EndPos, PipeDia, True)
'''
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInOutlet
'''
'''    Set objInOutlet = Nothing
'''
''' ' Place 3-Way
''' '---------------------------------------------------------------------------------------
'''    StartPos.Set CenterPos.x, CenterPos.y, CenterPos.z
'''    EndPos.Set StartPos.x, StartPos.y, StartPos.z - FacetoCenter  + FlangeThick
'''
'''    Dim obj3Way As IngrGeom3D.Projection3d
'''    Set obj3Way = PlaceCylinder(m_OutputColl, StartPos, EndPos, PipeDia, True)
'''
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), obj3Way
'''
         
 ' Place Stem
 '---------------------------------------------------------------------------------------
    Dim StartPos As New AutoMath.DPosition
    Dim EndPos As New AutoMath.DPosition
    
    StartPos.Set CenterPos.x, CenterPos.y, CenterPos.z
    EndPos.Set StartPos.x, StartPos.y, StartPos.z + OperatorHeight
    
    Dim objStem As IngrGeom3D.Projection3d
    Set objStem = PlaceCylinder(m_OutputColl, StartPos, EndPos, OperatorDia, True)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStem
    
    Set objStem = Nothing

 ' Place Handle 1
 ' ------------------------------------------------------------------------------------------
    StartPos.Set CenterPos.x - OperatorLength / 4, _
                 CenterPos.y, _
                 CenterPos.z + OperatorHeight - HandleDia / 2
    EndPos.Set StartPos.x + OperatorLength, StartPos.y, StartPos.z
    
    Dim objHandle As IngrGeom3D.Projection3d
    Set objHandle = PlaceCylinder(m_OutputColl, StartPos, EndPos, HandleDia, True)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandle
    
    Set objHandle = Nothing
    Set StartPos = Nothing
    Set EndPos = Nothing
 
 ' Place Handle 2
 ' ------------------------------------------------------------------------------------------
    StartPos.Set CenterPos.x, _
                 CenterPos.y, _
                 CenterPos.z + OperatorHeight - HandleDia / 2
    EndPos.Set StartPos.x, StartPos.y + OperatorLength / 4, StartPos.z
    
    Set objHandle = PlaceCylinder(m_OutputColl, StartPos, EndPos, HandleDia, True)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandle
    
    Set objHandle = Nothing
    Set StartPos = Nothing
    Set EndPos = Nothing
 
 ' Place Nozzle 1
 ' ------------------------------------------------------------------------------------------
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -FacetoCenter - CptOffset + Depth, 0, 0
    oDir.Set -1, 0, 0

    Set objNozzle = CreateNozzleWithLength( _
                        1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, FacetoCenter)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

 ' Place Nozzle 2
 ' ------------------------------------------------------------------------------------------
    RetrieveParameters 2, oPartFclt, m_OutputColl, _
                       PipeDia, FlangeThick, FlangeDia, CptOffset, Depth

    oPlacePoint.Set FacetoCenter + CptOffset - Depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzleWithLength( _
                        2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, FacetoCenter)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

 ' Place Nozzle 3
 ' ------------------------------------------------------------------------------------------
    RetrieveParameters 3, oPartFclt, m_OutputColl, _
                       PipeDia, FlangeThick, FlangeDia, CptOffset, Depth

    oPlacePoint.Set 0, FacetoCenter + CptOffset - Depth, 0
    oDir.Set 0, 1, 0

    Set objNozzle = CreateNozzleWithLength( _
                        3, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, FacetoCenter)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

Private Function CreateNozzleWithLength _
( _
        nozzleIndex As Integer, _
        ByRef partInput As PartFacelets.IJDPart, _
        ByVal objOutputColl As Object, _
        lDir As AutoMath.DVector, _
        lPos As AutoMath.DPosition, _
        NozzleLength As Double _
) As GSCADNozzleEntities.IJDNozzle

''===========================
''Construction of nozzle
''===========================

    Dim iLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort        As GSCADNozzleEntities.IJDistribPort
    
    Dim oPipePort           As GSCADNozzleEntities.IJDPipePort
    Dim NozzleFactory       As GSCADNozzleEntities.NozzleFactory
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    
    Set NozzleFactory = New GSCADNozzleEntities.NozzleFactory
    Set oNozzle = NozzleFactory.CreatePipeNozzleFromPart(partInput, nozzleIndex, _
                                            False, objOutputColl.ResourceManager)
    Set NozzleFactory = Nothing
    Set iLogicalDistPort = oNozzle
    Set iDistribPort = oNozzle
    
    iDistribPort.SetDirectionVector lDir
    Set oPipePort = oNozzle
    oNozzle.Length = NozzleLength
          
' Position of the nozzle should be the connect point of the nozzle
    iLogicalDistPort.SetCenterLocation lPos
     
    Set CreateNozzleWithLength = oNozzle

    Set oNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set oPipePort = Nothing
    
 End Function
