Attribute VB_Name = "basGeom3dPH"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'   basGeom3d.bas
'   Author:
'   Creation Date:
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   6 June 2003     Symbol Team -          Added functions RetrieveNozzleParameters, SmallerDim and ReportUnanticipatedError2
'                                          If Symbol uses only basGeom3d.bas file, for error report call function  "ReportUnanticipatedError2"
'                                          If Symbol uses only FullyParametricFun.bas file, for error report call function  "ReportUnanticipatedError3"
'                                          If Symbol uses both basGeom3d.bas and FullyParametricFun.bas files, for error report call function  "ReportUnanticipatedError2"
'                                          12th June 2003 MS  TR 42250: Added the statements "Exit Function" for both RetrieveNozzleParameters and SmallerDim.
'   09.Jul.2003     SymbolTeam(India)      Copyright Information, Header  is Added.
'   23.Jul.2003     SymbolTeam(India)      Added PlaceTrapezoidWithPlanes function to create a trapezoid with 6 planes.
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox
'   26.Jul.2004     ACM                  DI-61828 changed declaration of pipeport from IJDPipePort to IJCatalogPipePort
'   28.Nov.2004     JFF                    Created by duplication and edit from basGeom3d for SmartOccurrence
'                                          Stripped off the non nozzle creation functions and rename the other with PH
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Const MODULE = "Geom3dPH"
Const TOLERANCE = 0.0000001

Public Function CreateNozzlePH(nozzleIndex As Integer, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                lDir As AutoMath.DVector, _
                                lPos As AutoMath.DPosition) As IJDNozzle

''' This function places Nozzle based on 2 parameters:
''' direction and placePoint
''' 2 first parameters (partInput and objOutPutColl) are from symbol machinery
'''' This is example of nozzles output
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -lFaceToFace / 2 - hcgs, 0, 0
'''    oDir.Set -1, 0, 0
'''
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing

''===========================
''Construction of nozzle
''===========================

    Dim iLogicalDistPort    As IJLogicalDistPort
    Dim iDistribPort        As IJDistribPort
    
    Dim oPipePort           As IJDPipePort
    Dim NozzlePHFactory     As NozzlePHFactory
    Dim oNozzle             As IJDNozzle
    
    Set NozzlePHFactory = New NozzlePHFactory
    Set oNozzle = NozzlePHFactory.CreatePipeNozzlePHFromPart(partInput, nozzleIndex, _
                                        False, objOutputColl.ResourceManager)
    Set NozzlePHFactory = Nothing
    Set iLogicalDistPort = oNozzle
    Set iDistribPort = oNozzle
    
    iDistribPort.SetDirectionVector lDir
    Set oPipePort = oNozzle
    oNozzle.Length = oPipePort.FlangeOrHubThickness
          
' Position of the nozzle should be the connect point of the nozzle
    iLogicalDistPort.SetCenterLocation lPos
     
    Set CreateNozzlePH = oNozzle

    Set oNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set oPipePort = Nothing
    
 End Function

'Create Nozzle by defining the length when Length is different than Flange Thickness
Public Function CreateNozzlePHWithLength _
( _
        nozzleIndex As Integer, _
        ByRef partInput As PartFacelets.IJDPart, _
        ByVal objOutputColl As Object, _
        lDir As AutoMath.DVector, _
        lPos As AutoMath.DPosition, _
        NozzleLength As Double _
) As IJDNozzle

''===========================
''Construction of nozzle
''===========================

    Dim iLogicalDistPort    As IJLogicalDistPort
    Dim iDistribPort        As IJDistribPort
    
    Dim oPipePort           As IJDPipePort
    Dim NozzlePHFactory     As NozzlePHFactory
    Dim oNozzle             As IJDNozzle
    
    Set NozzlePHFactory = New NozzlePHFactory
    Set oNozzle = NozzlePHFactory.CreatePipeNozzlePHFromPart(partInput, nozzleIndex, _
                                            False, objOutputColl.ResourceManager)
    Set NozzlePHFactory = Nothing
    Set iLogicalDistPort = oNozzle
    Set iDistribPort = oNozzle
    
    iDistribPort.SetDirectionVector lDir
    Set oPipePort = oNozzle
    oNozzle.Length = NozzleLength
          
' Position of the nozzle should be the connect point of the nozzle
    iLogicalDistPort.SetCenterLocation lPos
     
    Set CreateNozzlePHWithLength = oNozzle

    Set oNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set oPipePort = Nothing
    
End Function

Public Function CreateNozzlePHJustaCircle(nozzleIndex As Integer, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                lDir As AutoMath.DVector, _
                                lPos As AutoMath.DPosition) As IJDNozzle

''' This function places Nozzle based on 2 parameters:
''' direction and placePoint
''' 2 first parameters (partInput and objOutPutColl) are from symbol machinery
'''' This is example of nozzles output
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -lFaceToFace / 2 - hcgs, 0, 0
'''    oDir.Set -1, 0, 0
'''
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing

''===========================
''Construction of nozzle
''===========================

    Dim iLogicalDistPort    As IJLogicalDistPort
    Dim iDistribPort        As IJDistribPort
    
    Dim oPipePort           As IJDPipePort
    Dim NozzlePHFactory     As NozzlePHFactory
    Dim oNozzle             As IJDNozzle
    
    Set NozzlePHFactory = New NozzlePHFactory
    Set oNozzle = NozzlePHFactory.CreatePipeNozzlePHFromPart(partInput, nozzleIndex, _
                                        True, objOutputColl.ResourceManager)
    Set NozzlePHFactory = Nothing
    Set iLogicalDistPort = oNozzle
    Set iDistribPort = oNozzle
    
    iDistribPort.SetDirectionVector lDir
    Set oPipePort = oNozzle
    oNozzle.Length = oPipePort.FlangeOrHubThickness
          
' Position of the nozzle should be the connect point of the nozzle
    iLogicalDistPort.SetCenterLocation lPos
     
    Set CreateNozzlePHJustaCircle = oNozzle

    Set oNozzle = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set oPipePort = Nothing
    
 End Function

''''Create Cable Tray Port
Public Function CreateCableTrayPortPH(ByRef oPart As PartFacelets.IJDPart, _
    dNozzleIndex As Long, _
    oBasePt As AutoMath.DPosition, _
    oAxis As AutoMath.DVector, oRadial As AutoMath.DVector, _
    ByVal objOutputColl As Object) As IJCableTrayPortOcc
    ' This subroutine creates a Cable Tray Port  and sets it's position and direction
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory
    Dim NullObj As Object
    Dim oDistribPort As IJDistribPort
    Dim oLogicalDistPort As IJLogicalDistPort
    Dim oCableTrayPort As IJCableTrayPortOcc
    
    Const METHOD = "CreateCableTrayPortPH:"

    On Error GoTo ErrHandler
    Set oCableTrayPort = oNozzlePHFactory.CreateCableTrayNozzlePHFromPart(oPart, dNozzleIndex, _
                                                                objOutputColl.ResourceManager)
    Set oLogicalDistPort = oCableTrayPort
    Set oDistribPort = oCableTrayPort
    

    oLogicalDistPort.SetCenterLocation oBasePt
    
    oDistribPort.SetDirectionVector oAxis
    oDistribPort.SetRadialOrient oRadial
    
    Set CreateCableTrayPortPH = oCableTrayPort
        
    Set oNozzlePHFactory = Nothing
    Set oDistribPort = Nothing
    Set oLogicalDistPort = Nothing
    Set oCableTrayPort = Nothing
    Set oAxis = Nothing
    Set oRadial = Nothing
    Set oBasePt = Nothing
    Exit Function
    
ErrHandler:
  ReportUnanticipatedError2 MODULE, METHOD
End Function

Public Function CreateConduitNozzlePH(oBasePt As AutoMath.DPosition, oAxis As AutoMath.DVector, ByVal objOutputColl As Object, ByRef oPart As PartFacelets.IJDPart, dNozzleIndex As Long) As IJConduitPortOcc
    ' This subroutine creates a ConduitNozzle  and sets it's position and direction
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory
    Dim NullObj As Object
    Dim oDistribPort As IJDistribPort
    Dim oLogicalDistPort As IJLogicalDistPort
    Dim oConduitNozzle As IJConduitPortOcc
    Const METHOD = "CreateConduitNozzlePH:"

    On Error GoTo ErrHandler
    
    Set oConduitNozzle = oNozzlePHFactory.CreateConduitNozzlePHFromPart(oPart, dNozzleIndex, _
                                                                objOutputColl.ResourceManager)
    Set oLogicalDistPort = oConduitNozzle
    Set oDistribPort = oConduitNozzle
    

    oLogicalDistPort.SetCenterLocation oBasePt
    
    oDistribPort.SetDirectionVector oAxis
    
    Set CreateConduitNozzlePH = oConduitNozzle
        
    Set oNozzlePHFactory = Nothing
    Set oDistribPort = Nothing
    Set oLogicalDistPort = Nothing
    Set oConduitNozzle = Nothing
    Exit Function
    
ErrHandler:
  ReportUnanticipatedError2 MODULE, METHOD
End Function
