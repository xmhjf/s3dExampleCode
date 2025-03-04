Attribute VB_Name = "basCPFLevTransmitter"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   basCPFLevTransmitter.cls
'   Author: Veena
'   Creation Date:  Friday, oct 6 2006
'
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   23.May.2008     VRK     CR-142762: Provide instrument transmitter and pressure transmitter symbols
'******************************************************************************
Option Explicit

Public Type InputType
    name        As String
    description As String
    properties  As IMSDescriptionProperties
    uomValue    As Double
End Type

Public Type OutputType
    name            As String
    description     As String
    properties      As IMSDescriptionProperties
    Aspect          As SymbolRepIds
End Type

Public Type AspectType
    name                As String
    description         As String
    properties          As IMSDescriptionProperties
    AspectId            As SymbolRepIds
End Type

'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub
Public Sub ReportUnanticipatedError(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub
Public Function RetrievePipeOD_1(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                ByRef lPipeDiam As Double)

    'Dim oPipePort       As GSCADNozzleEntities.IJDPipePort
    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    
    On Error Resume Next
    
    Set oCollection = partInput.GetNozzles()
    For pPortIndex = 1 To oCollection.Size
        
        Set oPipePort = oCollection.Item(pPortIndex)
        If Not oPipePort Is Nothing Then 'There is a cable nozzle which donot implement IJCatalogPipePort
            If oPipePort.PortIndex = index Then
                lPipeDiam = oPipePort.PipingOutsideDiameter
            End If
        End If
    Next pPortIndex
    
    Set oPipePort = Nothing
    
End Function
