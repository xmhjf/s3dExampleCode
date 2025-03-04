Attribute VB_Name = "basCEccentric"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   basCEccentric.bas
'   Author:          NN
'   Creation Date:  Saturday, Jul 21 2001
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   22.Feb.2007     RRK                   TR-113129 Added RetrieveParametersWithInsidePipeDiameter function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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
Public Sub ReportUnanticipatedError(InModule As String, InMethod As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

Public Function RetrieveParametersWithInsidePipeDiameter(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByVal objOutputColl As Object, _
                                ByRef lPipeDiam As Double, ByRef lFlangeThick As Double, _
                                ByRef lFlangeDiam As Double, ByRef lcptoffset As Double, _
                                ByRef lDepth As Double, ByRef lInsidePipeDiameter As Double)

    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    
    Set oCollection = partInput.GetNozzles()
    
    For pPortIndex = 1 To oCollection.Size
        Set oPipePort = oCollection.Item(pPortIndex)
            If oPipePort.PortIndex = index Then
                lPipeDiam = oPipePort.PipingOutsideDiameter
                lFlangeThick = oPipePort.FlangeOrHubThickness
                lFlangeDiam = oPipePort.FlangeOrHubOutsideDiameter
                lcptoffset = oPipePort.FlangeProjectionOrSocketOffset
                lDepth = oPipePort.SeatingOrGrooveOrSocketDepth
                lInsidePipeDiameter = oPipePort.PipingInsideDiameter
                Exit For
            End If
    Next pPortIndex
    
    Set oPipePort = Nothing
    
End Function

