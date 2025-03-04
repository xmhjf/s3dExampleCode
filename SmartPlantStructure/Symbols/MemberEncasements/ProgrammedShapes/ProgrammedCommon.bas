Attribute VB_Name = "ProgrammedCommon"
'---------------------------------------------------------------------------
'    Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'
'
'History
'    SS         10/29/08      Creation
'   MOH         12/31/08      Wrote MakeAttributedWireBody and use Erase for edgeids array
'---------------------------------------------------------------------------------------
Option Explicit

Public Const MODULE = "ProgrammedCommon"

Public Sub DefineEncasementSymbolOutput(strLibName As String, strLibProgId As String, pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    Const METHOD = "DefineEncasementSymbolOutput"
    On Error GoTo ErrorHandler
    
    Dim lLibCookie As Long
    Dim lMethodCookie As Long
    Dim LibDesc As DLibraryDescription
    Dim LibDescCustomMethods As IJDLibraryDescription
    
    'Defines the output of this symbol
    Dim oRep As IJDRepresentation
    Dim oOutput As IJDOutput

    ' Set the reference of myself as library
    Set LibDesc = pSymbolDefinition.IJDUserMethods.GetLibrary(strLibName)
    If LibDesc Is Nothing Then
        Set LibDescCustomMethods = New DLibraryDescription
        
        LibDescCustomMethods.Name = strLibName
        LibDescCustomMethods.Type = imsLIBRARY_IS_ACTIVEX
        LibDescCustomMethods.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
        LibDescCustomMethods.Source = strLibProgId
        pSymbolDefinition.IJDUserMethods.SetLibrary LibDescCustomMethods
        Set LibDesc = pSymbolDefinition.IJDUserMethods.GetLibrary(strLibName)
    End If

    If LibDesc Is Nothing Then
        WriteError METHOD, "IJUserMethods.GetLibrary failed for " & strLibName
        Exit Sub
    End If

    lLibCookie = LibDesc.Cookie
    lMethodCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMEvaluateEncasementCrossSection", lLibCookie)
    If lMethodCookie = 0 Then
        WriteError METHOD, "IJUserMethods.GetMethodCookie failed for " & strLibName & ", CMEvaluateEncasementCrossSection"
        Exit Sub
    End If

    Set oRep = New DRepresentation
    oRep.Name = ENCASEMENT_CROSSSECTION_REPNAME
    oRep.RepresentationId = 1
    oRep.Description = "Member Encasement"
    oRep.Properties = igCOLLECTION_VARIABLE
    oRep.IJDRepresentationStdCustomMethod.SetCMEvaluate lLibCookie, lMethodCookie

    Set oOutput = New DOutput
    oOutput.Name = ENCASEMENTSYMBOL_OUTPUTNAME
    oOutput.Description = "Cross Section Of Encasement"
    oRep.IJDOutputs.Add oOutput
    pSymbolDefinition.IJDRepresentations.Add oRep

    'as this symbol def has declared a graphic object as input
    ' GeomOption option will be set to igSYMBOL_GEOM_FIX_TO_ID by the symbol machinerary
    'Because of this the  outputs will be transformed during MDR and the Symbol geometry will
    ' end up in an incorrect location. So resetting the flag - DI226263
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' oOuter can be a wireBody ( already attributed ) or a GType geometry, and edgeIds are applied to it.
' elesHoles can be Nothing

Public Sub BuildEncasementSymbolOutput(oOuter As Object, edgeIds() As Long, _
                        elesHoles As IJElements, outputColl As IJDOutputCollection)

    Const METHOD = "BuildEncasementSymbolOutput"
    On Error GoTo ErrorHandler

    Dim oWireBody As Object
    Dim oProfile2dhelper As IJDProfile2dHelper
    Dim countSeg As Long, iSeg As Long, countHole As Long, iHole As Long
    Dim elesWires As IJElements
    Dim iSectionProfile As SectionProfile

    If TypeOf oOuter Is IJWireBody Then
        Set oWireBody = oOuter
    
    Else
        If elesHoles Is Nothing Then        ' make the wirebody persistent
            MakeAttributedWireBody outputColl.ResourceManager, oOuter, 1, edgeIds, oWireBody
        Else
            MakeAttributedWireBody Nothing, oOuter, 1, edgeIds, oWireBody
        End If
    End If

    If elesHoles Is Nothing Then
        outputColl.AddOutput ENCASEMENTSYMBOL_OUTPUTNAME, oWireBody
        Exit Sub
    End If

    Set oProfile2dhelper = New DProfile2dHelper
    Set elesWires = New JObjectCollection

    elesWires.Add oWireBody
    Set oWireBody = Nothing

    countHole = elesHoles.Count
    For iHole = 1 To countHole
        Set iSectionProfile = elesHoles(iHole)

        MakeAttributedWireBody Nothing, iSectionProfile, 0, edgeIds, oWireBody
    
        elesWires.Add oWireBody
        Set oWireBody = Nothing
    Next iHole

    Set oWireBody = oProfile2dhelper.MergeWireBodies(elesWires, outputColl.ResourceManager)

    outputColl.AddOutput ENCASEMENTSYMBOL_OUTPUTNAME, oWireBody

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub EvaluateContourCrossSection(pRepSCM As IJDRepresentationStdCustomMethod, bExposeTop As Boolean, _
                    bExposeRight As Boolean, bExposeBottom As Boolean, bExposeLeft As Boolean)
    Const METHOD = "EvaluateContourCrossSection"
    On Error GoTo ErrorHandler
 
    Dim oRepDuringGame As IMSSymbolEntities.IJDRepresentationDuringGame
    Dim oInput As IJDInput
    
    Dim countProfiles As Long
    Dim thickness As Double
    Dim iIJProxy As IJDProxy
    Dim iInsulationSpec As IStructInsulationSpec
    Dim iMemberPart As ISPSMemberPartPrismatic
    Dim countSeg As Long, iSeg As Long

    Dim ii As Long
    Dim xlo As Double, ylo As Double, xhi As Double, yhi As Double
    Dim elesProfiles As IJElements
    Dim XServices As New CrossSectionServices
    Dim iProfile As SectionProfile

    Dim oHoles As Object, elesProfileHoles As IJElements, elesAllHoles As IJElements
    Dim oComplexStringOffset As ComplexString3d
    
    'Dim pOutputs As IMSSymbolEntities.IJDOutputs
    Dim outputColl As IJDOutputCollection
    Dim oInsulation As IStructInsulation
    
    Set oRepDuringGame = pRepSCM
    Set outputColl = oRepDuringGame.outputCollection
    
    Set oInsulation = oRepDuringGame.definition.IJDDefinitionPlayerEx.PlayingSymbol

    Set oInput = oRepDuringGame.definition.IJDInputs.GetInputByIndex(1)
    Set iIJProxy = oInput.IJDInputDuringGame.Argument
    Set iInsulationSpec = iIJProxy.Source
    thickness = iInsulationSpec.thickness

    Set oInput = oRepDuringGame.definition.IJDInputs.GetInputByIndex(3)
    Set iMemberPart = oInput.IJDInputDuringGame.Argument
    
    XServices.GetProfiles iMemberPart.CrossSection.symbol, MEMBER_CROSSSECTION_REPNAME, elesProfiles

    Dim edgeIds() As Long
    Dim status As StructInsulations.StructInsulationServicesStatus
    Dim iServices As StructInsulations.IInsulationServices

    countProfiles = elesProfiles.Count

    For ii = 1 To countProfiles
        Set iProfile = elesProfiles(ii)
        
        '' accum all holes into one list
        Set oHoles = iProfile.Holes
        If Not oHoles Is Nothing Then
            Set elesProfileHoles = oHoles
            If elesAllHoles Is Nothing Then
                Set elesAllHoles = elesProfileHoles
            Else
                elesAllHoles.AddElements elesProfileHoles
            End If
            Set elesProfileHoles = Nothing
        End If
    
    Next ii

    Set iServices = New StructInsulations.InsulationServices
    
    If countProfiles = 1 Then
        
        Set iProfile = elesProfiles(1)

        status = iServices.CurveForContour(Nothing, iProfile, thickness, bExposeTop, bExposeRight, bExposeBottom, bExposeLeft, oComplexStringOffset, edgeIds)
        If status <> StructInsulationServices_Ok Then
            Select Case status
                Case Is = StructInsulationServices_UnexpectedError
                    oInsulation.ComputeStatus = StructInsulationInputHelper_UnexpectedError
                Case Is = StructInsulationServices_InvalidTypeOfObject
                    oInsulation.ComputeStatus = StructInsulationInputHelper_InvalidTypeOfObject
                Case Is = StructInsulationServices_InvalidInput
                    oInsulation.ComputeStatus = StructInsulationInputHelper_InvalidTypeOfObject
                Case Is = StructInsulationServices_DistanceExceedsCurveLength
                    oInsulation.ComputeStatus = StructInsulationInputHelper_BadGeometry
                Case Is = StructInsulationServices_CannotReadMemberCrossSectionSymbol
                    oInsulation.ComputeStatus = StructInsulationInputHelper_InvalidTypeOfObject
                Case Is = StructInsulationServices_CannotComputeInsulationCrossSection
                    oInsulation.ComputeStatus = StructInsulationInputHelper_BadGeometry
                Case Else
                    oInsulation.ComputeStatus = StructInsulationInputHelper_UnexpectedError
                    
            End Select
            Err.Raise E_FAIL, MODULE & " " & METHOD, "Cannot compute contour"
        End If

        BuildEncasementSymbolOutput oComplexStringOffset, edgeIds, elesAllHoles, outputColl
        
        Erase edgeIds

    Else

        Dim elesWires As IJElements
        Dim oWireBody As Object
        Dim oProfile2dhelper As IJDProfile2dHelper

        Set oProfile2dhelper = New DProfile2dHelper
        Set elesWires = New JObjectCollection

        For ii = 1 To countProfiles
    
            Set iProfile = elesProfiles(ii)
    
            status = iServices.CurveForContour(Nothing, iProfile, thickness, bExposeTop, bExposeRight, bExposeBottom, bExposeLeft, oComplexStringOffset, edgeIds)
            If status <> StructInsulationServices_Ok Then
                
                Select Case status
                    Case Is = StructInsulationServices_UnexpectedError
                        oInsulation.ComputeStatus = StructInsulationInputHelper_UnexpectedError
                    Case Is = StructInsulationServices_InvalidTypeOfObject
                        oInsulation.ComputeStatus = StructInsulationInputHelper_InvalidTypeOfObject
                    Case Is = StructInsulationServices_InvalidInput
                        oInsulation.ComputeStatus = StructInsulationInputHelper_InvalidTypeOfObject
                    Case Is = StructInsulationServices_DistanceExceedsCurveLength
                        oInsulation.ComputeStatus = StructInsulationInputHelper_BadGeometry
                    Case Is = StructInsulationServices_CannotReadMemberCrossSectionSymbol
                        oInsulation.ComputeStatus = StructInsulationInputHelper_InvalidTypeOfObject
                    Case Is = StructInsulationServices_CannotComputeInsulationCrossSection
                        oInsulation.ComputeStatus = StructInsulationInputHelper_BadGeometry
                Case Else
                    oInsulation.ComputeStatus = StructInsulationInputHelper_UnexpectedError
                        
                End Select
                Err.Raise E_FAIL, MODULE & " " & METHOD, "Cannot compute contour"
            End If
    
            MakeAttributedWireBody Nothing, oComplexStringOffset, 1, edgeIds, oWireBody

            Erase edgeIds

            elesWires.Add oWireBody
            Set oWireBody = Nothing
        
        Next ii

        If elesAllHoles Is Nothing Then
            Set oWireBody = oProfile2dhelper.MergeWireBodies(elesWires, outputColl.ResourceManager)
        Else
            Set oWireBody = oProfile2dhelper.MergeWireBodies(elesWires, Nothing)
        End If

        BuildEncasementSymbolOutput oWireBody, edgeIds, elesAllHoles, outputColl

    End If

    Exit Sub

ErrorHandler:
    Dim oEditErrors As IJEditErrors
    Set oEditErrors = New JServerErrors
    If Not oEditErrors Is Nothing Then
        oEditErrors.AddWithObjectInfo Err.Number, _
                                      Err.Source, _
                                      Err.Description, _
                                      oInsulation.ComputeStatus, _
                                      Err.HelpFile, _
                                      Err.HelpContext, _
                                      , , , , iMemberPart
    End If
    Set oEditErrors = Nothing
    'HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Sub

' oCurve can be a SectionProfile, or a GType curve complex string or other GType

Private Sub MakeAttributedWireBody(ResourceManager As Object, oCurve As Object, _
                                         InOut As Long, edgeIds() As Long, ByRef oWireBody As Object)
    Const METHOD = "MakeAttributedWireBody"
    On Error GoTo ErrorHandler

    Dim oProfile2dhelper As IJDProfile2dHelper
    Dim oIJGeometryMisc As IMSModelGeomOps.IJGeometryMisc
    
    Dim iSeg As Long, countSeg As Long
    Dim oComplexString As IJComplexString

    Dim edgeOprs() As Long
    Dim edgeOpts() As Long
    
    If TypeOf oCurve Is SectionProfile Then

        Dim localEdgeIds() As Long
        Dim iSectionProfile As SectionProfile
        Set iSectionProfile = oCurve
        Set oComplexString = iSectionProfile.Geometry
        countSeg = oComplexString.CurveCount
        ReDim localEdgeIds(countSeg)
        For iSeg = 0 To countSeg - 1
            localEdgeIds(iSeg) = iSectionProfile.EdgeId(iSeg + 1)
        Next iSeg

        MakeAttributedWireBody ResourceManager, oComplexString, InOut, localEdgeIds, oWireBody
        
        Erase localEdgeIds
    
        Exit Sub

    ElseIf TypeOf oCurve Is IJComplexString Then
        Set oComplexString = oCurve
        countSeg = oComplexString.CurveCount
    Else
        countSeg = 1
    End If

    Set oProfile2dhelper = New DProfile2dHelper
    Set oIJGeometryMisc = New IMSModelGeomOps.DGeomOpsMisc
    
    oIJGeometryMisc.CreateModelGeometryFromGType ResourceManager, oCurve, Nothing, oWireBody
    
    oProfile2dhelper.SetIntegerAttributeOnWireBodyLump oWireBody, "JSIn_Out", InOut
    
    ReDim edgeOprs(countSeg)
    ReDim edgeOpts(countSeg)
    For iSeg = 0 To countSeg - 1
        edgeOprs(iSeg) = iSeg + 1
        edgeOpts(iSeg) = 2000
    Next iSeg
    
    oProfile2dhelper.SetIntegerAttributesOnWireBodyEdges oWireBody, "JSXid", countSeg, edgeIds(0)
    oProfile2dhelper.SetIntegerAttributesOnWireBodyEdges oWireBody, "JSOpr", countSeg, edgeOprs(0)
    oProfile2dhelper.SetIntegerAttributesOnWireBodyEdges oWireBody, "JSOpt", countSeg, edgeOpts(0)

    Erase edgeOprs
    Erase edgeOpts

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Sub
