VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileCacheFuncs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------------------'
'    Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'Abstract
'
'This class module contains the code that provides custom methods used by the profile symbols
'They get called by the flavor manager to determine when to calculate a new representation
'
'History
'
'    M Holderer         12/03/08         Created
'
'--------------------------------------------------------------------------------------------'

Option Explicit

Private Const MODULE = "ProfileCacheFuncs"

Public Sub CMCacheSpecThickness(oInputCM As Object, bArgToCache As Boolean, oToConvert As Object, ByRef oOutput As Object)
    Const METHOD = "CMCacheSpecThickness"
    On Error GoTo ErrorHandler

    If bArgToCache Then
   
        Dim iIJProxy As IJDProxy
        Dim iInsulationSpec As IStructInsulationSpec
        
        Set iIJProxy = oToConvert
        Set iInsulationSpec = iIJProxy.Source

         'Convert the spec object to a parameter content thickness value
        Dim oPC As IJDParameterContent
        Set oPC = New DParameterContent
         
        oPC.Type = igValue
        oPC.UomValue = iInsulationSpec.Thickness
        Set oOutput = oPC
        Set oPC = Nothing
    
    Else
        Set oOutput = Nothing
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMCacheMaterial(oInputCM As Object, bArgToCache As Boolean, oToConvert As Object, ByRef oOutput As Object)
    Const METHOD = "CMCacheMaterial"
    On Error GoTo ErrorHandler

    If bArgToCache Then
       
        'Return a fixed string since the material does not matter to the profile
        Dim oPC As IJDParameterContent
        Set oPC = New DParameterContent
         
        oPC.Type = igString
        oPC.String = "Material"
        Set oOutput = oPC
        Set oPC = Nothing
    
    Else
        Set oOutput = Nothing
    End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMCacheMemberType(oInputCM As Object, bArgToCache As Boolean, oToConvert As Object, ByRef oOutput As Object)
    Const METHOD = "CMCacheMaterial"
    On Error GoTo ErrorHandler

    If bArgToCache Then

        'Return a fixed string since the member type does not matter to the profile
        Dim oPC As IJDParameterContent
        Set oPC = New DParameterContent

        oPC.Type = igString
        oPC.String = "MemberType"
        Set oOutput = oPC
        Set oPC = Nothing

    Else
        Set oOutput = Nothing
    End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'The insulation flavor manager needs check whether a flavor exists when:
'   (1) the member changes cross-section, or
'   (2) updating the cross-section via bulkload and resynch
'Return a string that is the member's cross-section's symbol's flavor Oid

Public Sub CMCacheMemberCrossSection(oInputCM As Object, bArgToCache As Boolean, oToConvert As Object, ByRef oOutput As Object)
    Const METHOD = "CMCacheMemberCrossSection"
    On Error GoTo ErrorHandler

    If bArgToCache Then
       
        Dim status As StructInsulationServicesStatus
        Dim iSPSCrossSection As iSPSCrossSection
        Dim iMemberPart As ISPSMemberPartPrismatic
        Dim iInsulationServices As iInsulationServices

        Set iInsulationServices = New StructInsulations.InsulationServices

        Dim oPC As IJDParameterContent
        Dim strMemberCrossSectionFlavorOid As String
    
        Set iMemberPart = oToConvert
        Set iSPSCrossSection = iMemberPart.CrossSection
        status = iInsulationServices.StringForMemberSymbolCache(iMemberPart.CrossSection, strMemberCrossSectionFlavorOid)

        Set oPC = New DParameterContent
        oPC.Type = igString
        oPC.String = strMemberCrossSectionFlavorOid
        Set oOutput = oPC
        Set oPC = Nothing
    
    Else
        Set oOutput = Nothing
    End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

''''Public Sub CMGetReferences(pSymbolDefinition As IJDSymbolDefinition, pSymbolOccurrence As IJDSymbol, ByRef ppEnumJDArgument As IEnumJDArgument)
''''    Const METHOD = "CMGetReferences"
''''    On Error GoTo ErrorHandler
''''
''''    Dim oRefColl As Object
''''
''''    Set oRefColl = GetRefCollection(pSymbolOccurrence)
''''    If Not oRefColl Is Nothing Then
''''        Set ppEnumJDArgument = oRefColl
''''    End If
''''
''''    Exit Sub
''''
''''ErrorHandler:
''''    HandleError MODULE, METHOD
''''End Sub

Private Sub HandleError(sModule As String, sMethod As String)
    Dim oEditErrors As IJEditErrors
    
    Set oEditErrors = New JServerErrors
    If Not oEditErrors Is Nothing Then
        oEditErrors.AddFromErr Err, "", sMethod, sModule
    End If
    Set oEditErrors = Nothing
End Sub

