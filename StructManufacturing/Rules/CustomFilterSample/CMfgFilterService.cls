VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilterService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJSMCustomFilter


'    TR-CP-227633  Few delivered MFG rules cannot be recompiled on an End user machine
'    Removing the functionality since IJSMCustomFilter is changed recently.



''This function is callled when an assembly is piccked from the command.
''Through this function, user can send the needed plates/profiles/members of the assembly to be manufactured
''NOTE:Below implementation is default where we return all parts of the assembly as required
Private Function IJSMCustomFilter_GetCustomPartsFromAssemblyForMfg(ByVal pAssembly As GSCADMfgRulesDefinitions.IJAssembly) As IMSCoreCollections.IJElements


    Const METHOD = "IJSMCustomFilter_GetCustomPartsFromAssemblyForMfg"
    On Error GoTo ErrorHandler

    'make a collection of parts with needed ones
    Dim oPartCollection As IJElements

    Set oPartCollection = New JObjectCollection
    'Get children of assembly

    Dim oChildren As IJDTargetObjectCol
    Set oChildren = pAssembly.GetChildren

    If oChildren.Count > 0 Then
        Dim Index As Long
        Dim oItem As Object

        'Get each assembly child and add child to elements list
        For Index = 1 To oChildren.Count
            'Get next item
            Set oItem = oChildren.Item(Index)

            If TypeOf oItem Is IJAssembly Then 'Assembly or block is child
                'get parts in this assembly
                oPartCollection.AddElements IJSMCustomFilter_GetCustomPartsFromAssemblyForMfg(oItem)
            Else 'plate/profile or member
                If TypeOf oItem Is IJPlatePart Or TypeOf oItem Is IJProfilePart Or TypeOf oItem Is ISPSMemberPartPrismatic Then
                    oPartCollection.Add oItem
                End If
            End If
        Next
    End If

    Set IJSMCustomFilter_GetCustomPartsFromAssemblyForMfg = oPartCollection

    Exit Function
ErrorHandler:
    Err.Raise Err.Number

End Function

'Private Function IJSMCustomFilter_MfgPartCustomFilterString(ByVal eMfgPartType As GSCADMfgRulesDefinitions.FilterMfgObjectTypes) As String
''   make sure that there are white spaces between the " AND NOT "
'    Select Case eMfgPartType
'        Case ManufacturingPlateType
'            IJSMCustomFilter_MfgPartCustomFilterString = " AND NOT [MfgObjectCustomFilter.CMfgIsIdentical,IsPlatePartIdentical]"
'        Case ManufacturingProfileType
'            IJSMCustomFilter_MfgPartCustomFilterString = " AND NOT [MfgObjectCustomFilter.CMfgIsIdentical,IsProfilePartIdentical]"
'        Case ManufacturingMemberType
'            IJSMCustomFilter_MfgPartCustomFilterString = " AND NOT [MfgObjectCustomFilter.CMfgIsIdentical,IsProfilePartIdentical]"
'    End Select
'
'End Function


Private Function IJSMCustomFilter_MfgPartCustomFilterString() As String

End Function
