VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BrktReinforceXSec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-------------------------------------------------------------------------------
'  Copyright (C) 2008, Intergraph Corporation.  All rights reserved.
'
'  FILE:  BrktReinforceXSec.cls
'
'  DESCRIPTION:
'  Rule to set the cross section size of the stiffener or ER reinforcement
'  on a bracket (tripping or by plane)
'
'  ORIGINAL AUTHOR:  Bill Cox
'
'-------------------------------------------------------------------------------

Implements IJBracketReinforcementXSecRule

Private Const E_FAIL = -2147467259
Private Const MODULE = "S:\ShipStructure\Data\BracketReinforcementRules\BrktReinforceXSec.cls"


Private Sub IJBracketReinforcementXSecRule_ComputeReinforcementValues(ByVal pBracketPlate As Object, ByVal pBracketPlateSO As Object, ByVal pReinforcement As Object)
                
    ' check type of stiffener to see if ER or BS
    If TypeOf pReinforcement Is IJERSystem Then
        
        Dim oERSystem As IJERSystem
        Set oERSystem = pReinforcement
        oERSystem.Position = ER_OnEdgeCentered
    
        ' set cross section for ER
        SetCrossSection pReinforcement, "FB", "F 100x10", LeftWeb, 25
        
    Else ' must be buckling stiffener
    
        ' set cross section for BS
        SetCrossSection pReinforcement, "FB", "F 100x10", BottomSideOfInferiorFlange, 3
    End If

End Sub

' Function:
'    CreateCrossSection
' Abstract:
'   Create a Cross section to be applied to the profile
' Description:
'   Creates a hard code profile section
Private Function CreateCrossSection(ByVal strSectionType As String, _
                                    ByVal strSectionName As String) As IJCrossSection
                                                      
    Dim oRefDataService As RefDataMiddleServices.StructServices
    Dim xsecColl As IJDCollection
    Dim oTempCrossSection As IJDAttributes
    Dim lSize As Long
    Dim collIndex As Integer
    Dim strName As String
    
    On Error GoTo ErrorHandler
    
    Set oRefDataService = New RefDataMiddleServices.StructServices
    Set xsecColl = oRefDataService.GetCrossSections("ShipShapes", strSectionType)
    Set oRefDataService = Nothing
    
    lSize = xsecColl.Size
    
    Dim i As Integer
    For i = 1 To lSize
        Set oTempCrossSection = xsecColl.Item(i)
        If oTempCrossSection.CollectionOfAttributes("ISTRUCTCrossSection").Item("SectionName").Value = strSectionName Then
            collIndex = i
            Exit For
        End If
    Next i
    
    Set CreateCrossSection = xsecColl.Item(collIndex)
    Set xsecColl = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
              
Private Sub SetCrossSection(oProfile As Object, _
                            ByVal strCrossSectionType As String, _
                            ByVal strCrossSectionName As String, _
                            ByVal eMountingFace As ProfileFaceName, _
                            ByVal loadPoint As Long)
                           
    Dim oIJProfile As IJProfile
    Dim oIJProfileAttributes As IJProfileAttributes
    Dim oCrossSection As IJCrossSection
    
    On Error GoTo ErrorHandler
    
    Set oIJProfile = oProfile
'    If oIJProfile Is Nothing Then
'        Err.Raise E_INVALIDARG, MODULE & ":" & "SetCrossSection", "Profile not found"
'    End If
    
    Set oIJProfileAttributes = New GSCADCreateModifyUtilities.ProfileUtils
    Set oCrossSection = CreateCrossSection(strCrossSectionType, strCrossSectionName)
    If oCrossSection Is Nothing Then
        Err.Raise E_FAIL, MODULE & ":" & "SetCrossSection", "Failed to get CrossSection"
    End If
    
    oIJProfileAttributes.SetCrossSection oIJProfile, oCrossSection, eMountingFace, loadPoint
    
    Set oIJProfile = Nothing
    Set oIJProfileAttributes = Nothing
    Set oCrossSection = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


