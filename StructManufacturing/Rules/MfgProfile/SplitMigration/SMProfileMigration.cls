VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SMProfileMigration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''
''   Copyright (C) 2002, Intergraph Corporation. All rights reserved.
''
''   Abstract:
''              MfgProfileRule.cls
''   ProgID:
''              StrMfgMigrationRules.SMProfileMigration
''   Author:
''              Ingr StructMfg team.
''   Creation Date:
''              19 April 2006
''   Description:
''      CR#84861: Struct Mfg should check split notification and keep the GNEST lot number.
''      When a part has already been nested, it has a Lot number to link the part between
''      IntelliShip and GNEST.
''
''      In a simplest test, a Profile "P1" is split into two smaller Profiles "P2" and "P3" and "P1" will be deleted.
''      Old Profile P1, along with new Profiles P2 and P3 will be passed to a user defined VB rule.  This rule will be
''      bulkloaded into the catalog data in StructMfgSetting.xls spread sheet.
''
''   Change History:
''   dd.mmm.yyyy     who     change description
''   -----------     ---     ------------------
''
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'
Option Explicit
Implements IJMfgSplitMigration
Private Const MODULE = "MfgProfileMigration"


Private Function IJMfgSplitMigration_MigrateObject(ByVal pOldObj As Object, ByVal pReplacingObjColl As GSCADStructMfgGlobals.IJElements, ByVal pMfgObject As Object, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
'Arguments:
' pOldObj ==>  Profile "P1" which is being deleted.
'  pReplacingObjColl ==> Collection of replacing new Profiles ( P2 and P3 ).
' pbCreateMfg ==>   Boolean flag indicating if software should automatically create MfgProfiles for the new Profiles.
' Output of this routine will be one of the Profile in pReplacingObjColl collection. "M1" in the above example will then be reattached to that Profile and thus maintain the lot# for this Profile part.
' NOTE:  The rule is not limited to two Profiles.  "n" number of Profiles could be passed to the rule after a split.  The rule would still only return one Profile meeting the criteria.
'
' pbCreateMfg   Output of MigrateObject()   Action
' ===========   =========================   ======
' False         Nothing           MfgProfile ( M1 ) will be deleted. NO new MfgProfiles will be created.
' True          Nothing           MfgProfile ( M1 ) will be deleted. New MfgProfiles will be created for all the new Profiles ( for P2 and P3 ).
' False         Object P2         MfgProfile ( M1 ) will be re-connected from P1 to P2. New MfgProfiles will NOT be created for the remaining new Profiles ( for P3 ).
' True          Object P2         MfgProfile ( M1 ) will be re-connected from P1 to P2.  New MfgProfiles will be created for the remaining new Profiles ( for P3 ).
'
    Const METHOD = "IJMfgSplitMigration_MigrateObject"
    On Error GoTo ErrorHandler

    Dim lNewProfileCount As Long
    Dim indProfile As Long
    Dim dMaxLength As Double
    Dim dMaxLengthIndex As Long

    'Create MfgProfiles automatically for the new Profiles.
    pbCreateMfg = True

    'Get the new Profile count
    lNewProfileCount = pReplacingObjColl.Count

    ' If the new Profile count is zero, there is nothing to do.
    If lNewProfileCount = 0 Then
        Exit Function
    End If

    ' Initialization.
    dMaxLength = -1# 'maximum landing curve length
    dMaxLengthIndex = 0 'index of the biggest Profile part in pReplacingObjColl

    
    ' Loop for each element in the new Profile collection.
    For indProfile = 1 To lNewProfileCount
        'Create ProfilePart support
        Dim oProfilePartSupport As IJProfilePartSupport
        Set oProfilePartSupport = New ProfilePartSupport
        
        Dim oPartSupport As IJPartSupport
        Set oPartSupport = oProfilePartSupport
        
        Dim dLength As Double
        'Get the base surface of this plate part
        Set oPartSupport.Part = pReplacingObjColl.Item(indProfile)

        dLength = oProfilePartSupport.ApproximateLength
        
        ' See if this length is larger than the previous one
        If dLength > dMaxLength Then
             ' If so, store this length as maximum and the element index
            dMaxLength = dLength
            dMaxLengthIndex = indProfile
        End If

        dLength = -1#
    Next indProfile

    'Something went wrong. Quit here
    If dMaxLengthIndex = 0 Then
        Exit Function
    End If

    'Return the Profile part with largest landing curve length.
    Set IJMfgSplitMigration_MigrateObject = pReplacingObjColl.Item(dMaxLengthIndex)
    
    ' clone the process and marking settings
    pOptionalArg = True

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2027, , "RULES")
End Function


Private Function IJMfgSplitMigration_ReverseMigrate(ByVal pReplacedPartColl As GSCADStructMfgGlobals.IJElements, ByVal pReplacingPart As Object, ByVal pMfgObjsColl As GSCADStructMfgGlobals.IJElements, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
    Const METHOD = "IJMfgSplitMigration_ReverseMigrate"
    On Error GoTo ErrorHandler
    
    'Return the plate part with largest surface area.
        
    pbCreateMfg = False ' set this to true only if migration is not needed but want to create a new mfg profile
    pOptionalArg = False ' this makes send to clone the settings only when the above the flag is true
    
    ' Initialization.
    Dim lNewProfileCount    As Long
    Dim indProfile          As Long
    Dim dMaxLength          As Double
    Dim dMaxLengthIndex     As Long

    dMaxLength = -1# 'maximum landing curve length
    dMaxLengthIndex = 0 'index of the biggest Profile part in pReplacingObjColl

    Dim oMfgUpdateInfo As IJMfgUpdateInfo
    
    ' Loop for each element in the manufactruing objects collection.
    For indProfile = 1 To pMfgObjsColl.Count
        Dim oProfilePartSupport As IJProfilePartSupport
        Dim oPartSupport        As IJPartSupport
    
        'Create ProfilePart support
        Set oProfilePartSupport = New ProfilePartSupport
        Set oPartSupport = oProfilePartSupport
                
        Dim dLength As Double
        Dim oMfgProfileAsMfgChild As IJMfgChild
        Set oMfgProfileAsMfgChild = pMfgObjsColl.Item(indProfile)
        'Get the detailed profile part part for this manufacturing profile
        Set oPartSupport.Part = oMfgProfileAsMfgChild.GetParent
        
        Set oMfgUpdateInfo = oMfgProfileAsMfgChild
        oMfgUpdateInfo.UpToDate = FLAG_PART_DELETED
        
        dLength = oProfilePartSupport.ApproximateLength

        ' See if this length is larger than the previous one
        If dLength > dMaxLength Then
             ' If so, store this length as maximum and the element index
            dMaxLength = dLength
            dMaxLengthIndex = indProfile
        End If

        dLength = -1#

    Next indProfile

    Set oProfilePartSupport = Nothing
    Set oPartSupport = Nothing
    
    Set oMfgUpdateInfo = pMfgObjsColl.Item(dMaxLengthIndex)
    oMfgUpdateInfo.UpToDate = FLAG_MIGRATED
    
    'Return the Profile part with largest landing curve length.
    Set IJMfgSplitMigration_ReverseMigrate = pMfgObjsColl.Item(dMaxLengthIndex)

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2027, , "RULES")
End Function
