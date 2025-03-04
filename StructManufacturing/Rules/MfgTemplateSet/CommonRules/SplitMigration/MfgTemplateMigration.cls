VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateMigration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''
''   Copyright (C) 2002, Intergraph Corporation. All rights reserved.
''
''   Abstract:
''              MfgTemplateMigration.cls
''   ProgID:
''              StrMfgMigrationRules.TemplateMigration
''   Author:
''              Suresh Kadambari
''   Creation Date:
''              12 Nov 2006
''   Description:
''      Struct Mfg should check split notification and keep the template set.
''
''      In a simplest test, a part "P1" is split into two smaller parts "P2" and "P3" and "P1" will be deleted.
''      Old part P1, along with new parts P2 and P3 will be passed to a user defined VB rule.  This rule will be
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
Private Const MODULE = "TemplateMigration"


Private Function IJMfgSplitMigration_MigrateObject(ByVal pOldObj As Object, ByVal pReplacingObjColl As GSCADStructMfgGlobals.IJElements, ByVal pMfgObject As Object, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
'Arguments:
' pOldObj ==>  Plate "P1" which is being deleted.
'  pReplacingObjColl ==> Collection of replacing new plates ( P2 and P3 ).
' pbCreateMfg ==>   Boolean flag indicating if software should automatically create template sets for the new plates.
' Output of this routine will be one of the plate in pReplacingObjColl collection.
' NOTE:  The rule is not limited to two plates.  "n" number of plates could be passed to the rule after a split.  The rule would still only return one plate meeting the criteria.
'
' pbCreateMfg   Output of MigrateObject()   Action
' ===========   =========================   ======
' False         Nothing           TemplateSet ( T1 ) will be deleted. NO new TemplateSets will be created.
' True          Nothing           TemplateSet ( T1 ) will be deleted. New TemplateSets will be created for all the new plates ( for P2 and P3 ).
' False         Object P2         TemplateSet ( T1 ) will be re-connected from P1 to P2. New TemplateSets will NOT be created for the remaining new plates ( for P3 ).
' True          Object P2         TemplateSet ( T1 ) will be re-connected from P1 to P2.  New TemplateSets will be created for the remaining new plates ( for P3 ).
    Const METHOD = "IJMfgSplitMigration_MigrateObject"
    On Error GoTo ErrorHandler

    Dim lNewPlateCount As Long
    Dim indPlate As Long
    Dim dMaxSurfaceArea As Double
    Dim lMaxSurfaceIndex As Long

    'Don't create TemplateSet automatically for the new plates.
    pbCreateMfg = False

    'Get the new Plate count
    lNewPlateCount = pReplacingObjColl.Count

    ' If the new plate count is zero, there is nothing to do.
    If lNewPlateCount = 0 Then
        Exit Function
    End If
    
    Dim oPlatePartSupport   As IJPlatePartSupport
    Dim oPartSupport        As IJPartSupport
    Dim oProfileWrapper     As MfgRuleHelpers.ProfilePartHlpr
    
    ' Initialization.
    dMaxSurfaceArea = -1# 'maximum surface area
    lMaxSurfaceIndex = 0 'index of the biggest plate part in pReplacingObjColl

    Dim bPlatePart  As Boolean
    If TypeOf pReplacingObjColl.Item(1) Is IJPlatePart Then
        bPlatePart = True
                
        'Create PlatePart support to get the base surface of the plate without any features/camfers
        Set oPlatePartSupport = New PlatePartSupport
        Set oPartSupport = oPlatePartSupport
    Else
        'Create ProfileWrapper to get the web left surface of the profile
        Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    End If
    
    ' Loop for each element in the new plate collection.
    For indPlate = 1 To lNewPlateCount
    
        Dim oPartSurface As Object
        
        If bPlatePart = True Then
            'Get the base surface of this plate part
            Set oPartSupport.Part = pReplacingObjColl.Item(indPlate)
            oPlatePartSupport.GetSurface PlateBaseSide, oPartSurface
        Else
            'Get the web left surface of this profile part
            Set oProfileWrapper.object = pReplacingObjColl.Item(indPlate)
            
            Dim oSurfacePort As IJPort
            Set oSurfacePort = oProfileWrapper.GetSurfacePort(JXSEC_WEB_LEFT)
            Set oPartSurface = oSurfacePort.Geometry
        End If
    
        Dim oBaseSurfaceModelBody As IJDModelBody
        Set oBaseSurfaceModelBody = oPartSurface

        'Calculate the surface area
        Dim dDummyLength As Double
        Dim dEstRelAccyAchieved As Double
        Dim dSurfaceArea As Double
        Dim dDummyVolume As Double

        oBaseSurfaceModelBody.GetDimMetrics 0.001, dEstRelAccyAchieved, dDummyLength, dSurfaceArea, dDummyVolume

        ' See if this area is larger than the previous one
        If dSurfaceArea > dMaxSurfaceArea Then
             ' If so, store this area as maximum and the element index
             dMaxSurfaceArea = dSurfaceArea
             lMaxSurfaceIndex = indPlate
        End If
        
        Set oBaseSurfaceModelBody = Nothing
        Set oPartSurface = Nothing
        
    Next indPlate

    Set oPlatePartSupport = Nothing
    Set oPartSupport = Nothing
    Set oProfileWrapper = Nothing
    
    'Something went wrong. Quit here
    If lMaxSurfaceIndex = 0 Then
        Exit Function
    End If

    'Return the plate part with largest surface area.
    Set IJMfgSplitMigration_MigrateObject = pReplacingObjColl.Item(lMaxSurfaceIndex)
    pOptionalArg = True

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3011, , "RULES")
End Function

Private Function IJMfgSplitMigration_ReverseMigrate(ByVal pReplacedPartColl As GSCADStructMfgGlobals.IJElements, ByVal pReplacingPart As Object, ByVal pMfgObjsColl As GSCADStructMfgGlobals.IJElements, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
    Const METHOD = "IJMfgSplitMigration_ReverseMigrate"
    On Error GoTo ErrorHandler

    'Return the template set on the plate part with largest surface area.
        
    pbCreateMfg = False ' set this to true only if migration is not needed but want to create a new template set
    pOptionalArg = False ' this makes send to clone the settings only when the above the flag is true
    
    ' Initialization.
    Dim indTemplateSet      As Long
    Dim dMaxSurfaceArea     As Double
    Dim lMaxSurfaceIndex    As Long
    
    dMaxSurfaceArea = -1# 'maximum surface area
    lMaxSurfaceIndex = 0 'index of the biggest plate part in pReplacingObjColl

    Dim oPartSupport        As IJPartSupport
    Dim oPlatePartSupport   As IJPlatePartSupport
    Dim oProfileWrapper     As MfgRuleHelpers.ProfilePartHlpr

    ' Check if type of parent is plate part/profile part
    Dim oMfgChild As IJMfgChild
    Set oMfgChild = pMfgObjsColl.Item(1)
    
    Dim bPlatePart  As Boolean
    If TypeOf oMfgChild.GetParent Is IJPlatePart Then
        bPlatePart = True
        
        'Create PlatePart support to get the base surface of the plate without any features/camfers
        Set oPlatePartSupport = New PlatePartSupport
        Set oPartSupport = oPlatePartSupport
    Else
        'Create ProfilePart wrapper to get the web left surface of the Profile part
        Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    End If
    
    Dim oMfgUpdateInfo As IJMfgUpdateInfo
    
    ' Loop for each element in the manufactruing objects collection.
    For indTemplateSet = 1 To pMfgObjsColl.Count
        
        Dim oTemplateSet As IJMfgChild
        Set oTemplateSet = pMfgObjsColl.Item(indTemplateSet)
        
        Set oMfgUpdateInfo = oTemplateSet
        oMfgUpdateInfo.UpToDate = FLAG_PART_DELETED
        
        Dim oPartSurface As Object
        
        If bPlatePart = True Then
        
            'Get the base surface of this plate part
            Set oPartSupport.Part = oTemplateSet.GetParent
            oPlatePartSupport.GetSurface PlateBaseSide, oPartSurface
        Else
            
            'Get the web left surface of this profile part
            Set oProfileWrapper.object = oTemplateSet.GetParent
            
            Dim oSurfacePort As IJPort
            Set oSurfacePort = oProfileWrapper.GetSurfacePort(JXSEC_WEB_LEFT)
            Set oPartSurface = oSurfacePort.Geometry
        End If
        
        Dim oBaseSurfaceModelBody As IJDModelBody
        Set oBaseSurfaceModelBody = oPartSurface

        'Calculate the surface area
        Dim dDummyLength            As Double
        Dim dEstRelAccyAchieved     As Double
        Dim dSurfaceArea            As Double
        Dim dDummyVolume            As Double

        oBaseSurfaceModelBody.GetDimMetrics 0.001, dEstRelAccyAchieved, dDummyLength, dSurfaceArea, dDummyVolume

        ' See if this area is larger than the previous one
        If dSurfaceArea > dMaxSurfaceArea Then
             ' If so, store this area as maximum and the element index
             dMaxSurfaceArea = dSurfaceArea
             lMaxSurfaceIndex = indTemplateSet
        End If
        
        Set oBaseSurfaceModelBody = Nothing
        Set oPartSurface = Nothing
        
    Next indTemplateSet
    
    Set oPlatePartSupport = Nothing
    Set oPartSupport = Nothing
    Set oProfileWrapper = Nothing
    
    Set oMfgUpdateInfo = pMfgObjsColl.Item(lMaxSurfaceIndex)
    oMfgUpdateInfo.UpToDate = FLAG_MIGRATED
    
    'Return the manufacturing part with largest surface area.
    Set IJMfgSplitMigration_ReverseMigrate = pMfgObjsColl.Item(lMaxSurfaceIndex)
   
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3011, , "RULES")

End Function
