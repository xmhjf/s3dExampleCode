Attribute VB_Name = "PropertyPageHelper2"
'********************************************************************
' Copyright (C) 2002 Intergraph Corporation.  All Rights Reserved.
'
' File: PropertyPageHelper2.cls
'
' Abstract: Helper subs / functions for property pages to use.
'
' Description:
'   This provides functions that can be included in property page projects to
'   ease the implementation.
'   Note:  This extends the original PropertyPageHelper.bas module by including support
'          for pages that implement IJPpHelper.
'
'History
'   Chris Gibble         01/24/02     Creation.
'********************************************************************

Option Explicit

Option Private Module    ' Makes public functions private outside of project
 
' Note:  All Booleans default to False, object variables default to Nothing.
Private m_bHasBeenInitialized As Boolean            ' Flag indicating if the data has been initialized
Private m_bInEditPropertiesMode As Boolean          ' Flag indicating this is in Edit Properties mode
Private m_bDoApplyChanges As Boolean                ' Flag indicating this should apply changes
Private m_oHolderContainer As IJHolderContainer     ' Points to specified HolderContainer.
 

'********************************************************************
' Routine: InitializeHelper
'
' Abstract: Sets up module variables.
'
' Description:  Asks Trader for the strExistingHolder -- if it exists,
' then this must be in Create mode (since the command has created it).
' Otherwise, it's in Edit Properties mode.
'********************************************************************
Public Sub InitializeHelper(strExistingHolder As String, oHolderContainerFactory As IJHolderContainerFactory)
    Dim oTrader As Trader
    Dim oPropertyHolder As IJPropertyHolder
    
    If Not m_bHasBeenInitialized Then    ' Continue only if this has not already been initialized
        m_bHasBeenInitialized = True
 
        'Check mode by looking for the presence of a property holder in the Trader
        m_bInEditPropertiesMode = False

        Set oTrader = New Trader
        Set oPropertyHolder = oTrader.Service(strExistingHolder, "")   'Look for an existing holder (default is Parent System)
        If oPropertyHolder Is Nothing Then
            m_bInEditPropertiesMode = True
            Set m_oHolderContainer = oHolderContainerFactory.CreateFromSelectSet
        End If
        
        Set oTrader = Nothing
        Set oPropertyHolder = Nothing
    End If
End Sub
 
'********************************************************************
' Routine: TerminateHelper
'
' Abstract: Cleans up module variables.
'
' Description:  Cleans up module variables.
'********************************************************************
Public Sub TerminateHelper()
    ' Reset variables
    Set m_oHolderContainer = Nothing
    m_bHasBeenInitialized = False
    m_bInEditPropertiesMode = False
End Sub
 
'********************************************************************
' Routine: HelperCheckpointValues
'
' Abstract: Checkpoints values on the passed-in Holders.
'
' Description:  Checkpoints values on the passed-in Holders.
'********************************************************************
Public Sub HelperCheckpointValues(oHolderCol As Collection)
    Dim oPropertyHolder As IJPropertyHolder
    
    For Each oPropertyHolder In oHolderCol
        oPropertyHolder.CheckpointCurrentValues
    Next
    
    Set oPropertyHolder = Nothing
End Sub

'********************************************************************
' Routine: HelperDisableControls
'
' Abstract: Disables the passed-in controls.
'
' Description:  Loops through the passed-in controls and disables them
'               (excluding labels and frames).
'********************************************************************
Public Sub HelperDisableControls(oControls As Object)
    Dim oControl As Control
    
    On Error Resume Next    ' In case a custom control does not have an "Enabled" property
    
    For Each oControl In oControls
        Select Case UCase$(TypeName(oControl))
            Case "LABEL", "FRAME"
                ' Leave these enabled
            Case Else
                oControl.Enabled = False
        End Select
    Next
    
    Set oControl = Nothing
End Sub

'********************************************************************
' Routine: HelperOKToCommit
'
' Abstract: Populates return flag with Boolean indicating whether Commit
'           can be called or not.
'
' Description:  Populates return flag with Boolean indicating whether Commit
'           can be called or not.  (Currently, this does nothing.)
'********************************************************************
Public Sub HelperOKToCommit(bCommitValue As Boolean, strErrorMsg As String, oHolderCol As Collection)
    ' Current logic is to always return True
    bCommitValue = True
    
    ' Set flag to actually apply changes
    m_bDoApplyChanges = True
End Sub


'********************************************************************
' Routine: HelperRefreshFromCheckpointValues
'
' Abstract: Refreshes from checkpointed values on the passed-in Holders.
'
' Description:  Refreshes from checkpointed values on the passed-in Holders.
'********************************************************************
Public Sub HelperRefreshFromCheckpointValues(oHolderCol As Collection)
    Dim oPropertyHolder As IJPropertyHolder
    
    ' Checkpoint all holders
    For Each oPropertyHolder In oHolderCol
        oPropertyHolder.RefreshFromCheckpointValues
    Next
    
    Set oPropertyHolder = Nothing
End Sub

'********************************************************************
' Routine: HelperPreCommit
'
' Abstract: Applies changes to the selected business objects in Edit properties mode.
'
' Description:  Applies changes to the selected business objects in Edit properties mode.
'********************************************************************
Public Sub HelperPreCommit(oHolderCol As Collection)
    ' In Edit Properties mode, apply changes if user has made any
    '(We should only get here in Modify mode.)
    If m_bInEditPropertiesMode Then
        HelperCheckpointValues oHolderCol
        
        ' Only want to apply changes if values were actually changed.
        
        ' Note: if m_bDoApplyChanges = True and we entered this sub
        ' as a result of a Cancel click, the subsequent Abort will
        ' "wipe out" any changes applied to the BO, so things still work.
        If m_bDoApplyChanges Then
            m_oHolderContainer.ApplyChangesToBO
            m_bDoApplyChanges = False   ' only need to do this stuff once for ALL property holders
        End If
    End If
End Sub


'********************************************************************
' Routine: HelperPostCommit
'
' Abstract: Checkpoints the current values in create mode.
'
' Description:  Checkpoints the current values in create mode.
'       Note:  this method is called when the user clicks on Apply or OK in create mode.
'       When the PropertyPage terminates, it calls RefreshFromCheckpointValues -- this
'       makes sure the selected values are kept.
'********************************************************************
Public Sub HelperPostCommit(oHolderCol As Collection)
    'Checkpoint values.
    '(We should only get here after an Apply or OK in Create mode.)
    If Not m_bInEditPropertiesMode Then
        HelperCheckpointValues oHolderCol
    End If
End Sub
