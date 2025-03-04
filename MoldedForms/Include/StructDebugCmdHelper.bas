Attribute VB_Name = "StructDebugCmdHelper"
'------------------------------------------------------------------------------------------------------------
' Copyright (C) 2002  Intergraph Corporation. All rights reserved.
'
' Project:
'     n/a
'
' File:
'     StructDebugCmdHelper.bas
'
' Description:
'      This common include module simply wraps the StructDebugCommand methods for
'      conditional compilation.
'
' Notes:
'     <notes>
'
' History:
'     S. B. Talbert     04/25/2002     Initial creation.
'------------------------------------------------------------------------------------------------------------

Option Explicit

'============================================================
' MODULE-LEVEL DECLARATIONS
'============================================================

#If DEBUG_COMPILE Then
Private SDC As New StructDebugCommand
#End If

'============================================================
' StructDebugCommand Wrapper Procedures
'============================================================

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     DEBUG_DEEP_BEGIN
'
' Description:
'     Increase deep message output.
'
' Arguments:
'     (none)
'------------------------------------------------------------------------------------------------------------
Public Sub DEBUG_DEEP_BEGIN()
    #If DEBUG_COMPILE Then
        SDC.DEBUG_DEEP_BEGIN
    #End If
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     DEBUG_DEEP_END
'
' Description:
'     Decrease deep message output.
'
' Arguments:
'     (none)
'------------------------------------------------------------------------------------------------------------
Public Sub DEBUG_DEEP_END()
    #If DEBUG_COMPILE Then
        SDC.DEBUG_DEEP_END
    #End If
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     DEBUG_DUMP_LIST
'
' Description:
'     Dump the given elements list.  The message will be "'ElementListInfo' count: <count>"
'     followed by the dump of each element in the list.
'
' Arguments:
'     ElementList            Reference to the element list being dumped
'     ElementListInfo      Element list description
'------------------------------------------------------------------------------------------------------------
Public Sub DEBUG_DUMP_LIST(ElementList As Variant, ElementListInfo As String)
    #If DEBUG_COMPILE Then
        SDC.DEBUG_DUMP_LIST ElementList, ElementListInfo
    #End If
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     DEBUG_DUMP_MATRIX
'
' Description:
'     Dump the given 4x4 matrix. The message will be "'MatrixInfo' matrix :" followed by a
'     dump of each value of the matrix.
'
' Arguments:
'     Matrix            Reference to the matrix being dumped
'     MatrixInfo      Matrix description
'------------------------------------------------------------------------------------------------------------
Public Sub DEBUG_DUMP_MATRIX(Matrix As IJDT4x4, MatrixInfo As String)
    #If DEBUG_COMPILE Then
        SDC.DEBUG_DUMP_MATRIX Matrix, MatrixInfo
    #End If
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     DEBUG_DUMP_OBJECT
'
' Description:
'     Dump the specified element. The message will be "'ElementInfo': <dump of element>".
'
' Arguments:
'     Element            Reference to the element being dumped
'     ElementInfo      Element description
'------------------------------------------------------------------------------------------------------------
Public Sub DEBUG_DUMP_OBJECT(Element As Object, ElementInfo As String)
    #If DEBUG_COMPILE Then
        SDC.DEBUG_DUMP_OBJECT Element, ElementInfo
    #End If
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     DEBUG_DUMP_POSITION
'
' Description:
'     Dump the X, Y, and Z values of either IJDPosition, IJDVector, or IJDTriple.  The message
'     will be "'PositionInfo' : (X,Y,Z)".
'
' Arguments:
'     Position            Reference to the position object being dumped
'     PositionInfo      Position object description
'------------------------------------------------------------------------------------------------------------
Public Sub DEBUG_DUMP_POSITION(Position As Object, PositionInfo As String)
    #If DEBUG_COMPILE Then
        SDC.DEBUG_DUMP_POSITION Position, PositionInfo
    #End If
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     DEBUG_ERROR_MSG
'
' Description:
'     Display message output in DebugLog window and in message box.
'
' Arguments:
'     Message      The message to be displayed
'------------------------------------------------------------------------------------------------------------
Public Sub DEBUG_ERROR_MSG(Message As String)
    #If DEBUG_COMPILE Then
        SDC.DEBUG_ERROR_MSG Message
    #End If
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     DEBUG_MSG
'
' Description:
'     Display message output.
'
' Arguments:
'     Message      The message to be displayed
'------------------------------------------------------------------------------------------------------------
Public Sub DEBUG_MSG(Message As String)
    #If DEBUG_COMPILE Then
        SDC.DEBUG_MSG Message
    #End If
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     DEBUG_MSG_ONCE
'
' Description:
'     Display the given message once only, until 'cookie' number is different or
'     DEBUG_MSG_ONCE_RESET is called. This method is useful when using it in a loop.
'
' Arguments:
'     Cookie        Basically, the message key
'     Message      The message to be displayed
'------------------------------------------------------------------------------------------------------------
Public Sub DEBUG_MSG_ONCE(Cookie As Integer, Message As String)
    #If DEBUG_COMPILE Then
        SDC.DEBUG_MSG_ONCE Cookie, Message
    #End If
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     DEBUG_MSG_ONCE_RESET
'
' Description:
'     Reset the cookie number used in procedure DEBUG_MSG_ONCE.
'
' Arguments:
'     (none)
'------------------------------------------------------------------------------------------------------------
Public Sub DEBUG_MSG_ONCE_RESET()
    #If DEBUG_COMPILE Then
        SDC.DEBUG_MSG_ONCE_RESET
    #End If
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     DEBUG_REFCOUNT_OBJECT
'
' Description:
'     Display the Reference Count value of the specified element.
'
' Arguments:
'     Element      The element
'------------------------------------------------------------------------------------------------------------
Public Sub DEBUG_REFCOUNT_OBJECT(Element As Object)
    #If DEBUG_COMPILE Then
        SDC.DEBUG_REFCOUNT_OBJECT Element
    #End If
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure (Function):
'     DEBUG_REGISTRY_VALUE
'
' Description:
'     Return value of registry key
'     \HKEY_CURRENT_USER\Software\Intergraph\DEBUG\StructDebugCommand_DEBUG.
'
' Arguments:
'     (none)
'------------------------------------------------------------------------------------------------------------
Public Function DEBUG_REGISTRY_VALUE() As Long
    #If DEBUG_COMPILE Then
        DEBUG_REGISTRY_VALUE = SDC.DEBUG_REGISTRY_VALUE
    #End If
End Function

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     DEBUG_SOURCE
'
' Description:
'     Set source file description.
'
' Arguments:
'     SourceInfo      Source file description
'------------------------------------------------------------------------------------------------------------
Public Sub DEBUG_SOURCE(SourceInfo As String)
    #If DEBUG_COMPILE Then
        SDC.DEBUG_SOURCE = SourceInfo
    #End If
End Sub
