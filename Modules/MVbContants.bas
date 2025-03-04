Attribute VB_Name = "MVbConstants"
'************************************************************************
'  Copyright (C) 2001, Intergraph Corporation.  All rights reserved.
'
'  Project: Stand-alone include
'  File:    MVbConstants.bas
'
'  Author:  Steven Mitchell
'
'  Desc:    Define common VB constants not provided by VB itself.
'
'  History:
'  - Initial Development
'       11/29/2001    Steven Mitchell
'       06/03/2002    Jehle - added constant for DBL_UNDEFINED and FLT_UNDEFINED
'************************************************************************

Option Explicit

' VB Trappable errors -- see MSDN, Visual Tools and Languages,
' Visual Studio Documentation, Visual Basic Documentation,
' Reference, Trappable Errors
Public Enum E_VB_ERRORS
    E_VB_RETURN_WITHOUT_GOSUB = &H3             ' 3
    E_VB_INVALID_PROC_CALL = &H5                ' 5
    E_VB_OVERFLOW = &H6                         ' 6
    E_VB_OUT_OF_MEMORY = &H7                    ' 7
    E_VB_SUBSCRIPT_OUT_OF_RANGE = &H9           ' 9
    E_VB_ARRAY_FIXED_OR_LOCKED = &HA            ' 10
    E_VB_DIV_BY_ZERO = &HB                      ' 11
    E_VB_TYPE_MISMATCH = &HD                    ' 13
    E_VB_OUT_OF_STRING_SPACE = &HE              ' 14
    E_VB_EXPRESSION_TOO_COMPLEX = &H10          ' 16
    E_VB_CANNOT_PERFORM_OPERATION = &H11        ' 17
    E_VB_USER_INTERRUPT = &H12                  ' 18
    E_VB_RESUME_WITHOUT_ERROR = &H14            ' 20
    E_VB_OUT_OF_STACK_SPACE = &H1C              ' 28
    E_VB_ROUTINE_NOT_DEFINED = &H23             ' 35
    E_VB_TOO_MANY_DLL_APP_CLIENTS = &H2F        ' 47
    E_VB_ERROR_LOADING_DLL = &H30               ' 48
    E_VB_BAD_DLL_CALL_CONVENTION = &H31         ' 49
    E_VB_INTERNAL_ERROR = &H33                  ' 51
    E_VB_BAD_FILE_NAME_OR_NUMBER = &H34         ' 52
    E_VB_CANNOT_FIND_SPECIFIED_FILE = &H35      ' 53
    E_VB_BAD_FILE_MODE = &H36                   ' 54
    E_VB_FILE_ALREADY_OPEN = &H37               ' 55
    E_VB_DEVICE_IO_ERROR = &H39                 ' 57
    E_VB_FILE_ALREADY_EXISTS = &H3A             ' 58
    E_VB_BAD_RECORD_LENGTH = &H3B               ' 59
    E_VB_DISK_FULL = &H3D                       ' 61
    E_VB_INPUT_PAST_EOF = &H3E                  ' 62
    E_VB_BAD_RECORD_NUMBER = &H3F               ' 63
    E_VB_TOO_MANY_FILES = &H43                  ' 67
    E_VB_DEVICE_UNAVAILABLE = &H44              ' 68
    E_VB_PERMISSION_DENIED = &H46               ' 70
    E_VB_DISK_NOT_READY = &H47                  ' 71
    E_VB_CANNOT_RENAME_DIFF_DRIVE = &H4A        ' 74
    E_VB_PATH_FILE_ACCESS_ERROR = &H4B          ' 75
    E_VB_PATH_NOT_FOUND = &H4C                  ' 76
    E_VB_OBJECT_VARIABLE_NOT_SET = &H5B         ' 91
    E_VB_FOR_LOOP_NOT_INIT = &H5C               ' 92
    E_VB_INVALID_PATTERN_STRING = &H5D          ' 93
    E_VB_INVALID_NULL = &H5E                    ' 94
    E_VB_SINK_FAIL_DUE_TO_MAX_EVENTS = &H60     ' 96
    E_VB_CANNOT_CALL_FRIEND_PROC = &H61         ' 97
    E_VB_PRIVATE_ROUTINE_PUBLIC_REF = &H62      ' 98
    E_VB_INVALID_FILE_FORMAT = &H141            ' 321
    E_VB_CANNOT_CREATE_TEMP_FILE = &H142        ' 322
    E_VB_SHOW_DISPLAYED_FORM_MODAL = &H190      ' 400
    E_VB_PROPERTY_NOT_FOUND = &H1A6             ' 422
    E_VB_AUTO_CREATE_OBJECT = &H1AD             ' 429
    E_VB_AUTO_NOT_SUPPORTED = &H1AE             ' 430
    E_VB_AUTO_FILE_OR_CLASS_NOT_FOUND = &H1B0   ' 432
    E_VB_AUTO_ROUTINE_NOT_SUPPORTED = &H1B6     ' 438
    E_VB_AUTO_ERROR = &H1B8                     ' 440
    E_VB_AUTO_REMOTE_CONNECTION_LOST = &H1BA    ' 442
    E_VB_AUTO_NO_DEFAULT_VALUE = &H1BB          ' 443
    E_VB_AUTO_ACTION_NOT_SUPPORTED = &H1BD      ' 445
    E_VB_AUTO_NAMED_ARGS_NOT_SUPPORTED = &H1BE  ' 446
    E_VB_AUTO_CURR_LCID_NOT_SUPPORTED = &H1BF   ' 447
    E_VB_AUTO_NAMED_ARG_NOT_FOUND = &H1C0       ' 448
    E_VB_AUTO_ARG_NOT_OPTIONAL = &H1C1          ' 449
    E_VB_AUTO_WRONG_NUMBER_ARGS = &H1C2         ' 450
    E_VB_PROPERTY_LET_GET_INCORRECT = &H1C3     ' 451
    E_VB_INVALID_ORDINAL = &H1C4                ' 452
    E_VB_DLL_FUNCTION_NOT_FOUND = &H1C5         ' 453
    E_VB_CODE_RESOURCE_NOT_FOUND = &H1C6        ' 454
    E_VB_CODE_RESOURCE_LOCK_ERROR = &H1C7       ' 455
    E_VB_COL_DUPLICATE = &H1C9                  ' 457
    E_VB_AUTO_UNSUPPORTED_TYPE = &H1CA          ' 458
    E_VB_EVENTS_NOT_SUPPORTED = &H1CB           ' 459
    E_VB_CANNOT_SAVE_FILE_TO_TEMP_DIR = &H2DF   ' 735
    E_VB_SEARCH_TEXT_NOT_FOUND = &H2E8          ' 744
    E_VB_REPLACEMENTS_TOO_LONG = &H2EA          ' 746
    E_VB_SYS_OUT_OF_MEMORY = &H7919             ' 31001
End Enum

' Define common windows errors in VB terms -- see
' \<DevStudio\VC98\Include\WinErrors.h
Public Const S_OK As Long = &H0

Public Enum E_WIN_ERRORS
    E_NOTIMPL = &H80004001                      ' 2147500033
    E_FAIL = &H80004005                         ' 2147500037
    E_UNEXPECTED = &H8000FFFF                   ' 2147549183
    E_OLE_REG_FAIL = &H8002801C                 ' 2147647516
    E_ACCESS_DENIED = &H80070005                ' 2147942405
    E_OUTOFMEMORY = &H8007000E                  ' 2147942414
    E_INVALIDARG = &H80070057                   ' 2147942487
    E_CANT_BINDTOSOURCE = &H8004000A
End Enum

' Define constants to undefine double and float variables
Public Const DBL_UNDEFINED As Double = 1.6E+308 ' arbitrary number near the maximum double value that can be stored
Public Const FLT_UNDEFINED = 3.3E+38            ' arbitrary number near the maximum float value that can be stored


