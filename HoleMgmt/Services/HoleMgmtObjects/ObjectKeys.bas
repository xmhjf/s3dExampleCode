Attribute VB_Name = "ObjectKeys"
Public Enum tagAttributeTypes
    SQL_VB_CHAR = 1  ' CHAR, VARCHAR, DECIMAL, NUMERIC = VT_BSTR = SQL_C_CHAR = SQL_CHAR
    SQL_VB_LONG = 4  ' long int = VT_I4 = SQL_C_LONG = SQL_INTEGER
    SQL_VB_SHORT = 5 ' shrt int = VT_I2 = SQL_C_SHORT = SQL_SMALLINT
    SQL_VB_FLOAT = 7 ' float = VT_R4 = SQL_C_FLOAT = SQL_REAL
    SQL_VB_DOUBLE = 8  ' double = VT_R8 = SQL_C_DOUBLE = SQL_DOUBLE
    SQL_VB_BIT = -7  ' boolean = VT_BOOL = SQL_C_BIT
    SQL_VB_DATE = 9  ' date = VT_DATE = SQL_C_DATE
End Enum

 