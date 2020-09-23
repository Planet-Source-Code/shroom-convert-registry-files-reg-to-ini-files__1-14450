Attribute VB_Name = "INIReadWrite"
Option Explicit
'Thanks to allapi.net for their wonderful API Guide

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'm_file is the file we operate under
'm_buffer is the number of characters to retrieve max
'   -- Need to set this high (over 5000) if you plan on
'      Read_Sections or Read_Keys a large INI
Dim m_File As String, m_Buffer As Long

Public Sub INISetup(FileName As String, BufferSize As Long)
    m_Buffer = BufferSize
    m_File = FileName
End Sub

Public Function Read_Ini(iSection As String, iKeyName As String, Optional iDefault As String)
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    Dim Ret As String, NC As Long
    
    'Create the buffer
    Ret = String(m_Buffer, 0)
    
    'Retrieve the string
    NC = GetPrivateProfileString(iSection, iKeyName, iDefault, Ret, m_Buffer, m_File)
    
    'NC is the number of characters copied to the buffer
    If NC <> 0 Then
        Ret = Left$(Ret, NC)
    Else
        'Make sure to cut it down to number of char's returned
        Ret = ""
    End If
    
    'Turn the funky vbcrlf string into VBCRLFs
    Ret = Replace(Ret, "%%&&Chr(13)&&%%", vbCrLf)
    
    'Return the setting
    Read_Ini = Ret
End Function

Public Sub Write_Ini(iSection As String, iKeyName As String, iValue As Variant)
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    
    'Make sure to change it to a String
    iValue = CStr(iValue)
    
    'Turn all vbcrlf's into that funky string
    iValue = Replace(iValue, vbCrLf, "%%&&Chr(13)&&%%")
    WritePrivateProfileString iSection, iKeyName, CStr(iValue), m_File
End Sub

Public Function Read_Sections()
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    Dim Ret As String, NC As Long
    
    'Create the buffer
    Ret = String(m_Buffer, 0)
    
    'Retrieve the string, return '[-na-]' if there is none
    NC = GetPrivateProfileString(vbNullString, vbNullString, vbNullString, Ret, m_Buffer, m_File)
    
    'NC is the number of characters returned
    If NC <> 0 Then
        Ret = Left$(Ret, NC - 1)
    Else
        Ret = ""
    End If
    
    'Return the sections
    Read_Sections = Ret
End Function

Public Function Read_Keys(iSection As String)
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    Dim Ret As String, NC As Long
    
    'Create the buffer
    Ret = String(m_Buffer, 0)
    
    'Retrieve the string, return '[-na-]' if there is none
    NC = GetPrivateProfileString(iSection, vbNullString, vbNullString, Ret, m_Buffer, m_File)
    
    'NC is the number of characters copied to the buffer
    If NC <> 0 Then
        Ret = Left$(Ret, NC - 1)
    Else
        Ret = ""
    End If
    'Return the sections
    Read_Keys = Ret
End Function

Public Function DeleteSection(iSection As String)
'Haven't tested these two myself =\
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    WritePrivateProfileString iSection, vbNullString, vbNullString, m_File
End Function

Function DeleteKey(iSection As String, iKeyName As String)
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    WritePrivateProfileString iSection, iKeyName, vbNullString, m_File
End Function

