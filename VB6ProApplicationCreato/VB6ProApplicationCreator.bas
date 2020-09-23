Attribute VB_Name = "modProsubs"
'------------------------------------------------------------
' modProsubs.BAS
'   DAO (Data Access Objects) in Visual Basic 6.0.
'
'------------------------------------------------------------
Option Explicit
'Public database variables

Public gdbCurrentDB  As Database    'main database object
Public gsAppName As String

'Public constants

Public Const APP_CATEGORY = "Visual Basic Program"
Public sDatabaseName As String
Public gsAppPath As String
Public gsMDBPath As String
Public cnn    As ADODB.Connection
Public dbpathe    As String              'Database Path

Public ErrGen As Boolean
Public Sub Connect()

  On Error GoTo CnnError

Dim strCnn As String
Dim psw    As String

    strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;"
    strCnn = strCnn & "Data Source=" & Chr$(34) & sDatabaseName & Chr$(34) & ";"
    strCnn = strCnn & "Jet OLEDB:Engine Type=5;"

    Set cnn = New ADODB.Connection
    cnn.Open strCnn

Exit Sub
CnnError:

    Select Case Err
    Case Is = -2147217843 'Database password incorrect
        '      psw = ObtainPassword
        strCnn = vbNullString
        strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;"
        strCnn = strCnn & "Data Source=" & Chr$(34) & sDatabaseName & Chr$(34) & ";"
        strCnn = strCnn & "Jet OLEDB:Engine Type=5;"
        strCnn = strCnn & psw
        If LenB(psw) = 0 Then
            Resume Next
        Else
            Resume
        End If
    Case Else
        MsgBox "Error Number : " & Err & Error, vbCritical, Err.Source
        End
    End Select

End Sub

'Copy the Access database file to project directory,
'This will create sub directory in the app's location
'and put all files (.frm, .bas, vbp, mdb,

Public Function AddBrackets(rObjName As String) As String

'------------------------------------------------------------
'this functions adds [] to object names that might need
'them because they have spaces in them
'------------------------------------------------------------

'add brackets to object names w/ spaces in them

    If InStr(rObjName, " ") > 0 And Mid$(rObjName, 1, 1) <> "[" Then
        AddBrackets = "[" & rObjName & "]"
    Else
        AddBrackets = rObjName
    End If

End Function

Public Sub CreateProjectDirectory()


Dim fso As FileSystemObject

    Set fso = CreateObject("Scripting.FileSystemObject")
    gsAppPath = App.Path & "\" & gsAppName

    If Not (fso.FolderExists(gsAppPath)) Then
        fso.CreateFolder (gsAppPath)
    End If

    frmPAWndc.lblAppPath = gsAppPath

End Sub

Public Function FileExists(ByVal sPathName As String) As Integer

'-----------------------------------------------------------
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
Dim intFileNum As Integer

    On Error Resume Next
    '
    'Remove trailing directory separator character
    '
    If Right$(sPathName, 1) = "\" Then
        sPathName = Left$(sPathName, Len(sPathName) - 1)
    End If
    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open sPathName For Input As intFileNum
    FileExists = IIf(Err, False, True)
    Close intFileNum
    Err = 0

End Function

Private Function GetINIString(ByVal vsItem As String, ByVal vsDefault As String) As String

'------------------------------------------------------------
'this function returns the INI file setting for the
'passed in item and section
'------------------------------------------------------------

    GetINIString = GetSetting(APP_CATEGORY, App.Title, vsItem, vsDefault)

End Function

Public Function RemoveCRLF(ByVal rvntVal As String) As String

' the function removes CR and LF from a string
'
Dim i    As Integer
Dim sTmp As String

    For i = 1 To Len(rvntVal)

        If Asc(Mid$(rvntVal, i, 1)) = 10 Then
            sTmp = sTmp & " "
        ElseIf Asc(Mid$(rvntVal, i, 1)) = 13 Then
        Else
            sTmp = sTmp & Mid$(rvntVal, i, 1)
        End If

    Next i
    RemoveCRLF = sTmp

End Function

Public Function RemoveSpace(ByVal rvntVal As Variant) As String

'--------------------------------------------------------------------------
' function remove all spaces from string
'--------------------------------------------------------------------------
Dim i    As Integer
Dim sTmp As String

    For i = 1 To Len(rvntVal)

        If Asc(Mid$(rvntVal, i, 1)) = 32 Or Mid$(rvntVal, i, 1) = "-" Then
            'skip
        Else
            sTmp = sTmp & Mid$(rvntVal, i, 1)
        End If

    Next i
    RemoveSpace = sTmp

End Function

Private Function StripBrackets(rsObjName As String) As String

'------------------------------------------------------------
'this function strips the [] off of data objects
'------------------------------------------------------------

'add brackets to object names w/ spaces in them

    If Mid$(rsObjName, 1, 1) = "[" Then
        StripBrackets = Mid$(rsObjName, 2, Len(rsObjName) - 2)
    Else
        StripBrackets = rsObjName
    End If

End Function

Private Function StripConnect(rsTblName As String) As String

'------------------------------------------------------------
'this function strips the attached table connect string off
'------------------------------------------------------------

    If InStr(rsTblName, "->") > 0 Then
        StripConnect = Left$(rsTblName, InStr(rsTblName, "->") - 2)
    Else
        StripConnect = rsTblName
    End If

End Function

Public Function Stripext(ByVal rsTblName As String) As String

'------------------------------------------------------------
'strips the owner off of ODBC table names
'------------------------------------------------------------

    If InStr(rsTblName, ".") > 0 Then
        rsTblName = Left$(rsTblName, InStr(rsTblName, ".") - 1)
    End If

    Stripext = rsTblName

End Function

Private Function StripFileName(ByVal rsFileName As String) As String

'------------------------------------------------------------
'this function strips the file name from a path\file string
'------------------------------------------------------------
Dim i As Integer

    On Error Resume Next

    For i = Len(rsFileName) To 1 Step -1

        If Mid$(rsFileName, i, 1) = "\" Then
            Exit For
        End If

    Next i
    StripFileName = Mid$(rsFileName, 1, i - 1)

End Function

Private Function StripNonAscii(rvntVal As Variant) As String

'------------------------------------------------------------
'this function strips the non ACSII chars off memo field
'data before displaying it (not sure this is always needed)
'------------------------------------------------------------
Dim i    As Integer
Dim sTmp As String

    'stubbed out to enable DBCS chars
    StripNonAscii = rvntVal

Exit Function

    For i = 1 To Len(rvntVal)

        If Asc(Mid$(rvntVal, i, 1)) < 32 Or Asc(Mid$(rvntVal, i, 1)) > 126 Then
            sTmp = sTmp & " "
        Else
            sTmp = sTmp & Mid$(rvntVal, i, 1)
        End If

    Next i
    StripNonAscii = sTmp

End Function

Private Function StripOwner(ByVal rsTblName As String) As String

'------------------------------------------------------------
'strips the owner off of ODBC table names
'------------------------------------------------------------

    If InStr(rsTblName, ".") > 0 Then
        rsTblName = Mid$(rsTblName, InStr(rsTblName, ".") + 1, Len(rsTblName))
    End If

    StripOwner = rsTblName

End Function

Public Function StripPath(ByVal T As String) As String

'--------------------------------------------------------------------------
'This will get only name of file from a complete
'file name with its directory
'--------------------------------------------------------------------------
Dim x  As Integer
Dim ct As Integer

    StripPath$ = T$
    x% = InStr(T$, "\")

    Do While x%
        ct% = x%
        x% = InStr(ct% + 1, T$, "\")
    Loop

    If ct% > 0 Then StripPath$ = Mid$(T$, ct% + 1)

End Function

Private Function stTrueFalse(rvntTF As Variant) As String

'------------------------------------------------------------
'returns the true or false string
'------------------------------------------------------------

    If rvntTF Then
        stTrueFalse = "True"
    Else
        stTrueFalse = "False"
    End If

End Function

Private Sub UnloadAllForms()

'------------------------------------------------------------
'this sub unloads all forms except for the
'SQL, Tables and MDI form
'------------------------------------------------------------
Dim i As Integer

    On Error Resume Next

    'close all forms except for the Tables and SQL forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next i

End Sub

