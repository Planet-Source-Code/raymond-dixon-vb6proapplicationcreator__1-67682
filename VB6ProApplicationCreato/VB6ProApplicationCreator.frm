VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPAWndc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pro Application Creator"
   ClientHeight    =   7980
   ClientLeft      =   1305
   ClientTop       =   1725
   ClientWidth     =   11385
   HelpContextID   =   2018517
   Icon            =   "VB6ProApplicationCreator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7980
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   35
      Top             =   6660
      Width           =   8055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   34
      Top             =   6180
      Width           =   8055
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options"
      ForeColor       =   &H0000FFFF&
      Height          =   2295
      Left            =   420
      TabIndex        =   25
      Top             =   2160
      Width           =   6435
      Begin VB.OptionButton optADO 
         Caption         =   "Create ADO Application"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   600
         TabIndex        =   33
         Top             =   1800
         Width           =   5475
      End
      Begin VB.OptionButton optDAO 
         Caption         =   "Create DAO Application"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   600
         TabIndex        =   32
         Top             =   1380
         Value           =   -1  'True
         Width           =   4995
      End
      Begin VB.CommandButton cmdClean 
         Caption         =   "CLEAN"
         Height          =   300
         Left            =   5160
         TabIndex        =   31
         Top             =   780
         Width           =   1155
      End
      Begin VB.CheckBox chkCopyMDB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Copy MDB to Project DIR"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   30
         Top             =   240
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CommandButton cmdREG 
         Caption         =   "REG"
         Height          =   300
         Left            =   5160
         TabIndex        =   29
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "HELP"
         Height          =   300
         Left            =   5160
         TabIndex        =   28
         Top             =   480
         Width           =   1155
      End
      Begin VB.CheckBox chkComments 
         Caption         =   "Add coments to code"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   600
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkCreateFilemod 
         Caption         =   "Create Sub Module code for creating Database"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   600
         TabIndex        =   26
         Top             =   960
         Value           =   1  'Checked
         Width           =   3735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10080
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressB 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   7665
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   3600
      TabIndex        =   6
      Top             =   4560
      Width           =   1515
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Cancel Clean Files"
      Height          =   435
      Index           =   1
      Left            =   9180
      TabIndex        =   15
      Top             =   4620
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Clean All Files"
      Height          =   435
      Index           =   0
      Left            =   7620
      TabIndex        =   14
      Top             =   4620
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton cmdBuildApp 
      Caption         =   "Create Application"
      Height          =   435
      Left            =   1680
      TabIndex        =   5
      Top             =   4560
      Width           =   1515
   End
   Begin VB.CommandButton cmdOpenDB 
      Caption         =   "&Open Database"
      Height          =   435
      Left            =   1140
      TabIndex        =   9
      Top             =   4560
      Width           =   1515
   End
   Begin VB.Frame Frame5 
      Height          =   3615
      Left            =   6960
      TabIndex        =   11
      Top             =   840
      Width           =   4035
      Begin VB.ListBox lstClean 
         Height          =   1815
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   12
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label lblProcess 
         BackStyle       =   0  'Transparent
         Height          =   555
         Left            =   300
         TabIndex        =   16
         Top             =   2820
         Width           =   3435
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Files Created  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   3915
      End
   End
   Begin MSComDlg.CommonDialog dlgDBOpen 
      Left            =   60
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   1800
      TabIndex        =   0
      Top             =   8280
      Visible         =   0   'False
      Width           =   6075
      Begin VB.ComboBox cboRecordSource 
         Height          =   315
         Left            =   600
         TabIndex        =   18
         Text            =   "cboRecordSource"
         Top             =   1440
         Width           =   4335
      End
      Begin VB.ListBox lstFields 
         Height          =   255
         Left            =   240
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.ListBox lstOLECtls 
         Height          =   255
         Left            =   2880
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Available Fields: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   120
         Width           =   1470
      End
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Database Destination Dir"
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   6840
      Width           =   2190
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Database Source Dir"
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   6360
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Updated in 2006 for vb6 By: Raymond E Dixon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   420
      Left            =   480
      TabIndex        =   24
      Top             =   1680
      Width           =   6315
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "App DIR"
      Height          =   195
      Left            =   360
      TabIndex        =   23
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label lblAppPath 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "app dir"
      Height          =   315
      Left            =   1320
      TabIndex        =   22
      Top             =   5220
      Width           =   9975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Written in VB4 1996 By: Raymond E Dixon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   420
      Left            =   480
      TabIndex        =   21
      Top             =   1200
      Width           =   6315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "No Data Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2700
      TabIndex        =   20
      Top             =   840
      Width           =   2220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   840
      X2              =   6540
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lbFormName 
      Alignment       =   2  'Center
      Caption         =   "App Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7080
      TabIndex        =   19
      Top             =   300
      Width           =   3855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Database Dir"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   5640
      Width           =   930
   End
   Begin VB.Label lblDatabaseName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Database Name"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   5580
      Width           =   9975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "VB6 Pro Application Creator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   1140
      TabIndex        =   4
      Top             =   180
      Width           =   4920
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "VB6 Pro Application Creator FINISHED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Visible         =   0   'False
      Width           =   6810
   End
End
Attribute VB_Name = "frmPAWndc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This code is FREE to use or modify, but if you sell it please send me a check !!

Option Explicit
Private currentRS As Recordset 'current recordset
Private sSQLStatement As String
Private giFormNumber As Integer ' number added to form name
Private gsFormName As String 'form name
Private giCurrentTableNumber As Integer ' current index number
Private giDataType As Integer
Private gsRecordTable As String

'constants used for the data type of the database
Const gnDT_ACCESS = 0

' This code is FREE to use or modify, but if you sell it please send me a check !!

Private Sub BuildVBPProjectCode()

'--------------------------------------------------------------------------
'This will generate reference and form to project
'--------------------------------------------------------------------------


Dim nFileHnd As Integer


    nFileHnd = FreeFile
    If FileExists(gsAppPath & "\vbp" & gsAppName & ".vbp") Then Kill gsAppPath & "\vbp" & gsAppName & ".vbp"
    'This will create a project file with extension .vbp
    Open gsAppPath & "\vbp" & gsAppName & ".vbp" For Output As #nFileHnd
    Print #nFileHnd, "Type=Exe"

    If optDAO.Value = True Then

        Print #nFileHnd, "Reference=*\G{00025E01-0000-0000-C000-000000000046}#5.0#0#..\..\..\..\Common" & " Files\Microsoft Shared\DAO\dao360.dll#Microsoft DAO 3.6 Object Library"
        Print #nFileHnd, "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\..\..\..\..\..\WINDOWS\system32\stdole2.tlb#OLE Automation"
        Print #nFileHnd, "Object={CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0; MSDATGRD.OCX"
        Print #nFileHnd, "Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX"

    Else
        Print #nFileHnd, "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\..\..\..\..\..\..\WINDOWS\system32\stdole2.tlb#OLE Automation"
        Print #nFileHnd, "Reference=*\G{00000201-0000-0010-8000-00AA006D2EA4}#2.1#0#..\..\..\..\..\..\Common Files\system\ado\msado21.tlb#Microsoft ActiveX Data Objects 2.1 Library"
        Print #nFileHnd, "Reference=*\G{00000600-0000-0010-8000-00AA006D2EA4}#2.1#0#..\..\..\..\..\..\Common Files\System\ado\msADOX.dll#Microsoft ADO Ext. 2.1 for DDL and Security"

    End If

    Print #nFileHnd, "Form=frmMDI.frm"
    gsRecordTable = ""

    For giCurrentTableNumber = 0 To cboRecordSource.ListCount - 1
        giFormNumber = giCurrentTableNumber
        gsRecordTable = cboRecordSource.List(giCurrentTableNumber)
        gsRecordTable = RemoveSpace(gsRecordTable & Trim$(Str$(giFormNumber)))
        DoEvents
        Print #nFileHnd, "Form=frm" & gsRecordTable & ".frm"
    Next giCurrentTableNumber

    Print #nFileHnd, "Form=frmAbout.frm"
    Print #nFileHnd, "Form=frmWait.FRM"
    Print #nFileHnd, "Form=frmSearch.FRM"
    Print #nFileHnd, "Module=PublicMod; modGlblMod.bas"
    Print #nFileHnd, "Object={F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0; COMDLG32.OCX"
    Print #nFileHnd, "IconForm=frmMDI"
    Print #nFileHnd, "Startup=`````Sub Main"
    Print #nFileHnd, "HelpFile="""""
    Print #nFileHnd, "Title=""vbp" & gsAppName & """"
    Print #nFileHnd, "ExeName32=""" & gsAppName & ".exe"""
    Print #nFileHnd, "Command32="""""
    Print #nFileHnd, "Name=""vbp" & gsAppName & """"
    Print #nFileHnd, "HelpContextID=""0"""
    Print #nFileHnd, "CompatibleMode=""0"""
    Print #nFileHnd, "MajorVer=1"
    Print #nFileHnd, "MinorVer=0"
    Print #nFileHnd, "RevisionVer=0"
    Print #nFileHnd, "AutoIncrementVer=0"
    Print #nFileHnd, "ServerSupportFiles=0"
    Print #nFileHnd, "VersionCompanyName="
    Print #nFileHnd, "CompilationType=0"
    Print #nFileHnd, "OptimizationType=0"
    Print #nFileHnd, "FavorPentiumPro(tm)=0"
    Print #nFileHnd, "CodeViewDebugInfo=0"
    Print #nFileHnd, "NoAliasing=0"
    Print #nFileHnd, "BoundsCheck=0"
    Print #nFileHnd, "OverflowCheck=0"
    Print #nFileHnd, "FlPointCheck=0"
    Print #nFileHnd, "FDIVCheck=0"
    Print #nFileHnd, "UnroundedFP=0"
    Print #nFileHnd, "StartMode=0"
    Print #nFileHnd, "Unattended=0"
    Print #nFileHnd, "Retained=0"
    Print #nFileHnd, "ThreadPerObject=0"
    Print #nFileHnd, "MaxNumberOfThreads=1"
    Close #nFileHnd

End Sub

Private Sub BuildNoDataFormdao()


Dim i          As Integer
Dim sTmp       As String
Dim nNumFlds   As Integer
Dim frmNewForm As Object
Dim nButtonTop As Integer
Dim nFileHnd   As Integer


    On Error GoTo BuildFErr

    'create and open the file

    nFileHnd = FreeFile
    gsFormName = gsRecordTable & Trim$(Str$(giFormNumber))
    Open gsAppPath & "\frm" & gsFormName & ".FRM" For Output As nFileHnd
    Print #nFileHnd, "VERSION 5.00"
    nNumFlds = lstFields.ListCount
    lstOLECtls.Clear
    Print #nFileHnd, "Begin VB.Form frm" & Left$(RemoveSpace(currentRS.Name), 37)
    Print #nFileHnd, "   BorderStyle  = 3 'fixed dialog"
    Print #nFileHnd, "   Caption = """ & Left$(currentRS.Name, 32) & """"

    If nNumFlds > 13 Then
        Print #nFileHnd, "   ClientHeight       = " & 1615 + (14 * 320)
    Else
        Print #nFileHnd, "   ClientHeight       = " & 1615 + (nNumFlds * 320)
    End If

    Print #nFileHnd, "   ClientLeft         = 700"
    Print #nFileHnd, "   Top          = 700"

    If nNumFlds > 13 And nNumFlds < 28 Then
        Print #nFileHnd, "   ClientWidth        = 8370"
    ElseIf nNumFlds > 27 Then
        Print #nFileHnd, "   ClientWidth        = 13350"
    Else
        Print #nFileHnd, "   ClientWidth        = 7215"
    End If

    Print #nFileHnd, "   MDIchild     = -1 'true"

    If nNumFlds > 13 Then
        Print #nFileHnd, "   ScaleHeight       = " & 1615 + (13 * 320)
    Else
        Print #nFileHnd, "   ScaleHeight       = " & 1615 + (nNumFlds * 320)
    End If

    Print #nFileHnd, "   Top          = 700"

    If nNumFlds > 13 And nNumFlds < 28 Then
        Print #nFileHnd, "   ScaleWidth        = 8370"
    ElseIf nNumFlds > 27 Then
        Print #nFileHnd, "   ScaleWidth        = 13350"
    Else
        Print #nFileHnd, "   ScaleWidth        = 7215"
    End If
    Print #nFileHnd, "      Begin VB.Label Label1"
    Print #nFileHnd, "       Alignment = 2          'Center"
    Print #nFileHnd, "     Caption = """ & Left$(currentRS.Name, 32) & """"
    Print #nFileHnd, "      BeginProperty Font"
    Print #nFileHnd, "        Name = "" MS Sans Serif """
    Print #nFileHnd, "        Size = 13.5"
    Print #nFileHnd, "       Charset = 0"
    Print #nFileHnd, "            Weight = 400"
    Print #nFileHnd, "            Underline = 0           'False"
    Print #nFileHnd, "            Italic = 0              'False"
    Print #nFileHnd, "            Strikethrough = 0       'False"
    Print #nFileHnd, "        EndProperty"
    Print #nFileHnd, "        ForeColor = &HFF&"
    Print #nFileHnd, "        Height = 375"
    Print #nFileHnd, "       Left = 0"
    Print #nFileHnd, "       TabIndex = 24"
    Print #nFileHnd, "       Top = 0"
    Print #nFileHnd, "       Width = 7215"
    Print #nFileHnd, "     End"

    For i = 0 To nNumFlds - 1
        sTmp = lstFields.List(i)
        Print #nFileHnd, "   Begin VB.Label lblLabels"
        Print #nFileHnd, "      Alignment  = 1"
        Print #nFileHnd, "      Autosize   = 1 'true"
        Print #nFileHnd, "      Caption = """ & sTmp & ":"""
        Print #nFileHnd, "      Height  = 255"
        Print #nFileHnd, "      Index   = " & i

        If i > 13 And i < 28 Then
            Print #nFileHnd, "      Left    = 4080"
            Print #nFileHnd, "      Top     = " & ((i - 14) * 320) + 420
        ElseIf i > 27 Then
            Print #nFileHnd, "      Left    = 9060"
            Print #nFileHnd, "      Top     = " & ((i - 28) * 320) + 420
        Else
            Print #nFileHnd, "      Left    = 120"
            Print #nFileHnd, "      Top     = " & (i * 320) + 420
        End If

        Print #nFileHnd, "      Width   = 1815"
        Print #nFileHnd, "   End"

        If currentRS.Fields(sTmp).Type = 1 Then
            'true/false field
            Print #nFileHnd, "   Begin VB.CheckBox chkField" ' & i
            Print #nFileHnd, "      Height     = 285"
            Print #nFileHnd, "      Index      = " & i

            If i > 13 And i < 28 Then
                Print #nFileHnd, "      Left    = 6060"
                Print #nFileHnd, "      Top     = " & ((i - 14) * 320) + 420
            ElseIf i > 27 Then
                Print #nFileHnd, "      Left    = 11000"
                Print #nFileHnd, "      Top     = " & ((i - 28) * 320) + 420
            Else
                Print #nFileHnd, "      Left    = 2120"
                Print #nFileHnd, "      Top     = " & (i * 320) + 420
            End If

            Print #nFileHnd, "      Width      = 1875"
            Print #nFileHnd, "   End"
        ElseIf currentRS.Fields(sTmp).Type = 11 Then
            'picture field
            Print #nFileHnd, "   Begin VB.OLE oleField" & i
            Print #nFileHnd, "      Height         = 285"
            Print #nFileHnd, "      OLETypeAllowed = 1"

            If i > 13 And i < 28 Then
                Print #nFileHnd, "      Left    = 6060"
                Print #nFileHnd, "      Top     = " & ((i - 14) * 320) + 420
            ElseIf i > 27 Then
                Print #nFileHnd, "      Left    = 11000"
                Print #nFileHnd, "      Top     = " & ((i - 28) * 320) + 420
            Else
                Print #nFileHnd, "      Left    = 2120"
                Print #nFileHnd, "      Top     = " & (i * 320) + 420
            End If

            Print #nFileHnd, "      Width          = 3375"
            Print #nFileHnd, "   End"
            lstOLECtls.AddItem i
        Else
            Print #nFileHnd, "   Begin VB.TextBox txtField" '& i

            If currentRS.Fields(sTmp).Type = 12 Then
                Print #nFileHnd, "      Height     = 310"
            Else
                Print #nFileHnd, "      Height     = 285"
            End If

            Print #nFileHnd, "      Index      = " & i
            Print #nFileHnd, "      Left       = 2040"

            If currentRS.Fields(sTmp).Type = 10 Then
                Print #nFileHnd, "      MaxLength   = " & currentRS.Fields(sTmp).Size
            End If

            If currentRS.Fields(sTmp).Type = 12 Then
                Print #nFileHnd, "      MultiLine   = True"
            End If

            If currentRS.Fields(sTmp).Type = 12 Then
                Print #nFileHnd, "      ScrollBars  = 2"
            End If

            If i > 13 And i < 28 Then
                Print #nFileHnd, "      Left    = 6060"
                Print #nFileHnd, "      Top     = " & ((i - 14) * 320) + 420
            ElseIf i > 27 Then
                Print #nFileHnd, "      Left    = 11000"
                Print #nFileHnd, "      Top     = " & ((i - 28) * 320) + 420
            Else
                Print #nFileHnd, "      Left    = 2040"
                Print #nFileHnd, "      Top     = " & (i * 320) + 420
            End If

            If currentRS.Fields(sTmp).Type < 10 Then
                'numeric or date
                Print #nFileHnd, "      Width      = 935"
            Else
                'string or memo
                Print #nFileHnd, "      Width      = 1875"
            End If

            Print #nFileHnd, "   End"
        End If

    Next i
    nButtonTop = (((i - 1) * 320) + 40) + 340
    '************************************
    ' start new code below
    ' Buttons added to Bottom of form
    '**************************************
    'add the data control and buttons
    Print #nFileHnd, "   Begin VB.PictureBox CommandBar"
    Print #nFileHnd, "      Align           =   2  'Align Bottom"; 1 'Align top"
    Print #nFileHnd, "      Height          =   375"
    Print #nFileHnd, "      Left            =   0"
    Print #nFileHnd, "      TabIndex        =   45"
    Print #nFileHnd, "      Top             =   0"
    Print #nFileHnd, "      BackColor       =   &H8000000C&"

    Print #nFileHnd, "   Begin VB.CommandButton cmdData"
    Print #nFileHnd, "      Caption         =   ""Exit"""
    Print #nFileHnd, "      Height          =   300"
    Print #nFileHnd, "      Index           =   0"
    Print #nFileHnd, "      Left            =   0"
    Print #nFileHnd, "      TabIndex        =   1"
    Print #nFileHnd, "      Top             =   0"
    Print #nFileHnd, "      Width           =   630"
    Print #nFileHnd, "   End"
    Print #nFileHnd, "   Begin VB.CommandButton cmdData"
    Print #nFileHnd, "      Caption         =   ""Print"""
    Print #nFileHnd, "      Height          =   300"
    Print #nFileHnd, "      Index           =   1"
    Print #nFileHnd, "      Left            =   630"
    Print #nFileHnd, "      TabIndex        =   2"
    Print #nFileHnd, "      Top             =   0"
    Print #nFileHnd, "      Width           =   630"
    Print #nFileHnd, "   End"
    Print #nFileHnd, "   Begin VB.CommandButton cmdData"
    Print #nFileHnd, "      Caption         =   ""Find"""
    Print #nFileHnd, "      Height          =   300"
    Print #nFileHnd, "      Index           =   2"
    Print #nFileHnd, "      Left            =   1260"
    Print #nFileHnd, "      TabIndex        =   3"
    Print #nFileHnd, "      Top             =   0"
    Print #nFileHnd, "      Width           =   630"
    Print #nFileHnd, "   End"
    Print #nFileHnd, "   Begin VB.CommandButton cmdData"
    Print #nFileHnd, "      Caption         =   ""Add"""
    Print #nFileHnd, "      Height          =   300"
    Print #nFileHnd, "      Index           =   3"
    Print #nFileHnd, "      Left            =   1890"
    Print #nFileHnd, "      TabIndex        =   4"
    Print #nFileHnd, "      Top             =   0"
    Print #nFileHnd, "      Width           =   630"
    Print #nFileHnd, "   End"
    Print #nFileHnd, "   Begin VB.CommandButton cmdData"
    Print #nFileHnd, "      Caption         =   ""Del"""
    Print #nFileHnd, "      Height          =   300"
    Print #nFileHnd, "      Index           =   4"
    Print #nFileHnd, "      Left            =   2520"
    Print #nFileHnd, "      TabIndex        =   5"
    Print #nFileHnd, "      Top             =   0"
    Print #nFileHnd, "      Width           =   630"
    Print #nFileHnd, "   End"
    Print #nFileHnd, "   Begin VB.CommandButton cmdData"
    Print #nFileHnd, "      Caption         =   ""||<< """
    Print #nFileHnd, "      Height          =   300"
    Print #nFileHnd, "      Index           =   5"
    Print #nFileHnd, "      Left            =   3150"
    Print #nFileHnd, "      TabIndex        =   6"
    Print #nFileHnd, "      Top             =   0"
    Print #nFileHnd, "      Width           =   630"
    Print #nFileHnd, "   End"
    Print #nFileHnd, "   Begin VB.CommandButton cmdData"
    Print #nFileHnd, "      Caption         =   ""<<"""
    Print #nFileHnd, "      Height          =   300"
    Print #nFileHnd, "      Index           =   6"
    Print #nFileHnd, "      Left            =   3780"
    Print #nFileHnd, "      TabIndex        =   7"
    Print #nFileHnd, "      Top             =   0"
    Print #nFileHnd, "      Width           =   630"
    Print #nFileHnd, "   End"
    Print #nFileHnd, "   Begin VB.CommandButton cmdData"
    Print #nFileHnd, "      Caption         =   "">>"""
    Print #nFileHnd, "      Height          =   300"
    Print #nFileHnd, "      Index           =   7"
    Print #nFileHnd, "      Left            =   4410"
    Print #nFileHnd, "      TabIndex        =   8"
    Print #nFileHnd, "      Top             =   0"
    Print #nFileHnd, "      Width           =   630"
    Print #nFileHnd, "   End"
    Print #nFileHnd, "   Begin VB.CommandButton cmdData"
    Print #nFileHnd, "      Caption         =   "">>||"""
    Print #nFileHnd, "      Height          =   300"
    Print #nFileHnd, "      Index           =   8"
    Print #nFileHnd, "      Left            =   5040"
    Print #nFileHnd, "      TabIndex        =   9"
    Print #nFileHnd, "      Top             =   0"
    Print #nFileHnd, "      Width           =   630"
    Print #nFileHnd, "   End"
    Print #nFileHnd, "   Begin VB.CommandButton cmdData"
    Print #nFileHnd, "      Caption         =   ""Can"""
    Print #nFileHnd, "      Height          =   300"
    Print #nFileHnd, "      Index           =   9"
    Print #nFileHnd, "      Left            =   5670"
    Print #nFileHnd, "      TabIndex        =   10"
    Print #nFileHnd, "      Top             =   0"
    Print #nFileHnd, "      Width           =   630"
    Print #nFileHnd, "   End"
    Print #nFileHnd, "   Begin VB.CommandButton cmdData"
    Print #nFileHnd, "      Caption         =   ""Updt"""
    Print #nFileHnd, "      Height          =   300"
    Print #nFileHnd, "      Index           =   10"
    Print #nFileHnd, "      Left            =   6300"
    Print #nFileHnd, "      TabIndex        =   11"
    Print #nFileHnd, "      Top             =   0"
    Print #nFileHnd, "      Width           =   630"
    Print #nFileHnd, "   End"
    Print #nFileHnd, "End"
    Print #nFileHnd, "   Begin VB.PictureBox StatusBar"
    Print #nFileHnd, "      Align           =   2  'Align Bottom"
    Print #nFileHnd, "      Height          =   315"
    Print #nFileHnd, "      Left            =   0"
    Print #nFileHnd, "      ScaleHeight     =   255"
    Print #nFileHnd, "      ScaleWidth      =   8190"
    Print #nFileHnd, "      TabIndex        =   46"
    Print #nFileHnd, "      Top             =   " & nButtonTop
    Print #nFileHnd, "      Begin VB.Label MsgBar"
    Print #nFileHnd, "         BackColor       =   &H00FFFFFF&"
    Print #nFileHnd, "         BorderStyle     =   1  'Fixed Single"
    Print #nFileHnd, "         Caption         =   ""MsgBar"""
    Print #nFileHnd, "         Height          =   255"
    Print #nFileHnd, "         Left            =   60"
    Print #nFileHnd, "         TabIndex        =   47"
    Print #nFileHnd, "         Top             =   0"
    Print #nFileHnd, "         Width           =   2715"
    Print #nFileHnd, "      End"
    Print #nFileHnd, "   End"
    ' end of form
    Print #nFileHnd, "End"
    '*********************************************
    ' end of added code
    '*************************************
    Print #nFileHnd, ""
    Print #nFileHnd, "Attribute VB_Name = ""frm" & Left$(RemoveSpace(currentRS.Name), 37) & """"
    Print #nFileHnd, "Attribute VB_Creatable = False"
    Print #nFileHnd, "Attribute VB_Exposed = False"
    Print #nFileHnd, "Option Explicit"
    Print #nFileHnd, ""
    'add the code to the form
    BuildNoDataCode nFileHnd
    Close nFileHnd

Exit Sub

BuildFErr:
    MsgBox Error$

End Sub

Private Sub cboRecordSource_Change()

    Set currentRS = Nothing
    lstFields.Clear

End Sub

Private Sub cboRecordSource_Click()

    Call cboRecordSource_LostFocus

End Sub

Private Sub cboRecordSource_LostFocus()

Dim Fld As Field

    On Error GoTo RSErr

    If Len(gsRecordTable) = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    If currentRS Is Nothing Then
        Set currentRS = gdbCurrentDB.OpenRecordset(gsRecordTable)

        For Each Fld In currentRS.Fields
            lstFields.AddItem Fld.Name
        Next Fld

    ElseIf currentRS.Name <> gsRecordTable Then
        lstFields.Clear
        Set currentRS = gdbCurrentDB.OpenRecordset(gsRecordTable)

        For Each Fld In currentRS.Fields
            lstFields.AddItem Fld.Name
        Next Fld

    End If
    Screen.MousePointer = vbDefault

Exit Sub

RSErr:
    Screen.MousePointer = vbDefault
    MsgBox Error$

End Sub

Private Sub chkCopyMDB_Click()
If chkCopyMDB.Value = vbChecked Then
Text1.Visible = True
Text2.Visible = True
Label9.Visible = True
Label10.Visible = True

Else
Text1.Visible = False
Text2.Visible = False
Label9.Visible = False
Label10.Visible = False

End If

End Sub

Private Sub cmdBuildApp_Click()

Dim mfh1         As Integer
Dim mfh3         As Integer
Dim cfh          As Integer
Dim PercentTotal As Integer
Dim fh1          As Integer
Dim Fld          As Field
Dim gfh          As Integer
Dim tmpcap       As String
Dim buff         As String
Dim FileLength   As Long

    On Error Resume Next
    ' file numbers for open files

    cmdBuildApp.Visible = False
    cmdOpenDB.Visible = False
    cmdClose.Visible = False
    tmpcap = Label4.Caption
    If FileExists(gsAppPath & "\cleanapp.tmp") Then Kill gsAppPath & "\cleanapp.tmp"
    cfh = FreeFile
    Open gsAppPath & "\cleanapp.tmp" For Append As cfh
    If FileExists(gsAppPath & "\frmMDI.FRM") Then Kill gsAppPath & "\frmMDI.FRM"
    If FileExists(gsAppPath & "\frmMDI.tmp") Then Kill gsAppPath & "\frmMDI.tmp"
    'create first section of Project
    Print #cfh, "vbp" & gsAppName & ".vbp"
    Call BuildVBPProjectCode
    'create first section of frmMDI
    mfh1 = FreeFile
    Open gsAppPath & "\frmMDI.FRM" For Output As mfh1
    BuildMDIProjectCode1 mfh1
    Close #mfh1
    mfh3 = FreeFile
    Open gsAppPath & "\frmMDI.tmp" For Output As mfh3
    BuildMDIProjectCode3 mfh3
    Close #mfh3
    mfh1 = FreeFile
    Open gsAppPath & "\frmMDI.FRM" For Append As mfh1
    mfh3 = FreeFile
    Open gsAppPath & "\frmMDI.tmp" For Append As mfh3
    'create and open the file
    gfh = FreeFile
    Open gsAppPath & "\modGlblMod.bas" For Output As gfh
    BuildPublicModNDC gfh
    Close gfh
    gfh = FreeFile
    Open gsAppPath & "\modGlblMod.bas" For Append As gfh
    'clean up before start
    giFormNumber = 0
    Set currentRS = Nothing
    lstFields.Clear
    PercentTotal = cboRecordSource.ListCount

    For giCurrentTableNumber = 0 To cboRecordSource.ListCount - 1
        giFormNumber = giCurrentTableNumber
        gsRecordTable = ""
        gsRecordTable = cboRecordSource.List(giCurrentTableNumber)
        ProgressB.Value = 1 / (PercentTotal - giCurrentTableNumber) * 100
        Label4.Caption = "Createing Table: " & gsRecordTable
        DoEvents

        If currentRS Is Nothing Then
            Set currentRS = gdbCurrentDB.OpenRecordset(gsRecordTable)

            For Each Fld In currentRS.Fields
                lstFields.AddItem Fld.Name
            Next Fld

        ElseIf currentRS.Name <> gsRecordTable Then
            lstFields.Clear
            Set currentRS = gdbCurrentDB.OpenRecordset(gsRecordTable)

            For Each Fld In currentRS.Fields
                lstFields.AddItem Fld.Name
            Next Fld

        End If

        If cboRecordSource.ItemData(giCurrentTableNumber) = 1 Then
            'copy the sql from the querydef to the SQL window
            sSQLStatement = gdbCurrentDB.QueryDefs(gsRecordTable).SQL
        Else
            sSQLStatement = ""
        End If

        If Len(gsRecordTable) = 0 Then
            MsgBox "You must enter a RecordSource!", 16
            Exit Sub
        End If

        gsRecordTable = RemoveSpace(gsRecordTable)

        If Len(gsRecordTable) = 0 Then
            MsgBox "Form Name cannot be blank!", 16
            Exit Sub
        End If

        If lstFields.ListCount = 0 Then
            MsgBox "You must include some Columns!", 16
            Exit Sub
        End If

        BuildNoDataFormdao
        ' we have to write mdi form in 4 parts
        ' to add all the menu items names
        BuildMDIProjectCode2 mfh1
        ' we have to write mdi form in 4 parts
        ' to add all the click subs
        BuildMDIProjectCode4 mfh3
        ' write code for forms

        If optDAO.Value = True Then

            BuildCreateDaoSub gfh
        End If

        Print #cfh, "frm" & gsFormName & ".FRM"
    Next giCurrentTableNumber

    'we are through with form module
    'start with code
    BuildSubMainCreatedb gfh
    If optADO.Value = True Then
        BuildCreateADOSub gfh
    End If

    Close gfh ' close Public module
    ' add the new form to the project
    Print #cfh, "modGlblMod.bas"
    ' search form is generic
    WriteSearchForm
    Print #cfh, "frmSearch.frm"
    ' about form is generic
    fh1 = FreeFile
    Open gsAppPath & "\" & "frmAbout.FRM" For Output As #fh1
    WriteAboutForm fh1
    Close #fh1
    Print #cfh, "frmAbout.frm"
    ' wait form is generic
    fh1 = FreeFile
    Open gsAppPath & "\" & "frmWait.FRM" For Output As #fh1
    WriteWaitForm fh1
    Close #fh1
    Print #cfh, "frmWait.frm"
    Close #mfh3
    mfh3 = FreeFile
    Open gsAppPath & "\frmMDI.tmp" For Input As mfh3
    FileLength = LOF(mfh3)
    ' copy all frmMDI to main form

    While Not EOF(mfh3)
        Line Input #mfh3, buff
        Print #mfh1, buff
    Wend

    Close #mfh1
    Close #mfh3
    'add the new form to the project
    Print #cfh, "frmMDI.frm"
    Close #cfh

    On Error Resume Next
    'clean up
    Kill gsAppPath & "\frmMDI.tmp"
    Frame4.Visible = True
    'set focus back to the form
    Label2.Visible = False
    Label5.Visible = True
    Me.SetFocus
    Beep ' let user know were through
    Label4.Caption = tmpcap
    cmdBuildApp.Visible = False
    cmdOpenDB.Visible = False
    cmdClose.Visible = True
    cmdClose.SetFocus
    Set gdbCurrentDB = Nothing

End Sub

Private Sub cmdClean_Click()

    Screen.MousePointer = 11 'set the hourglass
    cmdClose.Visible = False
    cmdOpenDB.Visible = False
    cmdBuildApp.Visible = False
    cmdDelete(0).Visible = True
    cmdDelete(1).Visible = True
    lstClean.ListIndex = 0
    Screen.MousePointer = 0 'unset the hourglass
    lblProcess.Caption = "Remove's All files created from database" & vbNewLine & """Reverse of Create" & " Application Creator"""

End Sub

Private Sub cmdClose_Click()

    On Error Resume Next
    gdbCurrentDB.Close
    Unload Me

End Sub

Private Sub cmdDelete_Click(Index As Integer)

Dim PercentTotal
Dim sFormName As String

    ProgressB.Min = 0
    Select Case Index
    Case 0 ' delete
        On Error Resume Next
        PercentTotal = lstClean.ListCount
        ProgressB.Max = PercentTotal

        For giCurrentTableNumber = 0 To lstClean.ListCount - 1
            giFormNumber = giCurrentTableNumber
            sFormName = ""
            sFormName = "\" & Trim$(lstClean.List(giCurrentTableNumber))
            ProgressB.Value = giCurrentTableNumber

            If FileExists(gsAppPath & sFormName) Then
                Kill gsAppPath & sFormName
            End If

        Next giCurrentTableNumber
        'clean up

        If FileExists(gsAppPath & "\cleanapp.tmp") Then
            Kill gsAppPath & "\cleanapp.tmp"
        End If

        lstClean.Clear
        ProgressB.Value = 0

    Case 1 ' cancel
        'cancel does nothing
    End Select

    cmdDelete(0).Visible = False
    cmdDelete(1).Visible = False
    lblProcess.Caption = ""
    cmdClean.Enabled = True
    cmdBuildApp.Visible = False
    cmdOpenDB.Visible = True
    cmdClose.Visible = True
    cmdClose.SetFocus
    Call Form_Load

End Sub

Private Sub cmdInfo_Click()

Dim msg      As String
Dim response As Integer

    On Error Resume Next

    msg = vbNewLine & "Builds complete application for All tables,queries and fields in  " & vbNewLine & "Access database.mdb including code to recreate database." & vbNewLine & "Makes app directory" & " and copies database and all project files to it.  " & vbNewLine & vbNewLine & "adds Number to form" & " name 'in case Querry' has same name. " & vbNewLine & vbNewLine & "Requirements: DAO 3.6 for access" & " 2000 database " & vbNewLine & "Requirements: Windows Scripting " & vbNewLine & "tvmancet@hotmail.com " & vbNewLine & vbNewLine & " "
    response = MsgBox(msg, vbOKOnly, "Quick Help")
    ' for real help
    '  Dim nRet As Integer
    ' ' nRet = OSWinHelp(Me.hWnd, App.HelpFile, HelpConstants.cdlHelpContents, 0)
    '  If Err Then
    '   ' ShowError
    '  End If

End Sub

Private Sub cmdOpenDB_Click()

Dim cfh      As Integer
Dim tdf      As TableDef
Dim qdf      As QueryDef
Dim cleanapp As String

    On Error GoTo OpenError

    '  Select Case giDataType
    dlgDBOpen.Filter = "Access DBs (*.mdb)|*.mdb|All Files (*.*)|*.*"
    dlgDBOpen.DialogTitle = "Open MS Access Database"

    With dlgDBOpen
        .FilterIndex = 1
        .FileName = ""
        .CancelError = True
        .Flags = cdlOFNExplorer
        .Action = 1
    End With

    sDatabaseName = dlgDBOpen.FileName
    gsMDBPath = dlgDBOpen.FileName
    lbFormName = Stripext(StripPath(sDatabaseName))
    gsAppName = Stripext(StripPath(sDatabaseName))
    
    CreateProjectDirectory
    
    If chkCopyMDB.Value = vbChecked Then
    Text1.Text = gsMDBPath
    Text2.Text = gsAppPath & "\" & StripPath(gsMDBPath)
    
        Call FileCopy(gsMDBPath, gsAppPath & "\" & StripPath(gsMDBPath))
    End If
    
    If FileExists(gsAppPath & "\cleanapp.tmp") Then
        cfh = FreeFile
        Open gsAppPath & "\cleanapp.tmp" For Input As #cfh

        Do While Not EOF(cfh)
            Line Input #cfh, cleanapp
            lstClean.AddItem cleanapp
        Loop

        Close #cfh
        cmdClean.Enabled = True
        cmdClean.Visible = True
    End If

    lblDatabaseName.Caption = sDatabaseName
    lstFields.Clear
    Me.Refresh       'repaint the form to get rid og the common dialog
    Screen.MousePointer = 11 'set the hourglass
    Set gdbCurrentDB = OpenDatabase(sDatabaseName, False, True)

    For Each tdf In gdbCurrentDB.TableDefs

        If (tdf.Attributes And &H80000002) = 0 Then
            cboRecordSource.AddItem tdf.Name
            cboRecordSource.ItemData(cboRecordSource.NewIndex) = 0
        End If

    Next tdf

    If giDataType = gnDT_ACCESS Then

        For Each qdf In gdbCurrentDB.QueryDefs
            cboRecordSource.AddItem qdf.Name
            cboRecordSource.ItemData(cboRecordSource.NewIndex) = 1
        Next qdf

    End If
    cmdOpenDB.Visible = False
    cmdBuildApp.Visible = True
    cmdBuildApp.SetFocus
    cboRecordSource.ListIndex = 0
    Screen.MousePointer = 0 'unset the hourglass


Exit Sub

OpenError:
    Screen.MousePointer = 0 'unset the hourglass
    If Err <> 32755 Then     'check for common dialog cancelled
        MsgBox Err.Description
    End If

End Sub

Private Sub cmdREG_Click()

'------ End of generate reference and form to project -----
Dim msg      As String
Dim response As Integer

    msg = " Register and you get all the power of VB DAO Engine" & vbNewLine & " without Data Control" & vbNewLine & vbNewLine & " More Create Options." & vbNewLine & " More Data control Options." & vbNewLine & " More Form control Options." & vbNewLine & vbNewLine & " For ONLY  free / (free with" & " source) Non Comerical." & vbNewLine & " For ONLY $169.95 (source included with Comerical)." & vbNewLine & vbNewLine & " YES to Print Registeration Form & Details."
    response = MsgBox(msg, 292, "Registeration")

    Select Case response
        ' Yes response.
    Case vbYes:
        PrintRegForm
        ' No response.
    Case vbNo:
    End Select

End Sub

Private Sub Form_Load()

    sSQLStatement = ""
    gsFormName = ""
    gsRecordTable = ""
    ProgressB.Value = 0
    'center it on the screen
    Me.Top = (Screen.Height - Me.Height) \ 2
    Me.Left = (Screen.Width - Me.Width) \ 2
    cboRecordSource.Clear
    lstClean.Clear
    cmdBuildApp.Visible = False
    cmdClean.Enabled = False
    Frame4.Visible = True
    Frame5.Visible = True
    '  On Error GoTo OpenError
    cmdOpenDB.Visible = True
    cmdBuildApp.Visible = False
    Screen.MousePointer = 0 'unset the hourglass
    Me.Show


Exit Sub

OpenError:
    Screen.MousePointer = 0 'unset the hourglass
    If Err <> 32755 Then     'check for common dialog cancelled
        MsgBox Err.Description
    End If

End Sub

Public Sub BuildCreateADOSub(ff As Integer)

  On Error GoTo ErrorGen

Dim cat     As ADOX.Catalog
Dim catTbl  As ADOX.Table
Dim catIdx1 As ADOX.Index
Dim Col     As ADOX.Column
Dim Fld     As ADODB.Fields
Dim ColNum  As Integer
Dim sGen    As String
Dim i       As Integer
Dim j       As Integer
Dim k       As Integer

    Screen.MousePointer = vbHourglass

    Call Connect
    'Open the Database Catalog from Actual Connection
    Set cat = New ADOX.Catalog
    cat.ActiveConnection = cnn

    Print #ff, "Private Sub CreateNewDatabase()"
    Print #ff,
    Print #ff, "On Error Goto ErrorCreateDB"
    Print #ff, "Dim Cat     As New ADOX.Catalog"
    Print #ff, "Dim Tbl(" & cat.Tables.Count - 1 & ") As ADOX.Table"
    Print #ff, "Dim Idx()   As ADOX.Index"
    Print #ff, "Dim msgErrR As integer"
    Print #ff, "Dim sCnn    As String "
    If chkCopyMDB.Value = vbChecked Then
     Print #ff, "sCnn =  ""Provider=Microsoft.Jet.OLEDB.4.0 ;Jet OLEDB:Engine Type=5;Data Source= "" &  App.Path & ""\" & StripPath(gsMDBPath) & Chr$(34) & """"
     Print #ff, "'sCnn =  ""Provider=Microsoft.Jet.OLEDB.4.0 ;Jet OLEDB:Engine Type=5;Data Source=" & sDatabaseName & Chr$(34) & ""
        Else
    Print #ff, "'sCnn =  ""Provider=Microsoft.Jet.OLEDB.4.0 ;Jet OLEDB:Engine Type=5;Data Source= "" &  App.Path & ""\" & StripPath(gsMDBPath) & Chr$(34) & """"
    Print #ff, "sCnn =  ""Provider=Microsoft.Jet.OLEDB.4.0 ;Jet OLEDB:Engine Type=5;Data Source=" & sDatabaseName & Chr$(34) & ""
End If
    Print #ff, "Cat.Create sCnn"

    'Get table names
    i = 0

    'Table Definitions
    For Each catTbl In cat.Tables

        If catTbl.Type = "TABLE" Then

            Print #ff, "  'Table Definition of " & catTbl.Name
            Print #ff, "  Set Tbl(" & Trim$(Str$(i)) & ")= New ADOX.Table"
            Print #ff, "  Tbl(" & Trim$(Str$(i)) & ").ParentCatalog = Cat"
            Print #ff, "  With Tbl(" & Trim$(Str$(i)) & ")" & vbTab
            Print #ff, "    .Name = """ & catTbl.Name & """"

            'Field Definitions
            For Each Col In catTbl.Columns

                Print #ff, "      .Columns.Append """ & Col.Name & """, " & Trim$(LoadResString(Col.Type)) & IIf(Col.DefinedSize <> 0 And Col.Type <> adBoolean, ", " & Col.DefinedSize, "")

                'Some Properties
                If Col.Properties("AutoIncrement").Value Then
                    Print #ff, "      .Columns(""" & Col.Name & """).Properties(""AutoIncrement"").Value = True"
                End If
                If Len(Col.Properties("Description").Value) > 0 Then
                    Print #ff, "      .Columns(""" & Col.Name & """).Properties(""Description"").Value = """ & Col.Properties("Description").Value & """"
                End If
                If Not Col.Properties("Nullable").Value = False Then
                    Print #ff, "      .Columns(""" & Col.Name & """).Properties(""Nullable"").Value = False"
                End If
                If Len(Col.Properties("Default").Value) > 0 Then
                    Print #ff, "      .Columns(""" & Col.Name & """).Properties(""Default"").Value = " & Col.Properties("Default").Value
                End If
                If Len(Col.Properties("Jet OLEDB:Column Validation Text").Value) > 0 Then
                    Print #ff, "      .Columns(""" & Col.Name & """).Properties(""Jet OLEDB:Column Validation Text"").Value = """ & Col.Properties("Jet OLEDB:Column Validation Text").Value & """ "
                End If
                If Len(Col.Properties("Jet OLEDB:Column Validation Rule").Value) > 0 Then
                    Print #ff, "      .Columns(""" & Col.Name & """).Properties(""Jet OLEDB:Column Validation Rule"").Value = """ & Col.Properties("Jet OLEDB:Column Validation Rule").Value & """ "
                End If
                If Not Col.Properties("Jet OLEDB:Compressed UNICODE Strings").Value Then
                    Print #ff, "      .Columns(""" & Col.Name & """).Properties(""Jet OLEDB:Compressed UNICODE Strings"").Value = " & Col.Properties("Jet OLEDB:Compressed UNICODE Strings").Value
                End If
                If Col.Properties("Jet OLEDB:Allow Zero Length").Value = True Then
                    Print #ff, "      .Columns(""" & Col.Name & """).Properties(""Jet OLEDB:Allow Zero Length"").Value = " & Col.Properties("Jet OLEDB:Allow Zero Length").Value
                End If

            Next Col

            Print #ff, "  End With"

            'Indexes
            If catTbl.Indexes.Count > 0 Then
                Print #ff, "  'Index Definitions of " & catTbl.Name
                Print #ff, "  ReDim Idx(" & Trim$(Str$(catTbl.Indexes.Count - 1)) & ")"
            End If

            j = 0

            For Each catIdx1 In catTbl.Indexes
                Print #ff, "  Set Idx(" & Trim$(Str$(j)) & ")= New ADOX.Index"
                Print #ff, "    Idx(" & Trim$(Str$(j)) & ").Name = """ & catIdx1.Name & """"
                If catIdx1.PrimaryKey Then
                    Print #ff, "    Idx(" & Trim$(Str$(j)) & ").PrimaryKey = True"
                End If
                If catIdx1.IndexNulls <> adIndexNullsDisallow Then
                    Select Case catIdx1.IndexNulls
                    Case Is = 0
                        Print #ff, "    Idx(" & Trim$(Str$(j)) & ").IndexNulls = adIndexNullsAllow"
                    Case Is = 2
                        Print #ff, "    Idx(" & Trim$(Str$(j)) & ").IndexNulls = adIndexNullsIgnore"
                    Case Is = 4
                        Print #ff, "    Idx(" & Trim$(Str$(j)) & ").IndexNulls = adIndexNullsIgnoreAny"
                    End Select
                End If
                If catIdx1.Unique = True Then
                    Print #ff, "    Idx(" & Trim$(Str$(j)) & ").Unique = True"
                End If

                If catIdx1.Columns.Count = 1 Then
                    'Single Column Index
                    Print #ff, "      Idx(" & Trim$(Str$(j)) & ").Columns.Append """ & catIdx1.Columns(0).Name & """"
                    If catIdx1.Columns.Item(0).SortOrder = adSortDescending Then
                        Print #ff, "          Idx(" & Trim$(Str$(j)) & ").Columns(""" & catIdx1.Columns(0).Name & """).SortOrder = adSortDescending"
                    End If
                ElseIf catIdx1.Columns.Count > 1 Then
                    'MultiColumn Index
                    For k = 0 To catIdx1.Columns.Count - 1
                        Print #ff, "      Idx(" & Trim$(Str$(j)) & ").Columns.Append """ & catIdx1.Columns(k).Name & """"
                        If catIdx1.Columns.Item(k).SortOrder = adSortDescending Then
                            Print #ff, "          Idx(" & Trim$(Str$(j)) & ").Columns(""" & catIdx1.Columns(k).Name & """).SortOrder = adSortDescending"
                        End If
                    Next k
                End If

                j = j + 1

            Next catIdx1

            If j > 1 Then
                Print #ff, "  For i = 0 to UBound(Idx)"
                Print #ff, "    Tbl(" & Trim$(Str$(i)) & ").Indexes.Append Idx(i)"
                Print #ff, "  Next i"
            ElseIf j = 1 Then
                Print #ff, "  Tbl(" & Trim$(Str$(i)) & ").Indexes.Append Idx(0)"
            End If

            Print #ff, "  Cat.Tables.Append Tbl(" & Trim$(Str$(i)) & ")"

        End If
        i = i + 1

    Next catTbl

    'Error code
    Print #ff, "  Set Cat = Nothing"
    Print #ff, "  Exit Sub"
    Print #ff, "  ErrorCreateDB:"
    Print #ff, "    msgErrR = MsgBox(""Error No. "" & Err & "" ""  & Error, vbCritical+vbAbortRetryIgnore, ""Code Gen Error""" & ")"
    Print #ff, "    Select Case msgErrR"
    Print #ff, "      Case Is = vbAbort"
    Print #ff, "      If Not (Cat is Nothing) Then"
    Print #ff, "        Set Cat = Nothing"
    Print #ff, "      Endif"
    Print #ff, "      Exit Sub"
    Print #ff, "     Case Is = vbRetry"
    Print #ff, "       Resume Next"
    Print #ff, "     Case Is = vbIgnore"
    Print #ff, "       Resume"
    Print #ff, "    End Select"
    Print #ff, "End Sub"

    Screen.MousePointer = vbDefault

Exit Sub
ErrorGen:

    MsgBox "Error No. " & Err & Error, vbCritical, "Error"
    ErrGen = True
    Screen.MousePointer = vbDefault

End Sub

Private Sub BuildFieldToText(fh As Integer)

Dim fieldtype As String
Dim sTmp      As String
Dim nNumFlds  As Integer
Dim i         As Integer

    On Error Resume Next
    nNumFlds = lstFields.ListCount
    ' here we need to see what mode were in and set defaults
    ' if field names have spaces or "-" then put brackets around name.
    ' create the data  table

    For i = 0 To nNumFlds - 1
        sTmp = lstFields.List(i)

        Select Case currentRS.Fields(i).Type
        Case 1
            fieldtype = "dbBoolean" 'Yes / No

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "        txtField(" & i & ") = cstr(!" & AddBrackets(sTmp) & ")"
            Else
                Print #fh, "        txtField(" & i & ") = cstr(!" & sTmp & ")"
            End If

        Case 2
            fieldtype = "dbByte" ' Byte

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "        txtField(" & i & ") = cstr(!" & AddBrackets(sTmp) & ")"
            Else
                Print #fh, "        txtField(" & i & ") = cstr(!" & sTmp & ")"
            End If

        Case 3
            fieldtype = "dbInteger" '   Integer

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "        txtField(" & i & ") = cstr(!" & AddBrackets(sTmp) & ")"
            Else
                Print #fh, "        txtField(" & i & ") = cstr(!" & sTmp & ")"
            End If

        Case 4
            fieldtype = "dbLong" ' Long

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "        txtField(" & i & ") = cstr(!" & AddBrackets(sTmp) & ")"
            Else
                Print #fh, "        txtField(" & i & ") = cstr(!" & sTmp & ")"
            End If

        Case 5
            fieldtype = "dbCurrency" '  Currency

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "        txtField(" & i & ") = cstr(!" & AddBrackets(sTmp) & ")"
            Else
                Print #fh, "        txtField(" & i & ") = cstr(!" & sTmp & ")"
            End If

        Case 6
            fieldtype = "dbSingle"  '  Single

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "        txtField(" & i & ") = cstr(!" & AddBrackets(sTmp) & ")"
            Else
                Print #fh, "        txtField(" & i & ") = cstr(!" & sTmp & ")"
            End If

        Case 7
            fieldtype = "dbDouble" '   Double

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "        txtField(" & i & ") = cstr(!" & AddBrackets(sTmp) & ")"
            Else
                Print #fh, "        txtField(" & i & ") = cstr(!" & sTmp & ")"
            End If

        Case 8
            fieldtype = "dbDate" 'Date / Time

        Case 10
            fieldtype = "dbText" ' Text

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "        txtField(" & i & ") = !" & AddBrackets(sTmp)
            Else
                Print #fh, "        txtField(" & i & ") = !" & sTmp
            End If

        Case 11

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "        txtField(" & i & ") = cstr(!" & AddBrackets(sTmp) & ")"
            Else
                Print #fh, "        txtField(" & i & ") = cstr(!" & sTmp & ")"
            End If

        Case 12
            fieldtype = "dbMemo" 'Memo

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "        txtField(" & i & ") = !" & AddBrackets(sTmp)
            Else
                Print #fh, "        txtField(" & i & ") = !" & sTmp
            End If

        End Select
    Next i

End Sub

Public Sub BuildCreateDaoSub(fh As Integer)


Dim sTmp      As String
Dim nNumFlds  As Integer
Dim i         As Integer
Dim j         As Integer
Dim idxobj    As Index
Dim tdf       As TableDef
Dim fieldtype As String

    On Error Resume Next
    nNumFlds = lstFields.ListCount
    'create all names and subs for creating new database

    If chkCreateFilemod.Value = vbChecked Then
        Print #fh, "Private Sub Create" & Left$(RemoveSpace(currentRS.Name), 37)

        Print #fh, "      Dim tdf As TableDef"
        Print #fh, "      Dim fld As Field"
        Print #fh, "      Dim idx As Index"
        Print #fh, ""

        If chkComments.Value = vbChecked Then
            Print #fh, "    ' set database for creation of tables"
        End If

        Print #fh, "    Set tdf = gDB.CreateTableDef(" & Chr$(34) & currentRS.Name & Chr$(34) & ")"
        Print #fh, ""
        'create the data  table

        For i = 0 To nNumFlds - 1
            sTmp = lstFields.List(i)

            Select Case currentRS.Fields(i).Type
            Case 1
                fieldtype = "dbBoolean" 'Yes / No
                Print #fh, "    Set fld = tdf.CreateField(" & Chr$(34) & sTmp & Chr$(34) & "," & fieldtype & ")"

            Case 2
                fieldtype = "dbByte" ' Byte
                Print #fh, "    Set fld = tdf.CreateField(" & Chr$(34) & sTmp & Chr$(34) & "," & fieldtype & ")"

            Case 3
                fieldtype = "dbInteger" '   Integer
                Print #fh, "    Set fld = tdf.CreateField(" & Chr$(34) & sTmp & Chr$(34) & "," & fieldtype & ")"

            Case 4
                fieldtype = "dbLong" ' Longdb
                Print #fh, "    Set fld = tdf.CreateField(" & Chr$(34) & sTmp & Chr$(34) & "," & fieldtype & ")"

            Case 5
                fieldtype = "dbCurrency" '  Currency
                Print #fh, "    Set fld = tdf.CreateField(" & Chr$(34) & sTmp & Chr$(34) & "," & fieldtype & ")"

            Case 6
                fieldtype = "dbSingle"  '  Single
                Print #fh, "    Set fld = tdf.CreateField(" & Chr$(34) & sTmp & Chr$(34) & "," & fieldtype & ")"

            Case 7
                fieldtype = "dbDouble" '   Double
                Print #fh, "    Set fld = tdf.CreateField(" & Chr$(34) & sTmp & Chr$(34) & "," & fieldtype & ")"

            Case 8
                fieldtype = "dbDate" 'Date / Time
                Print #fh, "    Set fld = tdf.CreateField(" & Chr$(34) & sTmp & Chr$(34) & "," & fieldtype & ")"

            Case 10
                fieldtype = "dbText" ' Text
                Print #fh, "    Set fld = tdf.CreateField(" & Chr$(34) & sTmp & Chr$(34) & "," & fieldtype & "," & currentRS.Fields(sTmp).Size & ")"
                Print #fh, "    fld.AllowZeroLength = True"

            Case 11
                fieldtype = "dbLongBinary"   ' Long Binary (OLE Object)
                Print #fh, "    Set fld = tdf.CreateField(" & Chr$(34) & sTmp & Chr$(34) & "," & fieldtype & ")"

            Case 12
                fieldtype = "dbMemo" 'Memo
                Print #fh, "    Set fld = tdf.CreateField(" & Chr$(34) & sTmp & Chr$(34) & "," & fieldtype & "," & currentRS.Fields(sTmp).Size & ")"
                Print #fh, "    fld.AllowZeroLength = True"
            End Select

            Print #fh, "   tdf.Fields.Append fld"
            Print #fh, ""
        Next i

        Print #fh, "    gDB.TableDefs.Append tdf"
        'create the indexes

        For i = 0 To gdbCurrentDB.TableDefs(currentRS.Name).Indexes.Count - 1
            Set idxobj = gdbCurrentDB.TableDefs(currentRS.Name).Indexes(i)
            ' Create new Index object.

            With idxobj
                Print #fh, "    Set idx = tdf.CreateIndex(" & Chr$(34) & .Name & Chr$(34) & ")"

                For j = 0 To .Fields.Count - 1
                    Print #fh, "    Set fld = idx.CreateField(" & Chr$(34) & .Fields(j).Name & Chr$(34) & ")"
                Next j

                Print #fh, "   idx.Unique = " & .Unique
                Print #fh, "   idx.Primary = " & .Primary
                Print #fh, "   idx.Required = " & .Required
                Print #fh, "   idx.IgnoreNulls = " & .IgnoreNulls
                Print #fh, "   idx.Clustered = " & .Clustered
                Print #fh, ""
            End With

            Print #fh, "   idx.Fields.Append fld"
            Print #fh, ""
        Next i

        If j > 0 Then 'if no indexes don't print
            Print #fh, "   tdf.Indexes.Append idx"
        End If

        Print #fh, "End Sub"
    End If

    '********** end *******************
Exit Sub

WCErr:
    MsgBox Error$

End Sub

Private Sub BuildMDIProjectCode1(fh As Integer)

Dim nNumFlds As Integer

    On Error GoTo WCErr
    nNumFlds = lstFields.ListCount
    'create and open the file
    Print #fh, "VERSION 5.00"
    Print #fh, "Object = ""{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0""; ""MSDATGRD.OCX"""
    Print #fh, "Object = ""{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0""; ""COMCTL32.OCX"""
    Print #fh, "Begin VB.MDIForm frmMDI"
    Print #fh, "   BackColor       =   &H8000000C&"

    If optDAO.Value = True Then
        Print #fh, "   Caption         =   " & Chr$(34) & "DAO Version 1.00" & Chr$(34)
    Else
        Print #fh, "   Caption         =   " & Chr$(34) & "ADO Version 1.00" & Chr$(34)
    End If

    Print #fh, "   ClientHeight    =   6075"
    Print #fh, "   ClientLeft      =   1785"
    Print #fh, "   ClientTop       =   1260"
    Print #fh, "   ClientWidth     =   6690"
    Print #fh, "   Height          =   6765"
    Print #fh, "   Left            =   1725"
    Print #fh, "   LinkTopic       =   " & Chr$(34) & "MDIForm1" & Chr$(34)
    Print #fh, "   Top             =   630"
    Print #fh, "   Width           =   6810"
    Print #fh, "   WindowState     =   2  'Maximized"
    Print #fh, "   Begin VB.PictureBox Picture1"
    Print #fh, "      Align           =   2  'Align Bottom"
    Print #fh, "      ClipControls    =   0   'False"
    Print #fh, "      Height          =   495"
    Print #fh, "      Left            =   0"
    Print #fh, "      ScaleHeight     =   435"
    Print #fh, "      ScaleWidth      =   6630"
    Print #fh, "      TabIndex        =   0"
    Print #fh, "      Top             =   5580"
    Print #fh, "      Visible         =   1   'True"
    Print #fh, "      Width           =   6690"
    Print #fh, "      Begin VB.Label MsgBar"
    Print #fh, "         BackColor       =   &H00FFFFFF&"
    Print #fh, "         BorderStyle     =   1  'Fixed Single"
    Print #fh, "         Caption         =   " & Chr$(34) & "MsgBar" & Chr$(34)
    Print #fh, "         ForeColor       =   &H00FF0000&"
    Print #fh, "         Height          =   315"
    Print #fh, "         Left            =   120"
    Print #fh, "         TabIndex        =   1"
    Print #fh, "         Top             =   60"
    Print #fh, "         Width           =   4275"
    Print #fh, "      End"
    Print #fh, "   End"
    Print #fh, "   Begin MSComDlg.CommonDialog ComDlg"
    Print #fh, "      Left            =   45"
    Print #fh, "      Top             =   5550"
    Print #fh, "      _Version        =   65536"
    Print #fh, "      _ExtentX        =   847"
    Print #fh, "      _ExtentY        =   847"
    Print #fh, "      _StockProps     =   0"
    Print #fh, "   End"
    Print #fh, "   Begin VB.Menu mnuFile"
    Print #fh, "      Caption         =   " & Chr$(34) & "&File" & Chr$(34)
    Print #fh, "      Begin VB.Menu mnuFilePrint"
    Print #fh, "         Caption         =   " & Chr$(34) & "&Print" & Chr$(34)
    Print #fh, "      End"
    Print #fh, "      Begin VB.Menu mnuFileSetup"
    Print #fh, "         Caption         =   " & Chr$(34) & "Printer &Setup" & Chr$(34)
    Print #fh, "      End"
    Print #fh, "      Begin VB.Menu mnuFileSep1"
    Print #fh, "         Caption         =   " & Chr$(34) & "-" & Chr$(34)
    Print #fh, "      End"
    Print #fh, "      Begin VB.Menu mnuFileExit"
    Print #fh, "         Caption         =   " & Chr$(34) & "E&xit" & Chr$(34)
    Print #fh, "      End"
    Print #fh, "   End"
    Print #fh, "   Begin VB.Menu mnuEdit"
    Print #fh, "      Caption         =   " & Chr$(34) & "&Edit" & Chr$(34)

Exit Sub

WCErr:
    MsgBox Error$

End Sub

Private Sub BuildMDIProjectCode2(fh As Integer)

    Print #fh, "   Begin VB.Menu mnuEdit" & Left$(RemoveSpace(currentRS.Name), 37)
    Print #fh, "         Caption = " & Chr$(34) & Left$(RemoveSpace(currentRS.Name), 37) & Chr$(34)
    Print #fh, "   End"

End Sub

Private Sub BuildMDIProjectCode3(fh As Integer)

'end of edit menu

    Print #fh, "   End"
    Print #fh, "   Begin VB.Menu mnuWindow"
    Print #fh, "      Caption         =   " & Chr$(34) & "&Window" & Chr$(34)
    Print #fh, "      Begin VB.Menu mnuWTile"
    Print #fh, "         Caption         =   " & Chr$(34) & "&Tile" & Chr$(34)
    Print #fh, "      End"
    Print #fh, "      Begin VB.Menu mnuWCascade"
    Print #fh, "         Caption         =   " & Chr$(34) & "&Cascade" & Chr$(34)
    Print #fh, "      End"
    Print #fh, "      Begin VB.Menu mnuArrange"
    Print #fh, "         Caption         =   " & Chr$(34) & "&Arrange Icons" & Chr$(34)
    Print #fh, "      End"
    Print #fh, "      Begin VB.Menu mnuWDash1"
    Print #fh, "         Caption         =   " & Chr$(34) & "-" & Chr$(34)
    Print #fh, "      End"
    Print #fh, "      Begin VB.Menu mnuWView"
    Print #fh, "         Caption         =   " & Chr$(34) & "Bring to &view" & Chr$(34)
    Print #fh, "         WindowList      =   -1  'True"
    Print #fh, "      End"
    Print #fh, "   End"
    Print #fh, "   Begin VB.Menu mnuHelp"
    Print #fh, "      Caption         =   " & Chr$(34) & "&Help" & Chr$(34)
    Print #fh, "      Begin VB.Menu mnuAbout"
    Print #fh, "         Caption         =   " & Chr$(34) & "&About" & Chr$(34)
    Print #fh, "      End"
    Print #fh, "   End"
    Print #fh, "End"
    Print #fh, "Attribute VB_Name = " & Chr$(34) & "frmMDI" & Chr$(34)
    Print #fh, "Attribute VB_Creatable = False"
    Print #fh, "Attribute VB_Exposed = False"
    Print #fh, "Option Explicit"
    Print #fh, ""
    Print #fh, "Private Sub MDIForm_Load()"

    Print #fh, "Show"
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub MDIForm_Unload(Cancel As Integer)"

    Print #fh, ""
    Print #fh, "    Dim nRV As Integer"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' When the user selects the File Exit option"
        Print #fh, "    ' display a message box to check whether or not the user really"
        Print #fh, "    ' wants to quit."
    End If

    Print #fh, ""
    Print #fh, "    nRV = MsgBox(" & Chr$(34) & "Do you want to exit?" & Chr$(34) & ", vbQuestion +" & " vbYesNo, " & Chr$(34) & "Exit" & Chr$(34) & ")"
    Print #fh, "    If nRV = vbNo Then "
    Print #fh, "    Cancel = True"
    Print #fh, "    Else"
    Print #fh, "    gDB.Close"
    Print #fh, "    End   "
    Print #fh, "    End IF"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub mnuArrange_Click()"

    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' Arrange any icon windows"
    End If

    Print #fh, "    frmMDI.Arrange vbArrangeIcons"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub mnuFileExit_Click()"

    Print #fh, "    Unload frmMDI"
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub mnuFilePrint_Click()"

    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub mnuFileSetup_Click()"

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' Display the print setup dialog"
    End If

    Print #fh, "    ComDlg.Flags = cdlPDPrintSetup"
    Print #fh, "    ComDlg.ShowPrinter"
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, ""
    Print #fh, "Private Sub mnuEdit_Click()"

    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub mnuFile_Click()"

    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, ""
    Print #fh, "Private Sub mnuWCascade_Click()"

    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' Cascade the windows on display within the MDI form"
    End If

    Print #fh, "    frmMDI.Arrange vbCascade"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub mnuAbout_Click()"

    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' Display the about About dialog box"
    End If

    Print #fh, "    Load frmAbout"
    Print #fh, "    frmAbout.Show 1"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub mnuWTile_Click()"

    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' Tile the windows on MDI form"
    End If

    Print #fh, "    frmMDI.Arrange vbTileHorizontal"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub mnuWView_Click()"

    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub Picture1_Click()"

    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, ""
    '********** end of code *******************

End Sub

Private Sub BuildMDIProjectCode4(fh As Integer)

' add all edit form names to menu

    Print #fh, "Private Sub mnuEdit" & Left$(RemoveSpace(currentRS.Name), 37) & "_Click()"

    Print #fh, "        Load frm" & Left$(RemoveSpace(currentRS.Name), 37)
    Print #fh, "        "
    Print #fh, "End sub"

End Sub

Private Sub BuildNoDataCode(fh As Integer)

Dim sTmp     As String
Dim nNumFlds As Integer
Dim i        As Integer

     On Error GoTo WCErr
    nNumFlds = lstFields.ListCount
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "  'status of record flags"
    End If

    Print #fh, "  Dim bAddingRecord As Boolean"
    Print #fh, "  Dim bEditingRecord As Boolean"
    Print #fh, ""

    Print #fh, "  Dim bCurrentRecord As Boolean"
    Print #fh, ""

    Print #fh, "  Dim gnOwnerID As String"
    If optDAO.Value = True Then Print #fh, "  Dim dataRS As Recordset"
    Print #fh, "  Dim tmpBM As Variant"
    Print #fh, ""

    Print #fh, "'   Program generated with Pro-APP Creator"
    Print #fh, "'   By: Raymond E Dixon"
    Print #fh, "'   tvmancet@hotmail.com "
    Print #fh, ""
    Print #fh, ""
    Print #fh, "Private Sub cmdData_MouseMove(Index As Integer,Button As Integer, Shift As Integer," & " x As Single, y As Single)"

    If chkComments.Value = vbChecked Then
        Print #fh, "' display message depending on where mouse is located"
    End If

    Print #fh, ""
    Print #fh, ""
    Print #fh, " On error resume next"
    Print #fh, "    Select Case Index"
    Print #fh, "        Case 0"
    Print #fh, "        dispstatus me," & Chr$(34) & "Exit" & Chr$(34)
    Print #fh, "        Case 1"
    Print #fh, "        dispstatus me," & Chr$(34) & "Print" & Chr$(34)
    Print #fh, "        Case 2"
    Print #fh, "        dispstatus me," & Chr$(34) & "Search" & Chr$(34)
    Print #fh, "        Case 3"
    Print #fh, "        dispstatus me," & Chr$(34) & "Add New Record" & Chr$(34)
    Print #fh, "        Case 4"
    Print #fh, "        dispstatus me," & Chr$(34) & "Delete Record" & Chr$(34)
    Print #fh, "        Case 5"
    Print #fh, "        dispstatus me," & Chr$(34) & "Move to First Record" & Chr$(34)
    Print #fh, "        Case 6"
    Print #fh, "        dispstatus me," & Chr$(34) & "Move to Previous Record" & Chr$(34)
    Print #fh, "        Case 7"
    Print #fh, "        dispstatus me," & Chr$(34) & "Move to Next Record" & Chr$(34)
    Print #fh, "        Case 8"
    Print #fh, "        dispstatus me," & Chr$(34) & "Move to Last Record" & Chr$(34)
    Print #fh, "        Case 9"
    Print #fh, "        dispstatus me," & Chr$(34) & "Cancel Edit" & Chr$(34)
    Print #fh, "        Case 10"
    Print #fh, "        dispstatus me," & Chr$(34) & "Update Record" & Chr$(34)
    Print #fh, "        Case Else"
    Print #fh, "        dispstatus me," & Chr$(34) & "" & Chr$(34)
    Print #fh, "    End Select"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub Form_Load()"

    Print #fh, "    on error resume next"
    Print #fh, "    Dim dbSQL As String"
    Print #fh, "    Dim nIndex As Integer"
    Print #fh, "    Dim msg As String"
    Print #fh, ""

    Print #fh, "    Screen.MousePointer = vbHourglass"
    Print #fh, "    Err.Clear"
    Print #fh, ""
    Print #fh, "      CenterMe Me 'form"
    Print #fh, ""
    If optDAO.Value = True Then
        If cboRecordSource.ItemData(giCurrentTableNumber) = 1 Then
            Print #fh, "    dbSQL = """ & RemoveCRLF(sSQLStatement)
            Print #fh, "    Set dataRS = gDB.OpenRecordset(dbSQL,dbOpenDynaset)"
        Else
            Print #fh, "    Set dataRS = gDB.OpenRecordset(""" & currentRS.Name & """,dbOpenDynaset)"
        End If
    Else
        Print #fh, "     dataRS.CursorLocation = adUseClient"
        Print #fh, "    dataRS.Open """ & currentRS.Name & """, gDB, adOpenDynamic, adLockOptimistic"
    End If
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' If there are no records in the recordset"
        Print #fh, "    ' disable all buttons except add"
    End If

    Print #fh, ""
    Print #fh, "    If dataRS.BOF And dataRS.EOF Then"
    Print #fh, "        For nIndex = gcExit to gcLast"
    Print #fh, "            cmdData(nIndex).Enabled = False"
    Print #fh, "        Next"
    Print #fh, "           cmdData(gcAdd).Enabled = True"
    Print #fh, "    Else"
    Print #fh, "        Refresh_Screen"
    Print #fh, "    End If"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' Make OK and Cancel buttons invisible"
    End If

    Print #fh, "      cmdData(gcCancel).Visible = False"
    Print #fh, "      cmdData(gcUpdate).Visible = False"
    Print #fh, ""
    Print #fh, "      Me.Show"
    Print #fh, "      "
    Print #fh, "      Screen.MousePointer = vbDefault"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub cmdData_Click(Index As Integer)"

    If chkComments.Value = vbChecked Then
        Print #fh, "'  handle all commands from here"
    End If

    Print #fh, "        Dim nIndex As Integer"
    Print #fh, "        Dim nRV As Integer"
    Print #fh, "        Dim msg As string"
    Print #fh, "        Dim i As Integer"
    Print #fh, "        Dim sBookMark As string"
    Print #fh, "        Dim sTmp As String"
    Print #fh, ""

    Print #fh, "        on error resume next"
    Print #fh, ""
    Print #fh, "        Select Case Index"
    Print #fh, "               Case 0 'exit"
    Print #fh, "               Unload frm" & Left$(RemoveSpace(currentRS.Name), 37)
    Print #fh, "               Case 1 ' Print"
    Print #fh, "               PrintAll" & RemoveSpace(gsRecordTable)
    Print #fh, "               Case 2 'search"
    Print #fh, "               On Error GoTo FindErr"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "               'load the field names into the search form"
        Print #fh, "               'we add all fields to search form "
    End If

    Print #fh, "               If frmSearch.lstFields.ListCount = 0 Then"
    Print #fh, "                 With dataRS"
    Print #fh, "                   For i = 0 To dataRS.Fields.Count - 1"
    Print #fh, "                   frmSearch.lstFields.AddItem .fields(i).Name"
    Print #fh, "                   Next"
    Print #fh, "                 End With "
    Print #fh, "               End If"
    Print #fh, "               FindStart:"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "               'reset the flags"
    End If
    
        If optADO.Value = True Then
                  Print #fh, "                             'set to original record"
                  Print #fh, "                If dataRS.BOF Or dataRS.EOF Then"
                  Print #fh, "                  dataRS.MoveFirst"
                  Print #fh, "               Else"
                  Print #fh, "                sBookMark = dataRS.Bookmark"
                  Print #fh, "               End If"
         End If
    
    Print #fh, "               gbFindFailed = False"
    Print #fh, "               gbNotFound = False"
    Print #fh, "               frmSearch.Caption = " & Chr$(34) & "Find " & Chr$(34) & "& Me.Caption"
    Print #fh, "               frmSearch.Show vbModal"
    Print #fh, "               dispstatus me,""Searching for New Record"""
    Print #fh, "               If gbFindFailed = True Then   'find cancelled"
    Print #fh, "                GoTo AfterWhile"
    Print #fh, "               End If"
    Print #fh, "               Screen.MousePointer = vbHourglass"
    Print #fh, "               i = frmSearch.lstFields.ListIndex"
    Print #fh, "               sBookMark = dataRS.Bookmark"

    If chkComments.Value = vbChecked Then
        Print #fh, "               'search for the record"
    End If
    If optDAO.Value = True Then

        Print #fh, "               If dataRS.Fields(i).Type = dbText Or dataRS.Fields(i).Type = dbMemo" & " Then"
        Print #fh, "                 sTmp = AddBrackets((dataRS.Fields(i).Name)) & "" "" & gsFindOp & """ & " '"" & gsFindExpr & ""'"""
        Print #fh, "               Else"
        Print #fh, "                 sTmp = AddBrackets((dataRS.Fields(i).Name)) + gsFindOp + gsFindExpr"
        Print #fh, "               End If"

        Print #fh, "               Select Case gnFindType"
        Print #fh, "                      Case 0"
        Print #fh, "                           dataRS.FindFirst sTmp"
        Print #fh, "                      Case 1"
        Print #fh, "                           dataRS.FindNext sTmp"
        Print #fh, "                      Case 2"
        Print #fh, "                           dataRS.FindPrevious sTmp"
        Print #fh, "                      Case 3"
        Print #fh, "                           dataRS.FindLast sTmp"
        Print #fh, "               End Select"
        Print #fh, "               gbNotFound = dataRS.NoMatch"
        Print #fh, "              AfterWhile:"
    Else 'ado
        Print #fh, "                 sTmp = AddBrackets((dataRS.Fields(i).Name)) & "" "" & gsFindOp & """ & " '"" & gsFindExpr & ""'"""
        Print #fh, "                                          dataRS.MoveFirst"

        Print #fh, "               Select Case gnFindType"
        Print #fh, "                      Case 0"
        Print #fh, "                           dataRS.Find sTmp"
        Print #fh, "                      Case 1"
        Print #fh, "                           dataRS.Find sTmp,adSearchForward,sBookMark"
        Print #fh, "                      Case 2"
        Print #fh, "                           dataRS.Find sTmp,adSearchBackward,sBookMark"
        Print #fh, "                      Case 3"
        Print #fh, "                          dataRS.movelast"
        Print #fh, "                           dataRS.Find sTmp, adSearchBackward"
        Print #fh, "               End Select"
        Print #fh, "  If dataRS.BOF Or dataRS.EOF Then gbNotFound = True"

        Print #fh, "              AfterWhile:"
    End If

    Print #fh, "              Screen.MousePointer = vbDefault"
    Print #fh, "              If gbFindFailed = True Then  "

    If chkComments.Value = vbChecked Then
        Print #fh, "              'set to original record"
    End If

   If optDAO.Value = True Then Print #fh, "                 dataRS.Bookmark = sBookMark"
    
    Print #fh, "                ElseIf gbNotFound Then"
    Print #fh, "                 Beep"
    Print #fh, "                 MsgBox ""Record Not Found"", 48"
    Print #fh, "                 dataRS.Bookmark = sBookMark"
    Print #fh, "                 GoTo FindStart"
    Print #fh, "              End If"
    Print #fh, "                 Refresh_Screen"
    Print #fh, "                 Exit Sub"
    Print #fh, "              Case 3 ' Add a new record"

    If chkComments.Value = vbChecked Then
        Print #fh, "                ' Clear all the text boxes"
        Print #fh, "                ' Make the OK and Cancel buttons visible"
    End If

    Print #fh, "                Clear_TextControls"
    Print #fh, "                cmdData(gcCancel).Visible = True"
    Print #fh, "                cmdData(gcUpdate).Visible = True"

    If chkComments.Value = vbChecked Then
        Print #fh, "                ' Disable all buttons on the toolbar"
    End If

    Print #fh, "                For nIndex = gcExit to gcLast"
    Print #fh, "                    cmdData(nIndex).Enabled = False"
    Print #fh, "                Next"

    If chkComments.Value = vbChecked Then
        Print #fh, "                ' Set a flag to indicate add"
    End If

    Print #fh, "                bAddingRecord = True"
    Print #fh, "                txtField(0).SetFocus"
    Print #fh, "            Case 4 ' Delete existing record"
    Print #fh, "                " ' If there is a record on show"
    Print #fh, "                If Not dataRS.EOF And Not dataRS.BOF Then"

    If chkComments.Value = vbChecked Then
        Print #fh, "                    ' Show the delete msgbox"
    End If

    Print #fh, "                    nRV = MsgBox(" & Chr$(34) & "Are you sure you want to Delete" & " record?" & Chr$(34) & ", vbYesNo + vbQuestion, " & Chr$(34) & "Delete" & Chr$(34) & ")"

    If chkComments.Value = vbChecked Then
        Print #fh, "                    ' If the user said yes then delete it"
    End If

    Print #fh, "                    If nRV = vbYes Then"

    If chkComments.Value = vbChecked Then
        Print #fh, "                        ' Delete the record"
    End If

    Print #fh, "                        dataRS.Delete"
    Print #fh, "                        If Err.Number <> 0 Then errHandler"

    If chkComments.Value = vbChecked Then
        Print #fh, "                        '  Update the display with the new record"
        Print #fh, "                        ' Enable error handling for move to the first record"
    End If

    Print #fh, ""
    Print #fh, "                        On Error Resume Next"
    Print #fh, "                        dataRS.MoveFirst"
    Print #fh, "                        On Error GoTo 0"
    Print #fh, "                        If dataRS.BOF And dataRS.EOF Then"

    If chkComments.Value = vbChecked Then
        Print #fh, "                            ' Disable all command buttons"
    End If

    Print #fh, "                            For nIndex = gcExit to gcLast"
    Print #fh, "                                cmdData(nIndex).Enabled = False"
    Print #fh, "                            Next"

    If chkComments.Value = vbChecked Then
        Print #fh, "                            'enable button 4 (add)"
    End If

    Print #fh, "                            cmdData(gcAdd).Enabled = True"

    If chkComments.Value = vbChecked Then
        Print #fh, "                            ' Clear the data area"
    End If

    Print #fh, "                            Clear_TextControls"
    Print #fh, "                        Else"

    If chkComments.Value = vbChecked Then
        Print #fh, "                            ' update screen after delete."
    End If

    Print #fh, "                            Refresh_Screen"
    Print #fh, "                        End If"
    Print #fh, "                    End If"
    Print #fh, "                End If"
    Print #fh, "            Case 5   ' Move to first record"
    Print #fh, "                On Error Resume Next"
    Print #fh, "                dataRS.MoveFirst"
    Print #fh, "                On Error GoTo 0"

    If chkComments.Value = vbChecked Then
        Print #fh, "                ' If there are no records, disable all commands except Addnew"
    End If

    Print #fh, "                If dataRS.BOF And dataRS.EOF Then"
    Print #fh, "                    For nIndex = gcExit to gcLast"
    Print #fh, "                        cmdData(nIndex).Enabled = False"
    Print #fh, "                    Next"
    Print #fh, "                    cmdData(gcAdd).Enabled = True"
    Print #fh, "                    Clear_TextControls"
    Print #fh, "                Else"

    If chkComments.Value = vbChecked Then
        Print #fh, "                ' disable first and previous"
    End If

    Print #fh, "                    Refresh_Screen"
    Print #fh, "                    cmdData(gcFirst).Enabled = False"
    Print #fh, "                    cmdData(gcPrevious).Enabled = False"
    Print #fh, "                End If"
    Print #fh, "            Case 6  ' Move to previous record"
    Print #fh, "                On Error Resume Next"
    Print #fh, "                dataRS.MovePrevious"
    Print #fh, "                On Error GoTo 0"

    If chkComments.Value = vbChecked Then
        Print #fh, "                ' at first record"
        Print #fh, "                ' disable if no records in the database"
    End If

    Print #fh, "                If dataRS.BOF Then"
    Print #fh, "                    dataRS.MoveFirst"
    Print #fh, "                    cmdData(gcFirst).Enabled = False"
    Print #fh, "                    cmdData(gcPrevious).Enabled = False"
    Print #fh, "                Else"
    Print #fh, "                    Refresh_Screen"
    Print #fh, "                End If"
    Print #fh, "            Case 7  ' Move to next record"
    Print #fh, ""
    Print #fh, "                On Error Resume Next"
    Print #fh, "                dataRS.MoveNext"
    Print #fh, "                On Error GoTo 0"

    If chkComments.Value = vbChecked Then
        Print #fh, "                ' at last record"
    End If

    Print #fh, "                If dataRS.EOF Then"
    Print #fh, "                    dataRS.MoveLast"
    Print #fh, "                    cmdData(gcNext).Enabled = False"
    Print #fh, "                    cmdData(gcLast).Enabled = False"
    Print #fh, "                Else"
    Print #fh, "                    Refresh_Screen"
    Print #fh, "                End If"
    Print #fh, "            Case 8  ' Move to last record"
    Print #fh, "                On Error Resume Next"
    Print #fh, "                dataRS.MoveLast"
    Print #fh, "                On Error GoTo 0"

    If chkComments.Value = vbChecked Then
        Print #fh, "                ' If no records, disable all except Add"
    End If

    Print #fh, "                If dataRS.BOF And dataRS.EOF Then"
    Print #fh, "                    For nIndex = gcExit to gcLast"
    Print #fh, "                        cmdData(nIndex).Enabled = False"
    Print #fh, "                    Next"
    Print #fh, "                    cmdData(gcAdd).Enabled = True"
    Print #fh, "                    Clear_TextControls"
    Print #fh, "                Else"

    If chkComments.Value = vbChecked Then
        Print #fh, "                ' disable first and previous"
    End If

    Print #fh, "                    Refresh_Screen"
    Print #fh, "                    cmdData(gcNext).Enabled = False"
    Print #fh, "                    cmdData(gcLast).Enabled = False"
    Print #fh, "                End If"
    Print #fh, "            Case 9"
    Print #fh, "                bAddingRecord = False"

    If chkComments.Value = vbChecked Then
        Print #fh, "                ' See if have a valid record."
        Print #fh, "                ' If not, then clear the display"
    End If

    Print #fh, "                If dataRS.BOF And dataRS.EOF Then"
    Print #fh, "                   Clear_TextControls"
    Print #fh, "                   cmdData(gcCancel).Visible = False"
    Print #fh, "                   cmdData(gcUpdate).Visible = False"
    Print #fh, "                   "
    Print #fh, "                    For nIndex = gcExit to gcLast"
    Print #fh, "                        cmdData(nIndex).Enabled = False"
    Print #fh, "                    Next"
    Print #fh, "                   cmdData(gcAdd).Enabled = True"
    Print #fh, "                Else"
    Print #fh, "                   cmdData(gcCancel).Visible = False"
    Print #fh, "                   cmdData(gcUpdate).Visible = False"
    Print #fh, "                   Refresh_Screen"
    Print #fh, "                     For nIndex = gcExit to gcLast"
    Print #fh, "                         cmdData(nIndex).Enabled = True"
    Print #fh, "                     Next"
    Print #fh, "                End If"
    Print #fh, "                bEditingRecord = False"
    Print #fh, "            Case 10"
    Print #fh, "                   on error resume next"
    Print #fh, "                  Screen.MousePointer = vbHourglass"
    Print #fh, "                   With dataRS"

    If chkComments.Value = vbChecked Then
        Print #fh, "                    ' If adding record"
    End If

    Print #fh, "                     If bAddingRecord = True Then"
    Print #fh, "                       .AddNew"
    If optDAO.Value = True Then Print #fh, "                     Else"

    If chkComments.Value = vbChecked Then
        Print #fh, "                    ' If editing record"
    End If

    If optDAO.Value = True Then Print #fh, "                       .Edit"
    Print #fh, "                      tmpBM = .Bookmark ' save current record"
    Print #fh, "                     End If"
    BuildTextToField fh

    If chkComments.Value = vbChecked Then
        Print #fh, "                     ' Update record"
    End If

    Print #fh, "                       .Update"
    Print #fh, "                       If Err.Number <> 0 Then errHandler"

    If chkComments.Value = vbChecked Then
        Print #fh, "                       'make sure we are at the same record"
    End If

    If optDAO.Value = True Then
        Print #fh, "                       .Move 0, .LastModified"
    Else
        Print #fh, "                       .Move 0"
    End If

    Print #fh, "                       If bAddingRecord = True Then .Bookmark = tmpBM"
    Print #fh, "                       Refresh_Screen"
    Print #fh, "                    End With"
    Print #fh, "                       For nIndex = gcExit to gcLast"
    Print #fh, "                       cmdData(nIndex).Enabled = True"
    Print #fh, "                       Next"
    Print #fh, "                       cmdData(gcCancel).Visible = False"
    Print #fh, "                       cmdData(gcUpdate).Visible = False"
    Print #fh, "                    bAddingRecord = False"
    Print #fh, "                    bEditingRecord = False"
    Print #fh, "                    Screen.MousePointer = vbDefault"
    Print #fh, "            Case Else   ' Do nothing"
    Print #fh, "        End Select"
    Print #fh, "   Exit Sub"
    Print #fh, ""
    Print #fh, "FindErr:"
    Print #fh, "      If not dataRS.EOF Then"
    Print #fh, "        If Err.Number <> 0 Then errHandler"
    Print #fh, "         Exit Sub"
    Print #fh, "      Else"
    Print #fh, "         gbNotFound = True"
    Print #fh, "         Resume Next"
    Print #fh, "   End If"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub Refresh_Screen()"

    Print #fh, ""
    Print #fh, "    Dim nIndex As Integer"
    Print #fh, "    on error resume next"

    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' Make sure we have a record to display."
        Print #fh, "    ' See if BOF or EOF are set"
        Print #fh, "    ' if not add new record "
    End If

    Print #fh, ""
    Print #fh, "    Clear_TextControls"
    Print #fh, "    With dataRS"
    Print #fh, "        If .EOF Or .BOF Then"
    Print #fh, "            Clear_TextControls"
    Print #fh, "            For nIndex = gcExit to gcLast"
    Print #fh, "                cmdData(nIndex).Enabled = False"
    Print #fh, "            Next"
    Print #fh, "            cmdData(gcAdd).Enabled = True"
    Print #fh, "        Else"
    Print #fh, "            bCurrentRecord = True"
    'here we need to see what mode were in and set defaults
    ' if field names have spaces or "-" then put brackets around name.
    BuildFieldToText fh
    Print #fh, "            bCurrentRecord = False"
    Print #fh, "            For nIndex = gcExit to gcLast"
    Print #fh, "                cmdData(nIndex).Enabled = True"
    Print #fh, "            Next"
    Print #fh, "        End If"
    Print #fh, "    End With"
    Print #fh, "    bEditingRecord = False"
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub Clear_TextControls()"

    Print #fh, "  on error resume next"
    Print #fh, ""
    Print #fh, "    bCurrentRecord = True"
    'write code for clearing fields

    For i = 0 To nNumFlds - 1
        sTmp = lstFields.List(i)
        Print #fh, "        txtField(" & i & ") = """
    Next i

    Print #fh, ""
    Print #fh, "        bCurrentRecord = False"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, ""
    Print #fh, "Private Sub SetUpdate_Mode()"

    Print #fh, "    Dim nIndex As Integer"
    Print #fh, "    on error resume next"

    Print #fh, "    If bEditingRecord Or bCurrentRecord Then Exit Sub"
    Print #fh, "    bEditingRecord = True"
    Print #fh, "    For nIndex = gcExit to gcLast"
    Print #fh, "        cmdData(nIndex).Enabled = False"
    Print #fh, "    Next"
    Print #fh, "    cmdData(gcCancel).Visible = True"
    Print #fh, "    cmdData(gcUpdate).Visible = True"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub Form_MouseMove(Button As Integer,Shift As Integer,x As Single,y As" & " Single)"

    Print #fh, "            on error resume next"
    Print #fh, "            dispstatus me," & Chr$(34) & "OK" & Chr$(34)
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub Form_Unload(Cancel As Integer)"

    Print #fh, " on error resume next"
    Print #fh, ""
    Print #fh, "    dataRS.Close"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub txtField_Change(Index as INTEGER)"

    Print #fh, "  on error resume next"
    Print #fh, "    SetUpdate_Mode"
    Print #fh, "End Sub"
    Print #fh, ""
    'write the code for the bound OLE client control(s)

    For i = 0 To lstOLECtls.ListCount - 1
        Print #fh, "Private Sub oleField" & lstOLECtls.List(i) & "_DblClick()"

        Print #fh, " on error resume next"

        If chkComments.Value = vbChecked Then
            Print #fh, "  'this is the way to get data into an empty ole control"
            Print #fh, "  'and have it saved back to the table"
        End If

        Print #fh, "  oleField" & lstOLECtls.List(i) & ".InsertObjDlg"
        Print #fh, "End Sub"
        Print #fh, ""
    Next i

    Print #fh, "Private Sub PrintAll" & RemoveSpace(gsRecordTable)

    Print #fh, "    Dim nIndex As Integer"
    Print #fh, "    Dim i As Integer"
    Print #fh, "    Dim dbSQL As String"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' Display the print dialog"
    End If

    Print #fh, "    frmMDI.comDlg.Flags = cdlPDPrintToFile + cdlPDNoSelection + cdlPDNoPageNums"
    Print #fh, "    frmMDI.comDlg.CancelError = True"
    Print #fh, "    On Error GoTo PrintCanceled"
    Print #fh, "    frmMDI.comDlg.ShowPrinter"
    Print #fh, "    On Error GoTo 0"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' start to printing"
        Print #fh, "    ' turn the cursor into an hourglass and set up the data"
    End If

    Print #fh, ""
    Print #fh, "    Screen.MousePointer = vbHourglass"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' If there are no records, disable buttons except addnew"
    End If

    Print #fh, "    If dataRS.BOF And dataRS.EOF Then"
    Print #fh, "        For nIndex = gcExit to gcLast"
    Print #fh, "            cmdData(nIndex).Enabled = False"
    Print #fh, "        Next"
    Print #fh, "        cmdData(gcAdd).Enabled = True"
    Print #fh, "    Else"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' loop through the records and print them out"
    End If

    Print #fh, "    On Error GoTo PrintCanceled"
    Print #fh, "    dataRS.MoveFirst"
    Print #fh, "    On Error GoTo 0"
    Print #fh, "    PrintHeading"
    Print #fh, ""
    Print #fh, "    Do While Not dataRS.EOF"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "        ' Print the record detail"
    End If

    Print #fh, "        With dataRS"
    ' check for memo field or multi line
    ' here we need to see what mode were in and set defaults
    ' if field names have spaces or "-" then put brackets around name.

    For i = 0 To nNumFlds - 1
        sTmp = lstFields.List(i)

        If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
            Print #fh, "       Printer.CurrentX = 5"
            Print #fh, "       Printer.Print " & Chr$(34) & sTmp & Chr$(34) & Chr$(34) & ": " & Chr$(34) & "!" & AddBrackets(sTmp)
        Else
            Print #fh, "       Printer.CurrentX = 5"
            Print #fh, "       Printer.Print " & Chr$(34) & sTmp & Chr$(34) & Chr$(34) & ": " & Chr$(34) & "!" & sTmp
        End If

    Next i
    Print #fh, "        End With"
    Print #fh, "       Printer.Print"
    Print #fh, "       Printer.Print"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "        ' near the bottom of the page so"
        Print #fh, "        ' start a new page and print the heading."
    End If

    Print #fh, ""
    Print #fh, "        If Printer.CurrentY > 45 Then"
    Print #fh, "       Printer.NewPage"
    Print #fh, "       PrintHeading"
    Print #fh, "       End If"
    Print #fh, ""
    Print #fh, "        dataRS.MoveNext"
    Print #fh, ""
    Print #fh, "    Loop"
    Print #fh, ""
    Print #fh, "       Printer.EndDoc"
    Print #fh, ""
    Print #fh, "       MsgBox " & Chr$(34) & "Windows  is printing all records." & Chr$(34) & "," & " vbInformation, " & Chr$(34) & "Print" & Chr$(34)
    Print #fh, ""
    Print #fh, "       Screen.MousePointer = vbDefault"
    Print #fh, ""
    Print #fh, "      End If"
    Print #fh, "Exit Sub"
    Print #fh, ""
    Print #fh, "PrintCanceled:"
    Print #fh, "    On Error GoTo 0"
    Print #fh, "    Exit Sub"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub PrintHeading()"

    Print #fh, ""
    Print #fh, "     Dim y as Integer"
    Print #fh, "     Printer.ScaleMode = 4 ' character mode"

    Print #fh, "     Printer.Font.Size = 10"
    Print #fh, "     Printer.Print " & Chr$(34) & "DATE :" & Chr$(34) & "Format(Now," & Chr$(34) & "Short Date" & Chr$(34) & ");"
    Print #fh, "     Printer.CurrentX = 52"
    Print #fh, "     Printer.Print ""REPORT: " & currentRS.Name & """; "
    Print #fh, "     Printer.CurrentX = 71"
    Print #fh, "     Printer.Print ""Page #"" & Printer.Page"
    Print #fh, ""
    Print #fh, "     y = Printer.CurrentY + 1  ' Set position for line."

    If chkComments.Value = vbChecked Then
        Print #fh, "     ' Draw a line across page."
    End If

    Print #fh, "     Printer.Line (0, y)-(Printer.ScaleWidth, y) ' Draw line."
    Print #fh, "     Printer.Print"
    Print #fh, "     Printer.Print"
    Print #fh, "     "
    Print #fh, "End Sub"
    Print #fh, ""
    '********** end *******************
Exit Sub

WCErr:
    MsgBox Error$

End Sub

Private Sub BuildPublicModNDC(fh As Integer)

Dim nNumFlds As Integer
Dim i        As Integer

    On Error GoTo WCErr
    nNumFlds = lstFields.ListCount
    Print #fh, "Attribute VB_Name = " & Chr$(34) & "PublicMod" & Chr$(34)
    Print #fh, " public gsFindField as String 'field to search"
    Print #fh, " public gsFindExpr as String  'search expression"
    Print #fh, " public gsFindOp as String    'search option"
    Print #fh, " public gbFindFailed as boolean 'failed flag "
    Print #fh, " public gbNotFound as boolean ' not found flag"
    Print #fh, " public gnFindType as Integer"
    Print #fh, " Public dbFile As String"
    Print #fh, "'works only with access  database"
    Print #fh, "'********************************************"

    If optDAO.Value = True Then

        Print #fh, " public gDB as DataBase"
    Else
        Print #fh, "'works only with access  database"
        Print #fh, "'********************************************"
        Print #fh, "public gDB   As New ADODB.Connection"
        Print #fh, "public dataRS  As New ADODB.Recordset"
        Print #fh, ""

    End If
    Print #fh, " Public Const vbCrLf = vbCr + vbLf"
    Print #fh,    ' define command button index number"

    Print #fh, "  ' make editing easier than remembering"
    Print #fh, "  ' button posistion "
    Print #fh, " Public Const gcExit = 0"
    Print #fh, " Public Const gcPrint = 1"
    Print #fh, " Public Const gcSearch = 2"
    Print #fh, " Public Const gcAdd = 3"
    Print #fh, " Public Const gcDelete = 4"
    Print #fh, " Public Const gcFirst = 5"
    Print #fh, " Public Const gcPrevious = 6"
    Print #fh, " Public Const gcNext = 7"
    Print #fh, " Public Const gcLast = 8"
    Print #fh, " Public Const gcCancel = 9"
    Print #fh, " Public Const gcUpdate = 10"
    Print #fh, ""

    Print #fh, "Public Sub dispstatus(f as form,s As String)"

    Print #fh, " ' displays message in label MsgBar an any form"
    Print #fh, " on error resume next"
    Print #fh, "           f.MsgBar.Caption = s"
    Print #fh, "End Sub"
    Print #fh, ""
    ' write error handler code using err object
    Print #fh, "  Public Sub errHandler()"

    Print #fh, "  dim msg as string"

    If chkComments.Value = vbChecked Then
        Print #fh, "' Check for error, then show message."
        Print #fh, "' uses error.object"
    End If

    Print #fh, "  msg = ""Error#"" & str(Err.Number) & "" was generated By "" & Err.Source & Chr(13)" & " & Err.Description "
    Print #fh, "  MsgBox msg" ', , Error, Err.HelpFile, Err.HelpContext"
    Print #fh, ""
    Print #fh, "  Err.Clear"
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "  Public Sub CenterMe (rfrm as Object)"

    Print #fh, " on error resume next"

    If chkComments.Value = vbChecked Then
        Print #fh, "     'center it on the MDI form"
    End If

    Print #fh, "      If rfrm.MDIChild = True Then"
    Print #fh, "        rfrm.Top = ((frmMDI.Height - rfrm.Height) \ 2) - 800"
    Print #fh, "        rfrm.Left = (frmMDI.Width - rfrm.Width) \ 2"
    Print #fh, "     Else"
    Print #fh, "        rfrm.Top = frmMDI.Top + (frmMDI.Height - rfrm.Height) \ 2"
    Print #fh, "       rfrm.Left = frmMDI.Left + (frmMDI.Width - rfrm.Width) \ 2"
    Print #fh, "     End If"
    Print #fh, "End Sub"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "'------------------------------------------------------------"
        Print #fh, "'this functions adds [] to object names that might need"
        Print #fh, "'them if they have spaces in them"
        Print #fh, "'------------------------------------------------------------"
    End If

    Print #fh, "Function AddBrackets(rObjName As String) As String"
    Print #fh, "  'add brackets to object names w/ spaces in them"
    Print #fh, "  If InStr(rObjName, "" "") > 0 Or InStr(rObjName, ""-"") > 0 And Mid(rObjName, 1," & " 1) <> ""["" Then"
    Print #fh, "    AddBrackets = ""["" & rObjName & ""]"""
    Print #fh, "  Else"
    Print #fh, "    AddBrackets = rObjName"
    Print #fh, "  End If"
    Print #fh, "End Function"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "'-----------------------------------------------------------"
        Print #fh, "' FUNCTION: FileExists"
        Print #fh, "' Determines whether the specified file exists"
        Print #fh, "' sPathName = file to check for"
        Print #fh, "'-----------------------------------------------------------"
        Print #fh, "'"
    End If

    Print #fh, "Function FileExists(ByVal sPathName As String) As Integer"
    Print #fh, "    Dim intFileNum As Integer"
    Print #fh, ""

    Print #fh, "    On Error Resume Next"

    If chkComments.Value = vbChecked Then
        Print #fh, "    '"
        Print #fh, "    'Remove  directory separator character"
        Print #fh, "    '"
    End If

    Print #fh, "    If Right$(sPathName, 1) = " & Chr$(34) & "\" & Chr$(34) & " then"
    Print #fh, "        sPathName = Left$(sPathName, Len(sPathName) - 1)"
    Print #fh, "    End If"
    Print #fh, "    '"

    If chkComments.Value = vbChecked Then
        Print #fh, "    'Attempt to open the file, return value of this function is False"
        Print #fh, "    'if an error occurs on open, True otherwise"
        Print #fh, "    '"
    End If

    Print #fh, "    intFileNum = FreeFile"
    Print #fh, "    Open sPathName For Input As intFileNum"
    Print #fh, "    FileExists = IIf(Err, False, True)"
    Print #fh, "    Close intFileNum"
    Print #fh, "    Err = 0"
    Print #fh, "End Function"
    Print #fh, ""
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "'----------------------------------"
        Print #fh, "' extract a filename from full path"
        Print #fh, "'----------------------------------"
    End If

    Print #fh, "Function ExtractFileName(ByVal s As String) As String"
    Print #fh, "    Dim i As Integer"
    Print #fh, ""

    Print #fh, "    On Error Resume Next"
    Print #fh, "    For i = Len(s) To 1 Step -1"
    Print #fh, "        If Mid$(s, i, 1) = " & Chr$(34) & "\" & Chr$(34) & " then"
    Print #fh, "            ExtractFileName = Right$(s, Len(s) - i)"
    Print #fh, "            Exit Function"
    Print #fh, "        End If"
    Print #fh, "    Next i"
    Print #fh, "    ExtractFileName = s"
    Print #fh, "End Function"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "'-----------------------------------"
        Print #fh, "' returns only numeric KeyPress"
        Print #fh, "'-----------------------------------"
    End If

    Print #fh, "Function ExtractNumber(KeyAscii As Integer) As Integer"
    Print #fh, "    On Error Resume Next"
    Print #fh, "    Select Case KeyAscii"
    Print #fh, "    Case Asc(" & Chr$(34) & "0" & Chr$(34) & ") To Asc(" & Chr$(34) & "9" & Chr$(34) & ")"
    Print #fh, "        ExtractNumber = KeyAscii"
    Print #fh, "    Case Asc(" & Chr$(34) & "." & Chr$(34) & ")"
    Print #fh, "        ExtractNumber = KeyAscii"
    Print #fh, "    Case Else"
    Print #fh, "        ExtractNumber = 0"
    Print #fh, "    End Select"
    Print #fh, "End Function"
    Print #fh, ""
    '********** end *******************

Exit Sub

WCErr:
    MsgBox Error$

End Sub

Private Sub BuildSubMainCreatedb(fh As Integer)

Dim nNumFlds As Integer


      On Error Resume Next
    nNumFlds = lstFields.ListCount
    'for windows before long filenames
    'If Len(gsRecordTable) > 7 Then
    'gsFormName = Left$(gsRecordTable, 7) & Trim(Str(giFormNumber))
    'Else
    gsFormName = gsRecordTable & Trim$(Str$(giFormNumber))
    'End If
    If optDAO.Value = True Then
        If chkCreateFilemod.Value = 1 Then
            Print #fh, "Sub CreateNewDatabase()"
            Print #fh, "    On Error Resume Next"
            Print #fh, ""
            Print #fh, "    Screen.MousePointer = vbDefault"
            Print #fh, ""

            If chkComments.Value = vbChecked Then
                Print #fh, "    ' Create the database"
            End If

            If chkComments.Value = vbChecked Then
                Print #fh, "    'uncomment next line for database in app dir"
                Print #fh, "    'Set gDB = CreateDatabase(dbFile, dbLangGeneral, dbVersio40" & ")"
                Print #fh, ""
                Print #fh, "    'comment next line for database in app dir"
            End If
           
           If chkCopyMDB.Value = vbChecked Then
            Print #fh, "    Set gDB = CreateDatabase(app.path &" & Chr$(34) & "\" & StripPath(gsMDBPath) & Chr$(34) & "," & " dbLangGeneral, dbVersio40" & ")"
              Else
            Print #fh, "    Set gDB = CreateDatabase(" & Chr$(34) & sDatabaseName & Chr$(34) & "," & " dbLangGeneral, dbVersio40" & ")"
           End If
            
            Print #fh, ""
            Print #fh, "    If gDB Is Nothing Then"
            Print #fh, "        Beep"
            Print #fh, "        MsgBox " & Chr$(34) & "Can't create database!" & Chr$(34) & "," & " vbExclamation"
            Print #fh, "        Exit Sub"
            Print #fh, "    End If"
            Print #fh, ""
            Print #fh, "    FrmWait.Show"
            Print #fh, "    "
            Print #fh, ""

            For giCurrentTableNumber = 0 To cboRecordSource.ListCount - 1
                giFormNumber = giCurrentTableNumber
                gsRecordTable = ""
                gsRecordTable = cboRecordSource.List(giCurrentTableNumber)

                If chkComments.Value = vbChecked Then
                    Print #fh, "    '*****************************************************"
                    Print #fh, "    '   create " & gsRecordTable & " database and index"
                    Print #fh, "    '*****************************************************"
                End If

                Print #fh, "    ProgressB FrmWait,( 1 / (" & cboRecordSource.ListCount & "-" & giCurrentTableNumber & ") * 100)"
                Print #fh, "    FrmWait.Status = " & Chr$(34) & "Createing Table: " & gsRecordTable & Chr$(34)
                Print #fh, "    Create" & RemoveSpace(gsRecordTable)
                Print #fh, "    If Err.Number <> 0 Then errHandler "
                Print #fh, ""
                Print #fh, "    "
                Print #fh, ""
            Next giCurrentTableNumber

            Print #fh, "    gDB.Close"
            Print #fh, "    If Err.Number <> 0 Then errHandler "
            Print #fh, ""
            Print #fh, "    FrmWait.Status = " & Chr$(34) & Chr$(34)
            Print #fh, "    unload FrmWait"
            Print #fh, "    Screen.MousePointer = vbDefault"
            Print #fh, ""
            Print #fh, "End Sub"
        End If
    End If 'end dao

    Print #fh, ""
    Print #fh, "Sub Main()"
    Print #fh, "    Dim nRV As Integer"
    Print #fh, "    Dim s As String"
    Print #fh, "    Dim n As Integer"
    Print #fh, "    Dim DatabaseFilePathe As String"
    Print #fh, "    On Error Resume Next"

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' Start"
    End If

    Print #fh, "    Screen.MousePointer = vbHourglass"
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    ' Load Mainform"
    End If

    Print #fh, "    Load frmMDI"
    Print #fh, "    frmMDI.Show"
    Print #fh, "    "
    Print #fh, ""
    Print #fh, "   On Error Resume Next"
    Print #fh, ""
    Print #fh, "    App.HelpFile = " & Chr$(34) & Chr$(34)
    Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "   'adds app path to file name"
        Print #fh, ""
        Print #fh, "    ' DatabaseFilePathe = DatabaseFilePathe & " & Chr$(34) & "\" & Chr$(34)
        Print #fh, "    'End If"
        Print #fh, "    'dbFile = DatabaseFilePathe & " & Chr$(34) & StripPath(gsAppPath) & ".mdb" & Chr$(34)
        Print #fh, ""
        Print #fh, "    ' if no database found then create."
        Print #fh, ""
        Print #fh, "     'search project and uncomment all dbFile code"
        Print #fh, "     'lines to use default app dir for database"
        Print #fh, "     'and comment the line with fixed database"
        Print #fh, ""
        Print #fh, "    'uncomment next line for database in app dir"
        Print #fh, "   'Do While Not FileExists(dbFile)"
        Print #fh, "    'comment line for database in app dir"
    End If

    If chkCopyMDB.Value = vbChecked Then
       Print #fh, "    Do While Not FileExists(app.path &" & Chr$(34) & "\" & StripPath(gsMDBPath) & Chr$(34) & ")"
       Print #fh, "    'Do While Not FileExists(" & Chr$(34) & sDatabaseName & Chr$(34) & ")"

        Else
       Print #fh, "    Do While Not FileExists(" & Chr$(34) & sDatabaseName & Chr$(34) & ")"
       Print #fh, "    'Do While Not FileExists(app.path &" & Chr$(34) & "\" & StripPath(gsMDBPath) & Chr$(34) & ")"

       End If
       
       Print #fh, ""

    If chkComments.Value = vbChecked Then
        Print #fh, "    'uncomment next line for database in app dir"
        Print #fh, "       'nRV = MsgBox(" & Chr$(34) & "Can't find Database Yes to create Database" & " in ""  & dbFile , vbCritical + vbYesNo)"
        Print #fh, "    'comment next line for database in app dir"
    End If

    Print #fh, "        nRV = MsgBox(" & Chr$(34) & "Can't find Database Yes to create Database in " & sDatabaseName & Chr$(34) & ", vbCritical + vbYesNo)"
    Print #fh, ""
    Print #fh, "        If (nRV = vbYes) Then"

    If chkComments.Value = vbChecked Then
        Print #fh, "          ' create New database"
    End If

    Print #fh, "            CreateNewDatabase"
    Print #fh, "        Else"
    Print #fh, "        End"
    Print #fh, "        End If"
    Print #fh, "    Loop"
    Print #fh, ""
    If optDAO.Value = True Then
        If chkComments.Value = vbChecked Then
            Print #fh, "    ' Open database."
            Print #fh, ""
            Print #fh, "    'uncomment next line for database in app dir"
            Print #fh, "   'Set gDB = Workspaces(0).OpenDatabase(dbFile)"
            Print #fh, ""
            Print #fh, "    'comment next line for database in app dir"
        End If
    If chkCopyMDB.Value = vbChecked Then
        Print #fh, "    Set gDB = Workspaces(0).OpenDatabase( app.path &" & Chr$(34) & "\" & StripPath(gsMDBPath) & Chr$(34) & ")"
    Else
            Print #fh, "    Set gDB = Workspaces(0).OpenDatabase(" & Chr$(34) & sDatabaseName & Chr$(34) & ")"

    End If
    Else
        Print #fh, "       With gDB"
        Print #fh, "        ' The following connection may or may not work on your computer."
        Print #fh, "          ' Alter it to find the your.mdb file"

        Print #fh, "         .Provider = ""Microsoft.Jet.OLEDB.4.0 """
        Print #fh, "    'uncomment next 2 lines for database in app dir"

        Print #fh, "    'dbFile = DatabaseFilePathe & " & Chr$(34) & StripPath(gsAppPath) & ".mdb" & Chr$(34)
        Print #fh, "    '      .Open dbFile"
        Print #fh, "    'comment next line for database in app dir"
        
        If chkCopyMDB.Value = vbChecked Then
        Print #fh, "          .Open app.path &" & Chr$(34) & "\" & StripPath(gsMDBPath) & Chr$(34)
        Print #fh, "          '.Open" & Chr$(34) & sDatabaseName & Chr$(34)
        Else
        Print #fh, "          .Open" & Chr$(34) & sDatabaseName & Chr$(34)
        Print #fh, "          '.Open app.path &" & Chr$(34) & "\" & StripPath(gsMDBPath) & Chr$(34)
        End If
        
        Print #fh, "     End With"
        Print #fh, "    'gDB.Open ""provider = microsoft.jet.oledb3.51DataSource = """ & Chr$(34) & sDatabaseName & Chr$(34)
    End If

    Print #fh, "    If (Err) Then"
    Print #fh, "        MsgBox " & Chr$(34) & "Can't open database file " & Chr$(34) & "," & " vbExclamation "
    Print #fh, "        End"
    Print #fh, "    End If"
    Print #fh, ""
    Print #fh, "    frmMDI.MsgBar = " & Chr$(34) & "Loaded" & Chr$(34)
    Print #fh, "    Screen.MousePointer = vbDefault"
    Print #fh, ""
    Print #fh, "End Sub"
    '
    Print #fh, ""
    Print #fh, "Private Sub ProgressB(f As Form, NV As Integer)"

    Print #fh, "'Drawmode = 6(this is very important.)"
    Print #fh, "'Name = ProgressBBar(change this if you like.)"
    Print #fh, "'If NV > 100 Or NV < 0 Then Exit Sub"
    Print #fh, " with f"
    Print #fh, "    .ProgressBBar.Cls ' Get rid of old value"
    Print #fh, "    .ProgressBBar.AutoRedraw = True"
    Print #fh, "    .ProgressBBar.DrawMode = 6"
    Print #fh, "    .ProgressBBar.FontSize = 12 ' Set this in code if needed"
    Print #fh, "    .ProgressBBar.ScaleMode = 0 ' Custom coordinates"
    Print #fh, "    .ProgressBBar.ScaleWidth = 100 ' 100 width because of percentages"
    Print #fh, "    .ProgressBBar.ScaleHeight = 10 ' Any value will do, but would have to chang next" & " line"
    Print #fh, "    .ProgressBBar.CurrentY = 2 ' A bit down"
    Print #fh, " ' Center it"
    Print #fh, "    .ProgressBBar.CurrentX = .ProgressBBar.ScaleWidth / 2 -" & " (.ProgressBBar.ScaleWidth / 15)"
    Print #fh, ""
    Print #fh, "    .ProgressBBar.Print Str(NV) & ""%"" ' Display value"
    Print #fh, " ' Draw the box. Any letters that were black and are covered now become white,"
    Print #fh, " ' while the others all stay the same."
    Print #fh, "    .ProgressBBar.Line (0, 0)-(NV, .ProgressBBar.ScaleHeight)," & " .ProgressBBar.FillColor, BF"
    Print #fh, " end with"
    Print #fh, "End Sub"

End Sub

Private Sub BuildTextToField(fh As Integer)


Dim sTmp     As String
Dim nNumFlds As Integer
Dim i        As Integer

    On Error Resume Next
    nNumFlds = lstFields.ListCount
    'here we need to see what mode were in and set defaults
    ' if field names have spaces or "-" then put brackets around name.
    'create the data  table

    For i = 0 To nNumFlds - 1
        sTmp = lstFields.List(i)

        Select Case currentRS.Fields(i).Type
        Case 1
            'fieldtype = "dbBoolean" 'Yes / No

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "                    !" & AddBrackets(sTmp) & " = cbool(txtField(" & i & "))"
                Print #fh, "                    If Err.Number <> 0 Then errHandler"
            Else
                Print #fh, "                    !" & sTmp & " = cbool(txtField(" & i & "))"
            End If

        Case 2
            'fieldtype = "dbByte" ' Byte

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "                    !" & AddBrackets(sTmp) & " = cbyte(txtField(" & i & "))"
            Else
                Print #fh, "                    !" & sTmp & " = cbyte(txtField(" & i & "))"
            End If

        Case 3
            '   fieldtype = "dbInteger" '   Integer

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "                    !" & AddBrackets(sTmp) & " = cInt(txtField(" & i & "))"
            Else
                Print #fh, "                    !" & sTmp & " = cInt(txtField(" & i & "))"
            End If

        Case 4
            '   fieldtype = "dbLong" ' Long

            If (currentRS.Fields(i).Attributes = dbAutoIncrField) Then
                If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                    Print #fh, "' auto inc           !" & AddBrackets(sTmp) & " = cLng(txtField(" & i & "))"
                Else
                    Print #fh, "' auto inc           !" & sTmp & " = cLng(txtField(" & i & "))"
                End If

            Else

                If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                    Print #fh, "                    !" & AddBrackets(sTmp) & " = cLng(txtField(" & i & "))"
                Else
                    Print #fh, "                    !" & sTmp & " = cLng(txtField(" & i & "))"
                End If

            End If

        Case 5
            ' fieldtype = "dbCurrency" '  Currency

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "                    !" & AddBrackets(sTmp) & " = ccur(txtField(" & i & "))"
            Else
                Print #fh, "                    !" & sTmp & " = ccur(txtField(" & i & "))"
            End If

        Case 6
            ' fieldtype = "dbSingle"  '  Single

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "                    !" & AddBrackets(sTmp) & " = cSng(txtField(" & i & "))"
            Else
                Print #fh, "                    !" & sTmp & " = cSng(txtField(" & i & "))"
            End If

        Case 7
            ' fieldtype = "dbDouble" '   Double

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "                    !" & AddBrackets(sTmp) & " = cdbl(txtField(" & i & "))"
            Else
                Print #fh, "                    !" & sTmp & " = cdbl(txtField(" & i & "))"
            End If

        Case 8
            ' fieldtype = "dbDate" 'Date / Time

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "                    !" & AddBrackets(sTmp) & " = cdate(txtField(" & i & "))"
            Else
                Print #fh, "                    !" & sTmp & " = cbool(txtField(" & i & "))"
            End If

        Case 10
            '  fieldtype = "dbText" ' Text

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "                    !" & AddBrackets(sTmp) & " = txtField(" & i & ")"
            Else
                Print #fh, "                    !" & sTmp & " = txtField(" & i & ")"
            End If

        Case 11 'dbLongBinary    Long Binary (OLE Object)
        Case 12
            ' fieldtype = "dbMemo" 'Memo

            If InStr(sTmp, " ") Or InStr(sTmp, "-") Then
                Print #fh, "                    !" & AddBrackets(sTmp) & " = txtField(" & i & ")"
            Else
                Print #fh, "                    !" & sTmp & " = txtField(" & i & ")"
            End If

        End Select
        Print #fh, "                    If Err.Number <> 0 Then errHandler"
    Next i

End Sub


