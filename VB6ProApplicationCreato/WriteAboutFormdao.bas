Attribute VB_Name = "WriteAboutFormdao"
Option Explicit

Public Sub WriteAboutForm(fh As Integer)

    Print #fh, "VERSION 5.00"
    Print #fh, "Begin VB.Form frmAbout"
    Print #fh, "   Appearance      =   0  'Flat"
    Print #fh, "   BackColor       =   &H00C0C0C0&"
    Print #fh, "   ClientHeight    =   2205"
    Print #fh, "   ClientLeft      =   1020"
    Print #fh, "   ClientTop       =   2640"
    Print #fh, "   ClientWidth     =   7455"
    Print #fh, "   BeginProperty Font"
    Print #fh, "      name            =   " & Chr$(34) & "MS Sans Serif" & Chr$(34)
    Print #fh, "      charset         =   0"
    Print #fh, "      weight          =   700"
    Print #fh, "      size            =   8.25"
    Print #fh, "      underline       =   0   'False"
    Print #fh, "      italic          =   0   'False"
    Print #fh, "      strikethrough   =   0   'False"
    Print #fh, "   EndProperty"
    Print #fh, "   ForeColor       =   &H80000008&"
    Print #fh, "   Height          =   2610"
    Print #fh, "   HelpContextID   =   1"
    Print #fh, "   Left            =   960"
    Print #fh, "   LinkTopic       =   " & Chr$(34) & "Form1" & Chr$(34)
    Print #fh, "   ScaleHeight     =   2205"
    Print #fh, "   ScaleWidth      =   7455"
    Print #fh, "   Top             =   2295"
    Print #fh, "   Width           =   7575"
    Print #fh, "   Begin VB.CommandButton Command1"
    Print #fh, "      Appearance      =   0  'Flat"
    Print #fh, "      BackColor       =   &H80000005&"
    Print #fh, "      Caption         =   " & Chr$(34) & "OK" & Chr$(34)
    Print #fh, "      Height          =   495"
    Print #fh, "      Left            =   6360"
    Print #fh, "      TabIndex        =   0"
    Print #fh, "      Top             =   1200"
    Print #fh, "      Width           =   885"
    Print #fh, "   End"
    Print #fh, "   Begin VB.Label Label3"
    Print #fh, "      Alignment       =   1  'Right Justify"
    Print #fh, "      Appearance      =   0  'Flat"
    Print #fh, "      BackColor       =   &H80000005&"
    Print #fh, "      BackStyle       =   0  'Transparent"
    Print #fh, "      Caption         =   " & Chr$(34) & "." & Chr$(34)
    Print #fh, "      ForeColor       =   &H80000008&"
    Print #fh, "      Height          =   255"
    Print #fh, "      Left            =   4020"
    Print #fh, "      TabIndex        =   6"
    Print #fh, "      Top             =   1200"
    Print #fh, "      Width           =   90"
    Print #fh, "   End"
    Print #fh, "   Begin VB.Label Label1"
    Print #fh, "      Alignment       =   1  'Right Justify"
    Print #fh, "      Appearance      =   0  'Flat"
    Print #fh, "      BackColor       =   &H80000005&"
    Print #fh, "      BackStyle       =   0  'Transparent"
    Print #fh, "      Caption         =   " & Chr$(34) & "Version " & Chr$(34)
    Print #fh, "      ForeColor       =   &H80000008&"
    Print #fh, "      Height          =   255"
    Print #fh, "      Left            =   2940"
    Print #fh, "      TabIndex        =   5"
    Print #fh, "      Top             =   1200"
    Print #fh, "      Width           =   750"
    Print #fh, "   End"
    Print #fh, "   Begin VB.Label Label4"
    Print #fh, "      Alignment       =   2  'Center"
    Print #fh, "      Appearance      =   0  'Flat"
    Print #fh, "      BackColor       =   &H80000005&"
    Print #fh, "      BackStyle       =   0  'Transparent"
    Print #fh, "      Caption         =   " & Chr$(34) & "Copyright  © 1996" & Chr$(34)
    Print #fh, "      ForeColor       =   &H80000008&"
    Print #fh, "      Height          =   285"
    Print #fh, "      Left            =   1440"
    Print #fh, "      TabIndex        =   2"
    Print #fh, "      Top             =   1560"
    Print #fh, "      Width           =   4695"
    Print #fh, "   End"
    Print #fh, "   Begin VB.Label pVersion"
    Print #fh, "      Alignment       =   1  'Right Justify"
    Print #fh, "      Appearance      =   0  'Flat"
    Print #fh, "      BackColor       =   &H80000005&"
    Print #fh, "      BackStyle       =   0  'Transparent"
    Print #fh, "      Caption         =   " & Chr$(34) & "maj" & Chr$(34)
    Print #fh, "      ForeColor       =   &H80000008&"
    Print #fh, "      Height          =   255"
    Print #fh, "      Left            =   3240"
    Print #fh, "      TabIndex        =   3"
    Print #fh, "      Top             =   1200"
    Print #fh, "      Width           =   750"
    Print #fh, "   End"
    Print #fh, "   Begin VB.Label Label2"
    Print #fh, "      Alignment       =   2  'Center"
    Print #fh, "      Appearance      =   0  'Flat"
    Print #fh, "      BackColor       =   &H00C0C0C0&"
    Print #fh, "      Caption         =   " & Chr$(34) & gsAppName & Chr$(34)
    Print #fh, "      BeginProperty Font"
    Print #fh, "         name            =   " & Chr$(34) & "MS Sans Serif" & Chr$(34)
    Print #fh, "         charset         =   0"
    Print #fh, "         weight          =   700"
    Print #fh, "         size            =   24"
    Print #fh, "         underline       =   0   'False"
    Print #fh, "         italic          =   0   'False"
    Print #fh, "         strikethrough   =   0   'False"
    Print #fh, "      EndProperty"
    Print #fh, "      ForeColor       =   &H00FF0000&"
    Print #fh, "      Height          =   615"
    Print #fh, "      Left            =   1200"
    Print #fh, "      TabIndex        =   4"
    Print #fh, "      Top             =   240"
    Print #fh, "      Width           =   5055"
    Print #fh, "   End"
    Print #fh, "   Begin VB.Label pRevision"
    Print #fh, "      Appearance      =   0  'Flat"
    Print #fh, "      BackColor       =   &H80000005&"
    Print #fh, "      BackStyle       =   0  'Transparent"
    Print #fh, "      Caption         =   " & Chr$(34) & "min" & Chr$(34)
    Print #fh, "      ForeColor       =   &H80000008&"
    Print #fh, "      Height          =   255"
    Print #fh, "      Left            =   4140"
    Print #fh, "      TabIndex        =   1"
    Print #fh, "      Top             =   1200"
    Print #fh, "      Width           =   870"
    Print #fh, "   End"
    Print #fh, "   Begin VB.Image Image1"
    Print #fh, "      Appearance      =   0  'Flat"
    Print #fh, "      Height          =   15"
    Print #fh, "      Left            =   0"
    Print #fh, "      Top             =   0"
    Print #fh, "      Width           =   15"
    Print #fh, "   End"
    Print #fh, "End"
    Print #fh, "Attribute VB_Name = " & Chr$(34) & "frmAbout" & Chr$(34)
    Print #fh, "Attribute VB_Creatable = False"
    Print #fh, "Attribute VB_Exposed = False"
    Print #fh, ""
    Print #fh, "Option Explicit"
    Print #fh, "Private Sub Command1_Click()"
    Print #fh, "    Unload Me"
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub Form_Load()"
    Print #fh, "    CenterMe Me"
    Print #fh, ""
    Print #fh, "    pVersion.Caption = App.Major"
    Print #fh, "    pRevision.Caption = App.Minor"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, "Private Sub reversion_Click()"
    Print #fh, ""
    Print #fh, "End Sub"
    Print #fh, ""
    Print #fh, ""

End Sub

