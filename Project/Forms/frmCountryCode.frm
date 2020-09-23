VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCountryCode 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Country Code"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin prjChameleon.chameleonButton cmdAbout 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "About"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmCountryCode.frx":0000
   End
   Begin prjChameleon.chameleonButton cmdExit 
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmCountryCode.frx":001C
   End
   Begin MSFlexGridLib.MSFlexGrid mfgList 
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6376
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "Code                   |  Description                                                                               "
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.ComboBox cboDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   4095
      End
      Begin VB.ComboBox cboCountryCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Country Code:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.Frame fraAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   2040
      TabIndex        =   9
      Top             =   1920
      Width           =   2895
      Begin prjChameleon.chameleonButton cmdOK 
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   1560
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "&OK"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         MPTR            =   0
         MICON           =   "frmCountryCode.frx":0038
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Email:Arkanghel@hotmail.com"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   2145
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Author: Michael Dungog"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1725
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   " About..."
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.Frame fraMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   -120
      TabIndex        =   8
      Top             =   -240
      Width           =   6495
   End
End
Attribute VB_Name = "frmCountryCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cnxn As New ADODB.Connection, strCnxn As String
Dim rsCountryCode As New ADODB.Recordset

Private Sub cboCountryCode_Click()
    rsCountryCode.Open "SELECT * FROM countrycode WHERE code='" & Trim(cboCountryCode) & "'", Cnxn, adOpenKeyset, adLockOptimistic, adCmdText
        cboDescription = rsCountryCode!Description
    rsCountryCode.Close
End Sub

Private Sub cboDescription_Click()
    rsCountryCode.Open "SELECT * FROM countrycode WHERE description='" & Trim(cboDescription) & "'", Cnxn, adOpenKeyset, adLockOptimistic, adCmdText
        cboCountryCode = rsCountryCode!code
    rsCountryCode.Close
End Sub

Private Sub Connection()
    strCnxn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Project\Database\CountryCode.mdb;Persist Security Info=False"
    Cnxn.Open strCnxn
End Sub

Private Sub chameleonButton1_Click()
    
End Sub

Private Sub cmdAbout_Click()
    fraMain.Enabled = False
    fraAbout.Enabled = True
    fraAbout.Visible = True
    fraAbout.ZOrder
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    fraAbout.Visible = False
    fraMain.Enabled = True
End Sub

Private Sub Form_Load()
    Call Connection
    rsCountryCode.Open "SELECT * FROM countrycode", Cnxn, adOpenKeyset, adLockOptimistic, adCmdText
        Do While rsCountryCode.EOF = False
            cboCountryCode.AddItem rsCountryCode!code
            cboDescription.AddItem rsCountryCode!Description
            rsCountryCode.MoveNext
        Loop
    rsCountryCode.Close
    
    rsCountryCode.Open "SELECT * FROM countrycode", Cnxn, adOpenKeyset, adLockOptimistic, adCmdText
    mfgList.Rows = 1
        If rsCountryCode.RecordCount <> 0 Then
            Do While rsCountryCode.EOF = False
                mfgList.AddItem rsCountryCode!code & vbTab & rsCountryCode!Description
                rsCountryCode.MoveNext
            Loop
        End If

    rsCountryCode.Close

End Sub

Private Sub mfgList_DblClick()
Dim strCode As String
    strCode = mfgList.TextMatrix(mfgList.Row, 0)
    
    rsCountryCode.Open "SELECT * FROM countrycode WHERE code='" & strCode & "'", Cnxn, adOpenKeyset, adLockOptimistic, adCmdText
        cboCountryCode.Text = rsCountryCode!code
        cboDescription.Text = rsCountryCode!Description
    rsCountryCode.Close
End Sub
