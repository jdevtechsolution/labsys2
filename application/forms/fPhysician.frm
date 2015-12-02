VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fPhysician 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fPhysician.frx":0000
   ScaleHeight     =   9690
   ScaleWidth      =   15465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9495
      Left            =   120
      ScaleHeight     =   9465
      ScaleWidth      =   15225
      TabIndex        =   0
      Top             =   120
      Width           =   15255
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   8775
         Left            =   3000
         ScaleHeight     =   8745
         ScaleWidth      =   12105
         TabIndex        =   14
         Top             =   480
         Width           =   12135
         Begin VB.TextBox txtSearch 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8040
            TabIndex        =   15
            Top             =   120
            Width           =   3975
         End
         Begin prjLabSys.isButton cmdNew 
            Height          =   330
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            Icon            =   "fPhysician.frx":1C8B9
            Style           =   6
            Caption         =   "&Create New Record"
            iNonThemeStyle  =   0
            Object.ToolTipText     =   ""
            ToolTipTitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            RoundedBordersByTheme=   0   'False
         End
         Begin prjLabSys.LynxGrid lgPhysician 
            Height          =   8175
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   14420
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorSel    =   12937777
            ForeColorSel    =   16777215
            CustomColorFrom =   16572875
            CustomColorTo   =   14722429
            GridColor       =   16367254
            FocusRectColor  =   9895934
            Appearance      =   0
            ColumnHeaderSmall=   -1  'True
            TotalsLineShow  =   0   'False
            FocusRowHighlightKeepTextForecolor=   0   'False
            ShowRowNumbers  =   0   'False
            ShowRowNumbersVary=   0   'False
            AllowColumnResizing=   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Search here :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6960
            TabIndex        =   18
            Top             =   120
            Width           =   1005
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   120
         ScaleHeight     =   6225
         ScaleWidth      =   2745
         TabIndex        =   1
         Top             =   3000
         Width           =   2775
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2175
            Left            =   120
            ScaleHeight     =   2145
            ScaleWidth      =   2505
            TabIndex        =   2
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Details"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   30
            Width           =   1245
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   2640
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label lblFullname 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   2880
            Width           =   2325
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Full Name :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   2640
            Width           =   2325
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Address :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   3120
            Width           =   2325
         End
         Begin VB.Label lblAddress 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   9
            Top             =   3360
            Width           =   2325
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   4320
            Width           =   2325
         End
         Begin VB.Label lblMobile 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   4560
            Width           =   2325
         End
         Begin VB.Label lblBirthdate 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   4080
            Width           =   2325
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Birthdate :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   3840
            Width           =   2325
         End
         Begin VB.Label lblEmail 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   5040
            Width           =   2325
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Email Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   4800
            Width           =   2325
         End
      End
      Begin MSComctlLib.TreeView tvOptions 
         Height          =   2460
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4339
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   617
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "SmallIcons"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Physician Management"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   2220
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   15120
         Y1              =   360
         Y2              =   360
      End
   End
   Begin MSComctlLib.ImageList SmallIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1C8D5
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1CC27
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1CF79
            Key             =   "Save As"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1D2CB
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1D61D
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1D96F
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1DCC1
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1E013
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1E365
            Key             =   "Replace"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1E6B7
            Key             =   "Go To"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1EA09
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1EFA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1F53D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1FAD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":20071
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":2060B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":20BA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":2113F
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":216D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":21C73
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":2220D
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":227A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":22D41
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":232DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":23875
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":23FEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":24A01
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":2517B
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":25B8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":26307
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":26A81
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fPhysician"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNew_Click()
    With fPhysicianEntry
        .Mode = eAddNew
        .CallingForm = ePhysician
        .Show 1
    End With
End Sub

Private Sub Form_Load()
    InitializeConnectionString
    Call OpenConnection
    Center Me
    CreateNodeOptions
    InitializeGridProperty
    CreateHeaders
    ShowPhysicianList
End Sub


Private Sub CreateHeaders()
    Dim nWidth As Integer
    With lgPhysician
        nWidth = .Width
        .AddColumn "Physician", nWidth * 0.34
        .AddColumn "Hospital / Clinic", nWidth * 0.26
        .AddColumn "Mobile", nWidth * 0.15
        .AddColumn "Email", nWidth * 0.2
    End With
End Sub


Public Sub ShowPhysicianList(Optional vCri As Variant)
    Dim rsPhysician As New ADODB.Recordset, vSQL As String
    With rsPhysician
        vSQL = "SELECT CONCAT_WS(' ',a.title,a.surname,a.firstname,a.middlename) as Physician," & _
        "a.mobile_no,a.email,a.hospital " & _
        " FROM physicians as a WHERE a.is_deleted=FALSE " & _
        IIf(IsMissing(vCri), "", " AND (a.surname LIKE '" & CStr(vCri) & "%' OR a.firstname LIKE '" & CStr(vCri) & _
        "%' OR a.middlename LIKE '" & CStr(vCri) & "%')")
        
        OpenTable vSQL, rsPhysician
        lgPhysician.Redraw = False
        lgPhysician.Clear
        If .RecordCount > 0 Then
            Do Until .EOF
                lgPhysician.AddItem !Physician & vbTab & !hospital & vbTab & !mobile_no & vbTab & !email
                .MoveNext
            Loop
        End If
        lgPhysician.Redraw = True
        .Close
        Set rsPhysician = Nothing
    End With
End Sub



Private Sub CreateNodeOptions()
    With tvOptions
        .Nodes.Add , , "1A", "Physician Information", 1
        .Nodes.Add , , "2A", "Recent Activity", 1
    End With
End Sub


Private Sub InitializeGridProperty()
    With lgPhysician
        '.ImageList = ImageList1
        .Redraw = False                                 'do not draw
        .AllowEdit = False
        .AllowDelete = True
        .AllowColumnResizing = False
        .ScrollBarStyle = Style_Regular
        .FocusRectStyle = lgFRHeavy
        .FocusRectMode = lgCol
        .FocusRectColor = vbYellow
        .FocusRowHighlight = True                       'this will highlight whole row
        .AllowColumnSort = False                        'header will not be clickable to sort
        .BackColorEvenRowsEnabled = True
        .BackColorBkg = &HFFFFFF
        .BackColorEdit = &HF2FEFF
        .Redraw = True 'grid is ready so draw
    End With
End Sub





















Private Sub txtSearch_Change()
    If txtSearch.Text = "" Then ShowPhysicianList
End Sub

Private Sub txtSearch_GotFocus()
    BackColorOnFocus txtSearch, vbYellow
End Sub



Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not txtSearch.Text = "" Then
            ShowPhysicianList txtSearch
        End If
    End If
End Sub





Private Sub txtSearch_LostFocus()
    BackColorOnFocus txtSearch, vbWhite
End Sub
