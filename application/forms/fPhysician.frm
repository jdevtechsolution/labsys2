VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fPhysician 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fPhysician.frx":0000
   ScaleHeight     =   10050
   ScaleWidth      =   15465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnNew 
      Caption         =   "&N"
      Height          =   495
      Left            =   15960
      TabIndex        =   22
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9855
      Left            =   120
      ScaleHeight     =   9825
      ScaleWidth      =   15225
      TabIndex        =   0
      Top             =   120
      Width           =   15255
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6735
         Left            =   120
         ScaleHeight     =   6705
         ScaleWidth      =   2745
         TabIndex        =   9
         Top             =   3000
         Width           =   2775
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   2175
            Left            =   120
            Picture         =   "fPhysician.frx":1C8B9
            Stretch         =   -1  'True
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   3840
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
            TabIndex        =   15
            Top             =   4080
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
            TabIndex        =   14
            Top             =   4320
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
            TabIndex        =   13
            Top             =   3360
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
            TabIndex        =   12
            Top             =   3600
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
            TabIndex        =   11
            Top             =   4800
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
            TabIndex        =   10
            Top             =   4560
            Width           =   2325
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   8775
         Left            =   3000
         ScaleHeight     =   8745
         ScaleWidth      =   12105
         TabIndex        =   1
         Top             =   960
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
            TabIndex        =   2
            Top             =   120
            Width           =   3975
         End
         Begin prjLabSys.isButton cmdNew 
            Height          =   330
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            Icon            =   "fPhysician.frx":1D4C7
            Style           =   6
            Caption         =   "Create &New Record"
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
            TabIndex        =   4
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
            TabIndex        =   5
            Top             =   120
            Width           =   1005
         End
      End
      Begin MSComctlLib.TreeView tvOptions 
         Height          =   1740
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   3069
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
         Caption         =   "Menus"
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
         TabIndex        =   21
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Physician Information"
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
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   2100
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   120
         Picture         =   "fPhysician.frx":1D4E3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   60
         Left            =   960
         Top             =   720
         Width           =   14175
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Create new, edit, remove and view Physician Information."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   960
         TabIndex        =   7
         Top             =   480
         Width           =   4200
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   60
         Left            =   5280
         Top             =   600
         Width           =   9855
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
            Picture         =   "fPhysician.frx":1DFA9
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1E2FB
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1E64D
            Key             =   "Save As"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1E99F
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1ECF1
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1F043
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1F395
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1F6E7
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1FA39
            Key             =   "Replace"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":1FD8B
            Key             =   "Go To"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":200DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":20677
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":20C11
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":211AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":21745
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":21CDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":22279
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":22813
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":22DAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":23347
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":238E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":23E7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":24415
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":249AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":24F49
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":256C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":260D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":2684F
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":27261
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":279DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysician.frx":28155
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit selected"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove selected"
      End
   End
End
Attribute VB_Name = "fPhysician"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnNew_Click()
    cmdNew_Click
End Sub

Private Sub cmdNew_Click()
    With fPhysicianEntry
        .mode = eAddNew
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
    Dim rsPhysician As New ADODB.Recordset, vSQL As String, vRowIndex As Integer
    With rsPhysician
        vSQL = "SELECT a.physician_id,CONCAT_WS(' ',a.title,a.surname,a.firstname,a.middlename) as Physician," & _
        "a.mobile_no,a.email,a.hospital " & _
        " FROM physicians as a WHERE a.is_deleted=FALSE " & _
        IIf(IsMissing(vCri), "", " AND (a.surname LIKE '" & CStr(vCri) & "%' OR a.firstname LIKE '" & CStr(vCri) & _
        "%' OR a.middlename LIKE '" & CStr(vCri) & "%')")
        
        OpenTable vSQL, rsPhysician
        lgPhysician.Redraw = False
        lgPhysician.Clear
        If .RecordCount > 0 Then
            Do Until .EOF
                vRowIndex = lgPhysician.AddItem(!Physician & vbTab & !hospital & vbTab & !mobile_no & vbTab & !email)
                lgPhysician.RowTag(vRowIndex) = !physician_id
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





















Private Sub lgPhysician_DblClick()
    mnuEdit_Click
End Sub

Private Sub lgPhysician_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then mnuEdit_Click
End Sub

Private Sub lgPhysician_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuOptions
    End If
End Sub

Private Sub mnuEdit_Click()
    If lgPhysician.ItemCount > 0 Then
        Dim vID As String
        vID = lgPhysician.RowTag(lgPhysician.row)
        fPhysicianEntry.vSelectedID = vID
        fPhysicianEntry.mode = eEdit
        fPhysicianEntry.CallingForm = ePhysician
        SetPhysicianDetails vID
        fPhysicianEntry.Show 1
    Else
        MsgBox "No record selected.          ", vbExclamation, system_title
    End If
End Sub



Private Sub SetPhysicianDetails(vID As String)
    Dim rs As New ADODB.Recordset, vSQL As String
    With rs
        vSQL = "SELECT a.* FROM physicians as a WHERE a.physician_id=" & val(vID)
        OpenTable vSQL, rs
        If .RecordCount > 0 Then
            With fPhysicianEntry
                .txtSurname.Text = "" & rs!surname
                .txtFirstname.Text = "" & rs!firstname
                .txtMiddlename.Text = "" & rs!middlename
                .txtTitle.Text = "" & rs!Title
                .txtMobile.Text = "" & rs!mobile_no
                .txtLandline.Text = "" & rs!landline
                .txtEmail.Text = "" & rs!email
                .txtAddress.Text = "" & rs!address
                .txtHospital.Text = "" & rs!hospital
                .txtSchool.Text = "" & rs!school
                .txtDepartment.Text = "" & rs!department
                .txtExpertise.Text = "" & rs!expertise
                
            End With
        End If
        .Close
        Set rs = Nothing
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
            If lgPhysician.ItemCount > 0 Then lgPhysician.row = 0
        End If
    End If
End Sub





Private Sub txtSearch_LostFocus()
    BackColorOnFocus txtSearch, vbWhite
End Sub
