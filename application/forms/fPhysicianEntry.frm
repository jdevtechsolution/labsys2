VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fPhysicianEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fPhysicianEntry.frx":0000
   ScaleHeight     =   7860
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   120
      ScaleHeight     =   6945
      ScaleWidth      =   6945
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   360
      Width           =   6975
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   4920
         ScaleHeight     =   1545
         ScaleWidth      =   1785
         TabIndex        =   32
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   1200
         TabIndex        =   6
         Top             =   3240
         Width           =   5535
      End
      Begin VB.TextBox txtExpertise 
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
         Height          =   645
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   5880
         Width           =   5535
      End
      Begin VB.TextBox txtDepartment 
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
         Left            =   1200
         TabIndex        =   10
         Top             =   5520
         Width           =   5535
      End
      Begin VB.TextBox txtSchool 
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
         Left            =   1200
         TabIndex        =   9
         Top             =   5160
         Width           =   5535
      End
      Begin VB.TextBox txtLandline 
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
         Left            =   4680
         TabIndex        =   5
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox txtMobile 
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
         Left            =   1200
         TabIndex        =   4
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox txtHospital 
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
         Left            =   1200
         TabIndex        =   8
         Top             =   4200
         Width           =   5535
      End
      Begin VB.TextBox txtAddress 
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
         Height          =   525
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3600
         Width           =   5535
      End
      Begin VB.TextBox txtFirstname 
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
         Left            =   1200
         TabIndex        =   1
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtMiddlename 
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
         Left            =   1200
         TabIndex        =   2
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtSurname 
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
         Left            =   1200
         TabIndex        =   0
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtTitle 
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
         Left            =   1200
         TabIndex        =   3
         Top             =   1800
         Width           =   3615
      End
      Begin prjLabSys.isButton cmdBrowse 
         Height          =   330
         Left            =   4920
         TabIndex        =   33
         Top             =   2220
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         Icon            =   "fPhysicianEntry.frx":1C8B9
         Style           =   6
         Caption         =   "&Browse"
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
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Email :"
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
         Left            =   240
         TabIndex        =   29
         Top             =   3240
         Width           =   885
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Expertise :"
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
         Left            =   0
         TabIndex        =   28
         Top             =   5880
         Width           =   1125
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Department :"
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
         Left            =   0
         TabIndex        =   27
         Top             =   5520
         Width           =   1125
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "School :"
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
         Left            =   240
         TabIndex        =   26
         Top             =   5160
         Width           =   885
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Educational Information"
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
         TabIndex        =   25
         Top             =   4800
         Width           =   2040
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   6720
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Landline :"
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
         Left            =   3720
         TabIndex        =   24
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mobile :"
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
         Left            =   240
         TabIndex        =   23
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hospital :"
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
         Left            =   240
         TabIndex        =   22
         Top             =   4200
         Width           =   885
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Address :"
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
         Left            =   240
         TabIndex        =   21
         Top             =   3600
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Personal Information"
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
         TabIndex        =   19
         Top             =   240
         Width           =   1800
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   6840
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Firstname :"
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
         Left            =   0
         TabIndex        =   18
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Middlename :"
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
         Left            =   0
         TabIndex        =   17
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "* Surname :"
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
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Title :"
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
         Left            =   0
         TabIndex        =   15
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   6720
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contact Information"
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
         TabIndex        =   14
         Top             =   2400
         Width           =   1725
      End
   End
   Begin prjLabSys.isButton cmdClose 
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "fPhysicianEntry.frx":1C8D5
      Style           =   6
      Caption         =   "&Close"
      iNonThemeStyle  =   0
      Object.ToolTipText     =   ""
      ToolTipTitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
   Begin MSComctlLib.ImageList SmallIcons 
      Left            =   720
      Top             =   2160
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
            Picture         =   "fPhysicianEntry.frx":1C8F1
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1CC43
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1CF95
            Key             =   "Save As"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1D2E7
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1D639
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1D98B
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1DCDD
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1E02F
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1E381
            Key             =   "Replace"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1E6D3
            Key             =   "Go To"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1EA25
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1EFBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1F559
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":1FAF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":2008D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":20627
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":20BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":2115B
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":216F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":21C8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":22229
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":227C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":22D5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":232F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":23891
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":2400B
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":24A1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":25197
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":25BA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":26323
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPhysicianEntry.frx":26A9D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin prjLabSys.isButton cmdSave 
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Top             =   7440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Icon            =   "fPhysicianEntry.frx":27217
      Style           =   6
      Caption         =   "&Save Changes"
      iNonThemeStyle  =   0
      Object.ToolTipText     =   ""
      ToolTipTitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Physician Adding / Editing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   2520
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ">   Please provide all required information."
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
      Left            =   2760
      TabIndex        =   30
      Top             =   150
      Width           =   3060
   End
End
Attribute VB_Name = "fPhysicianEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private eMode As eTxnMode
Private eActiveCallingForm As eCallingForm




Private Type PhysicianInfo
    vPhysicianID As String
    vPhysicianName As String
    vMobileNo As String
    vLandLine As String
    vEmail As String
    vHospital As String
End Type




Public Property Let Mode(eStat As eTxnMode)
    eMode = eStat
End Property

Public Property Get Mode() As eTxnMode
    Mode = eMode
End Property

Public Property Let CallingForm(eCallForm As eCallingForm)
    eActiveCallingForm = eCallForm
End Property

Public Property Get CallingForm() As eCallingForm
    CallingForm = eActiveCallingForm
End Property


Public Function StrEscape(vString As String) As String
    StrEscape = Replace(Trim(vString), "'", "\'")
    StrEscape = Replace(Trim(StrEscape), ",", "\,")
    StrEscape = Replace(Trim(StrEscape), ";", "\;")
End Function






Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If validateRequiredFields Then
        If Mode = eAddNew Then
            If CreatePhysicianInfo Then
                MsgBox "Physician record successfully created.         ", vbInformation, system_title
                ClearControls
                txtSurname.SetFocus
                
                If CallingForm = ePhysician Then 'parent form is Physician List
                    
                Else
                
                End If
                
                
            End If
        Else
        
        End If
    End If
End Sub


Private Function CreatePhysicianInfo() As Boolean
    On Error GoTo err
    Dim vSQL As String
    
    vSQL = "INSERT INTO physicians SET surname='" & StrEscape(txtSurname) & "',firstname='" & StrEscape(txtFirstname) & "'," & _
    "middlename='" & StrEscape(txtMiddlename) & "',title='" & StrEscape(txtTitle) & "',mobile_no='" & StrEscape(txtMobile) & "'," & _
    "landline='" & StrEscape(txtLandline) & "',email='" & StrEscape(txtEmail) & "',address='" & StrEscape(txtAddress) & _
    "',hospital='" & StrEscape(txtHospital) & "',school='" & StrEscape(txtSchool) & "',department='" & StrEscape(txtDepartment) & "'," & _
    "expertise='" & StrEscape(txtExpertise) & "'"
    dbCon.Execute vSQL
    CreatePhysicianInfo = True
    
    Exit Function
err:
    CreatePhysicianInfo = False
    MsgBox err.Description, vbExclamation, ""
End Function





Private Function validateRequiredFields() As Boolean
    If txtSurname.Text = "" Then
        MsgBox "Surname is required.     ", vbExclamation, system_title
        txtSurname.SetFocus
        validateRequiredFields = False
        Exit Function
    End If
    
    validateRequiredFields = True
End Function





Private Sub Form_Load()
    Call OpenConnection
    Center Me
    
End Sub


Private Sub txtAddress_GotFocus()
    BackColorOnFocus txtAddress, vbYellow
End Sub

Private Sub txtAddress_LostFocus()
    BackColorOnFocus txtAddress, vbWhite
End Sub

Private Sub txtDepartment_GotFocus()
        BackColorOnFocus txtDepartment, vbYellow
End Sub

Private Sub txtDepartment_LostFocus()
        BackColorOnFocus txtDepartment, vbWhite
End Sub

Private Sub txtEmail_GotFocus()
    BackColorOnFocus txtEmail, vbYellow
End Sub

Private Sub txtEmail_LostFocus()
    BackColorOnFocus txtEmail, vbWhite
End Sub

Private Sub txtExpertise_GotFocus()
        BackColorOnFocus txtExpertise, vbYellow
End Sub

Private Sub txtExpertise_LostFocus()
        BackColorOnFocus txtExpertise, vbWhite
End Sub

Private Sub txtFirstname_GotFocus()
    BackColorOnFocus txtFirstname, vbYellow
End Sub

Private Sub txtFirstname_LostFocus()
    BackColorOnFocus txtFirstname, vbWhite
End Sub

Private Sub txtHospital_GotFocus()
    BackColorOnFocus txtHospital, vbYellow
End Sub

Private Sub txtHospital_LostFocus()
    BackColorOnFocus txtHospital, vbWhite
End Sub

Private Sub txtLandline_GotFocus()
    BackColorOnFocus txtLandline, vbYellow
End Sub

Private Sub txtLandline_LostFocus()
    BackColorOnFocus txtLandline, vbWhite
End Sub

Private Sub txtMiddlename_GotFocus()
    BackColorOnFocus txtMiddlename, vbYellow
End Sub

Private Sub txtMiddlename_LostFocus()
    BackColorOnFocus txtMiddlename, vbWhite
End Sub

Private Sub txtMobile_GotFocus()
    BackColorOnFocus txtMobile, vbYellow
End Sub

Private Sub txtMobile_LostFocus()
    BackColorOnFocus txtMobile, vbWhite
End Sub

Private Sub txtSchool_GotFocus()
    BackColorOnFocus txtSchool, vbYellow
End Sub

Private Sub txtSchool_LostFocus()
    BackColorOnFocus txtSchool, vbWhite
End Sub

Private Sub txtSurname_GotFocus()
    BackColorOnFocus txtSurname, vbYellow
End Sub

Private Sub txtSurname_LostFocus()
    BackColorOnFocus txtSurname, vbWhite
End Sub

Private Sub txtTitle_GotFocus()
        BackColorOnFocus txtTitle, vbYellow
End Sub

Private Sub txtTitle_LostFocus()
        BackColorOnFocus txtTitle, vbWhite
End Sub


Private Sub ClearControls()
    Dim c As Control
    For Each c In Me.Controls
        If TypeOf c Is TextBox Then
            c.Text = ""
        End If
    Next
End Sub





Private Function GetLastAffectedRowDetails(vID As String) As PhysicianInfo
    Dim rsLastRow As New ADODB.Recordset, vSQL As String
    With rsLastRow
        vSQL = "SELECT a.physician_id,CONCAT_WS(' ',a.title,a.surname,a.firstname,a.middlename) as Physician," & _
        "a.mobile_no,a.email,a.hospital FROM physicians as a WHERE a.physician_id=" & val(vID)
        OpenTable vSQL, rsLastRow
        If .RecordCount > 0 Then
            GetLastAffectedRowDetails.vPhysicianID = "" & !physician_id
            GetLastAffectedRowDetails.vMobileNo = "" & !mobile_no
            GetLastAffectedRowDetails.vEmail = "" & !email
            GetLastAffectedRowDetails.vHospital = "" & !hospital
        Else
            GetLastAffectedRowDetails.vPhysicianID = ""
            GetLastAffectedRowDetails.vMobileNo = ""
            GetLastAffectedRowDetails.vEmail = ""
            GetLastAffectedRowDetails.vHospital = ""
        End If
        .Close
        Set rsLastRow = Nothing
    End With
End Function



















