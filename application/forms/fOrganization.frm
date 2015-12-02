VERSION 5.00
Begin VB.Form fOrganization 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   14745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   120
      ScaleHeight     =   8145
      ScaleWidth      =   14505
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   14535
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
         Left            =   10440
         TabIndex        =   0
         Top             =   480
         Width           =   3975
      End
      Begin prjLabSys.LynxGrid lgOrg 
         Height          =   7215
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   12726
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
         Left            =   9360
         TabIndex        =   8
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Organization Management"
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
         TabIndex        =   7
         Top             =   120
         Width           =   2550
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   14400
         Y1              =   360
         Y2              =   360
      End
   End
   Begin prjLabSys.isButton cmdNew 
      Height          =   330
      Left            =   7560
      TabIndex        =   1
      Top             =   8400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      Icon            =   "fOrganization.frx":0000
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
   Begin prjLabSys.isButton cmdEdit 
      Height          =   330
      Left            =   9360
      TabIndex        =   2
      Top             =   8400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      Icon            =   "fOrganization.frx":001C
      Style           =   6
      Caption         =   "&Edit Record"
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
   Begin prjLabSys.isButton cmdDelete 
      Height          =   330
      Left            =   11160
      TabIndex        =   3
      Top             =   8400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      Icon            =   "fOrganization.frx":0038
      Style           =   6
      Caption         =   "&Delete Record"
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
   Begin prjLabSys.isButton cmdClose 
      Height          =   330
      Left            =   12960
      TabIndex        =   4
      Top             =   8400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      Icon            =   "fOrganization.frx":0054
      Style           =   6
      Caption         =   "Cl&ose"
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
End
Attribute VB_Name = "fOrganization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sql As String
Private rsOrg As New ADODB.Recordset

Private Sub cmdDelete_Click()
     With lgOrg
        If .ItemCount > 0 Then
            If .SelectedCount > 0 Then
                If MsgBox("Are you sure you want to delete this organization ?                  " & vbNewLine _
                    & "[ " & .CellText(.row, 1) & "] ", vbYesNo + vbQuestion, system_title) = vbYes Then
                    DeleteOrganization CInt(.CellText(.row, 0))
                End If
            Else
                MsgBox "Please select an item to delete.                      ", vbOKOnly + vbExclamation, system_title
            End If
        Else
            MsgBox "Please select an item to delete.                      ", vbOKOnly + vbExclamation, system_title
        End If
    End With
End Sub

Private Sub cmdEdit_Click()
    With lgOrg
        If .ItemCount > 0 Then
            If .SelectedCount > 0 Then
                fOrganizationEntry.mode = "edit"
                fOrganizationEntry.var_form = "Organization"
                fOrganizationEntry.selected_org_id = CInt(.CellText(.row, 0))
                fOrganizationEntry.txtOrgName.Text = .CellText(.row, 1)
                fOrganizationEntry.txtOrgAddress.Text = .CellText(.row, 2)
                fOrganizationEntry.txtOrgContactPerson.Text = .CellText(.row, 3)
                fOrganizationEntry.txtOrgTelephone.Text = .CellText(.row, 4)
                fOrganizationEntry.txtOrgMobile.Text = .CellText(.row, 5)
                fOrganizationEntry.txtOrgEmail.Text = .CellText(.row, 6)
                fOrganizationEntry.Show vbModal
            Else
                MsgBox "Please select an item to edit.                      ", vbOKOnly + vbExclamation, system_title
            End If
        Else
            MsgBox "Please select an item to edit.                      ", vbOKOnly + vbExclamation, system_title
        End If
    End With
End Sub

Private Sub cmdNew_Click()
    fOrganizationEntry.mode = "add"
    fOrganizationEntry.var_form = "Organization"
    fOrganizationEntry.Show vbModal
End Sub

Private Sub Form_Load()
    Center Me
    InitializeConnectionString
    OpenConnection
    InitializeGridProperty
    CreateHeaders
    LoadOrganization
End Sub

Private Sub InitializeGridProperty()
    With lgOrg
        '.ImageList = ImageList1
        .Redraw = False                                 'do not draw
        .AllowEdit = True
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

Private Sub CreateHeaders()
    With lgOrg
        n = .Width
        .AddColumn "orgID", 0                               '0
        .AddColumn "Organization Name", n * 0.25            '1
        .AddColumn "Organization Address", n * 0.3          '2
        .AddColumn "Contact Person", n * 0.15               '3
        .AddColumn "Telephone #", n * 0.13                  '4
        .AddColumn "Mobile #", n * 0.13                     '5
        .AddColumn "Email Address", n * 0.17                '6
        
        
        .ColLocked(0) = True
        .ColLocked(1) = True
        .ColLocked(2) = True
        .ColLocked(3) = True
        .ColLocked(4) = True
        .ColLocked(5) = True
        .ColLocked(6) = True
        
        .AddRow ""
    End With
End Sub

Private Sub LoadOrganization(Optional search_text As String = "")
    With rsOrg
        lgOrg.Redraw = False
        lgOrg.Clear
        
        sql = "SELECT * FROM organization " _
            & "WHERE is_deleted = 0 and is_active = 1 " _
            & "and org_name LIKE '%" & search_text & "%' " _
            & "ORDER BY org_name "
            
        OpenTable sql, rsOrg
        If .EOF = False Then
            While .EOF = False
                
                lgOrg.AddRow !org_id & vbTab _
                        & !org_name & vbTab _
                        & !org_address & vbTab _
                        & !org_contact_person & vbTab _
                        & !org_telephone & vbTab _
                        & !org_mobile & vbTab _
                        & !org_email
                
                .MoveNext
            Wend
        End If
        lgOrg.Redraw = True
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub DeleteOrganization(org_id As Integer)
    Dim last_id As Integer
    On Error GoTo err_hndler
    dbCon.BeginTrans
    
    sql = "UPDATE organization " _
        & "SET " _
        & "is_deleted = 1," _
        & "deleted_by = " & current_user_id & "," _
        & "deleted_datetime = '" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "' " _
        & "WHERE org_id = " & org_id
    
    dbCon.Execute sql                                                                   '** execute insert statement
    dbCon.CommitTrans
    
    With lgOrg
        .RemoveRow (.row)
    End With
    
    MsgBox "Record has been deleted.                      ", vbOKOnly + vbInformation, system_title
    txtSearch.SetFocus
    Exit Sub
err_hndler:
    dbCon.RollbackTrans
    MsgBox "Unexpected Error Occurred.Saving has been cancelled.                    ", vbOKOnly + vbCritical, system_title
End Sub






































Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LoadOrganization txtSearch.Text
        Exit Sub
    End If
    
    KeyAscii = IntStrInputText(KeyAscii)
End Sub

Private Sub txtSearch_GotFocus()
    BackColorOnFocus txtSearch, vbYellow
End Sub

Private Sub txtSearch_LostFocus()
    BackColorOnFocus txtSearch, vbWhite
End Sub
