VERSION 5.00
Begin VB.Form fOrganizationEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3945
      ScaleWidth      =   6465
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Width           =   6495
      Begin VB.TextBox txtOrgEmail 
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
         Left            =   1920
         TabIndex        =   5
         Top             =   3480
         Width           =   4335
      End
      Begin VB.TextBox txtOrgMobile 
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
         Left            =   1920
         TabIndex        =   4
         Top             =   3120
         Width           =   4335
      End
      Begin VB.TextBox txtOrgTelephone 
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
         Left            =   1920
         TabIndex        =   3
         Top             =   2760
         Width           =   4335
      End
      Begin VB.TextBox txtOrgContactPerson 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   2400
         Width           =   4335
      End
      Begin VB.TextBox txtOrgAddress 
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
         Height          =   765
         Left            =   1920
         TabIndex        =   1
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox txtOrgName 
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
         Left            =   1920
         TabIndex        =   0
         Top             =   600
         Width           =   4335
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   6360
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Organization Contact"
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
         Top             =   1920
         Width           =   1785
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Email Address# :"
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
         TabIndex        =   17
         Top             =   3480
         Width           =   1605
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mobile # :"
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
         Top             =   3120
         Width           =   1605
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Telephone # :"
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
         TabIndex        =   15
         Top             =   2760
         Width           =   1605
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contact Person :"
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
         Top             =   2400
         Width           =   1605
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1605
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6360
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Organization Details"
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
         Top             =   120
         Width           =   1710
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "* Organization Name :"
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
         Top             =   600
         Width           =   1605
      End
   End
   Begin prjLabSys.isButton cmdSave 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   4560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Icon            =   "fOrganizationEntry.frx":0000
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
   Begin prjLabSys.isButton cmdClose 
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "fOrganizationEntry.frx":001C
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
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Organization Adding/Editing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "> Please provide all required information."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   2880
      TabIndex        =   8
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "fOrganizationEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sql As String
Public mode As String
Public selected_org_id As Integer
Public var_form As String



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Center Me
    InitializeConnectionString
    OpenConnection

End Sub

Private Function ValidateFields() As Boolean
    
    If Trim(txtOrgName.Text) = "" Then
        ValidateFields = False
        MsgBox "Organization Name is required.                      ", vbOKOnly + vbExclamation, system_title
        txtOrgName.SetFocus
        Exit Function
    End If
    
    ValidateFields = True
    
End Function

Private Sub ClearControls()
    txtOrgAddress.Text = ""
    txtOrgName.Text = ""
    txtOrgContactPerson.Text = ""
    txtOrgTelephone.Text = ""
    txtOrgMobile.Text = ""
    txtOrgEmail.Text = ""
End Sub

Private Sub cmdSave_Click()
    If ValidateFields = True Then
        If mode = "add" Then
            InsertOrganization
        ElseIf mode = "edit" Then
            UpdateOrganization
        End If
    End If
End Sub

Private Sub InsertOrganization()
Dim last_id As Integer
    On Error GoTo err_hndler
    dbCon.BeginTrans
    
    sql = "INSERT INTO organization " _
        & "SET " _
            & "org_name = '" & Trim(txtOrgName.Text) & "'," _
            & "org_address = '" & Trim(txtOrgAddress.Text) & "'," _
            & "org_contact_person = '" & Trim(txtOrgContactPerson.Text) & "'," _
            & "org_telephone = '" & Trim(txtOrgTelephone.Text) & "'," _
            & "org_mobile = '" & Trim(txtOrgMobile.Text) & "'," _
            & "org_email = '" & Trim(txtOrgEmail.Text) & "'," _
            & "created_by = " & current_user_id & "," _
            & "created_datetime = '" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "' "
    
    dbCon.Execute sql                                                                   '** execute insert statement
    
    last_id = GetLastInsertedId("organization", dbCon)
    
    dbCon.CommitTrans
    
    If var_form = "Organization" Then
        With fOrganization.lgOrg
            .Redraw = False
            .AddRow last_id & vbTab _
                & Trim(txtOrgName.Text) & vbTab _
                & Trim(txtOrgAddress.Text) & vbTab _
                & Trim(txtOrgContactPerson.Text) & vbTab _
                & Trim(txtOrgTelephone.Text) & vbTab _
                & Trim(txtOrgMobile.Text) & vbTab _
                & Trim(txtOrgEmail.Text)
            .Redraw = True
        End With
    End If
    
    MsgBox "Record has been saved.                      ", vbOKOnly + vbInformation, system_title
    ClearControls
    txtOrgName.SetFocus
    Exit Sub
err_hndler:
    dbCon.RollbackTrans
    MsgBox "Unexpected Error Occurred.Saving has been cancelled.                    ", vbOKOnly + vbCritical, system_title
    MsgBox Error
End Sub

Private Sub UpdateOrganization()
    On Error GoTo err_hndler
    dbCon.BeginTrans
    
    sql = "UPDATE organization " _
        & "SET " _
            & "org_name = '" & Trim(txtOrgName.Text) & "'," _
            & "org_address = '" & Trim(txtOrgAddress.Text) & "'," _
            & "org_contact_person = '" & Trim(txtOrgContactPerson.Text) & "'," _
            & "org_telephone = '" & Trim(txtOrgTelephone.Text) & "'," _
            & "org_mobile = '" & Trim(txtOrgMobile.Text) & "'," _
            & "org_email = '" & Trim(txtOrgEmail.Text) & "'," _
            & "created_by = " & current_user_id & "," _
            & "created_datetime = '" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "' " _
        & "WHERE " _
        & "org_id = " & selected_org_id
    
    dbCon.Execute sql                                                                   '** execute update statement

    dbCon.CommitTrans
    
    
    If var_form = "Organization" Then
        With fOrganization.lgOrg
            .Redraw = False
            .CellText(.row, 1) = Trim(txtOrgName.Text)
            .CellText(.row, 2) = Trim(txtOrgAddress.Text)
            .CellText(.row, 3) = Trim(txtOrgContactPerson.Text)
            .CellText(.row, 4) = Trim(txtOrgTelephone.Text)
            .CellText(.row, 5) = Trim(txtOrgMobile.Text)
            .CellText(.row, 6) = Trim(txtOrgEmail.Text)
            .Redraw = True
        End With
    End If
    
    MsgBox "Record has been saved.                      ", vbOKOnly + vbInformation, system_title
    Unload Me
    Exit Sub
err_hndler:
    dbCon.RollbackTrans
    MsgBox "Unexpected Error Occurred.Saving has been cancelled.                    ", vbOKOnly + vbCritical, system_title
    MsgBox Error
End Sub










































































'** start form object lost focus events ######################################################################

Private Sub txtOrgName_LostFocus()
    BackColorOnFocus txtOrgName, vbWhite
End Sub

Private Sub txtOrgAddress_LostFocus()
    BackColorOnFocus txtOrgAddress, vbWhite
End Sub

Private Sub txtOrgContactPerson_LostFocus()
    BackColorOnFocus txtOrgContactPerson, vbWhite
End Sub

Private Sub txtOrgTelephone_LostFocus()
    BackColorOnFocus txtOrgTelephone, vbWhite
End Sub

Private Sub txtOrgMobile_LostFocus()
    BackColorOnFocus txtOrgMobile, vbWhite
End Sub

Private Sub txtOrgEmail_LostFocus()
    BackColorOnFocus txtOrgEmail, vbWhite
End Sub

'** end form object lost focus events #######################################################################



'** start form object got focus events ######################################################################

Private Sub txtOrgName_GotFocus()
    BackColorOnFocus txtOrgName, vbYellow
End Sub

Private Sub txtOrgAddress_GotFocus()
    BackColorOnFocus txtOrgAddress, vbYellow
End Sub

Private Sub txtOrgContactPerson_GotFocus()
    BackColorOnFocus txtOrgContactPerson, vbYellow
End Sub

Private Sub txtOrgTelephone_GotFocus()
    BackColorOnFocus txtOrgTelephone, vbYellow
End Sub

Private Sub txtOrgMobile_GotFocus()
    BackColorOnFocus txtOrgMobile, vbYellow
End Sub

Private Sub txtOrgEmail_GotFocus()
    BackColorOnFocus txtOrgEmail, vbYellow
End Sub

'** end form object got focus events #######################################################################

