VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4005
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3735
      Begin VB.TextBox txtIP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E7E7E7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   0
         Top             =   300
         Width           =   1500
      End
      Begin VB.TextBox txtEsquema 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E7E7E7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   1
         Top             =   700
         Width           =   1500
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1920
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmLogin.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "F10   Salir"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtUsuario 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E7E7E7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1100
         Width           =   1500
      End
      Begin VB.TextBox txtPass 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E7E7E7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1500
         Width           =   1500
      End
      Begin VB.CommandButton cmdaceptar 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   360
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmLogin.frx":1054
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "F10   Salir"
         Top             =   2160
         Width           =   1350
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION IP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   350
         Width           =   1185
      End
      Begin VB.Label lblBaseDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "ESQUEMA (DB):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   255
         TabIndex        =   9
         Top             =   750
         Width           =   1305
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTRASEÑA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   255
         TabIndex        =   8
         Top             =   1550
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   255
         TabIndex        =   7
         Top             =   1150
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' EJEMPLO DE CODIGO PARA CONECTARSE CON UN SERVIDOR ISERIES EN SU VERSION 5.2.0 EN ADELANTE
' DEBE ALIMENTAR LOS SIGUIENTES VALORES
' DIRECCION IP
' ESQUEMA O NOMBRE DE LA BASE DE DATOS
' USUARIO
' CONTRASEÑA

' PUEDE EJECUTAR UNA CONSULTA BASICA, NO PODRA HACER UPDATES NI DELETES


'*************************
'* author:                *
'* philemon_@hotmail.com *
'*************************

Private Sub cmdaceptar_Click()
Dim pb_flag As Boolean
    
On Error GoTo getThe_Error
    
    If Conectar = 1 Then
        frmQuery.Show
        Unload Me
    Else
        MsgBox "Realice los cambios necesarios"
    End If

Exit Sub
getThe_Error:
    MsgBox "Se ha producido el siguiente error: " & Err.Description
    Err.Clear
End Sub

Private Sub cmdSalir_Click()
    End
End Sub


'**** VALIDACIONES
Private Sub txtIP_Validate(Cancel As Boolean)
    strDirIP = Trim$(txtIP.Text)
End Sub

Private Sub txtEsquema_Validate(Cancel As Boolean)
    strEsquema = Trim$(txtEsquema.Text)
End Sub

Private Sub txtUsuario_Validate(Cancel As Boolean)
    strUsuario = Trim$(txtUsuario.Text)
End Sub

Private Sub txtPass_Validate(Cancel As Boolean)
    strPassword = Trim$(txtPass.Text)
End Sub

