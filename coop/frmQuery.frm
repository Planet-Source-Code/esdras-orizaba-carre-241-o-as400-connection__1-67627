VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmQuery 
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   Icon            =   "frmQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Height          =   375
      Left            =   4560
      Picture         =   "frmQuery.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid msfResults 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2778
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   4680
      Picture         =   "frmQuery.frx":1054
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtQuery 
      Height          =   855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Teclea una consulta a ejecutar (solo consultas):"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strQuery As String

Private Sub cmdQuery_Click()
    Dim n As Integer, m As Long, i As Integer, j As Integer

    
On Error GoTo getThe_Error

    Set rstResults = getRecordset(strQuery)
    
    If rstResults.EOF = True Then
        MsgBox "La consulta no encontro créditos", vbInformation
    Else
        With rstResults
            n = .Fields.Count
            m = .RecordCount
            
            msfResults.Cols = n
            For i = 0 To (n - 1)
                msfResults.TextMatrix(0, i) = .Fields(i).Name
            Next
            .MoveFirst
            For j = 1 To (m)
                For i = 0 To (n - 1)
                    msfResults.TextMatrix(j, i) = .Fields(i).Value
                Next
                msfResults.AddItem ""
                .MoveNext
            Next
        End With
    End If
    
Exit Sub
getThe_Error:
    Set rstResults = Nothing
    MsgBox "Se ha producido el siguiente error: " & Err.Description
    Err.Clear
End Sub

Private Sub cmdSalir_Click()
    End
End Sub

Private Sub txtQuery_Validate(Cancel As Boolean)
    strQuery = txtQuery
End Sub


