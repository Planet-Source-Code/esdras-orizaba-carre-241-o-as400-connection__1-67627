Attribute VB_Name = "mdlConexion"
Option Explicit
Global objConx As New ADODB.Connection
Global strDirIP As String
Global strEsquema As String
Global strUsuario As String
Global strPassword As String
Global rstResults As ADODB.Recordset

Public Function Conectar() As Long
  Dim strConx As String

On Error GoTo getThe_Error
    
    strConx = ""
    strConx = strConx & "Provider=MSDataShape.1;DRIVER=iSeries Access ODBC Driver;"
    strConx = strConx & "SYSTEM=" & strDirIP & ";UID=" & strUsuario & ";PWD=" & strPassword & ";"
    strConx = strConx & "DBQ=" & strEsquema & ";DFTPKGLIB=;XLATEDLL=;LANGUAGEID=ENU;"
    strConx = strConx & "DBQ=" & strEsquema & ";SORTTABLE=;PKG=" & strEsquema & "/DEFAULT(IBM),2,0,1,0,6000;"
    strConx = strConx & "QAQQINILIB=;'DATABASE=;SQDIAGCODE=;CMT=0;DFT=6;DSP=2;TFT=3;DEBUG=64;"

    objConx.Open (strConx)
    If objConx.State <> 1 Then
        MsgBox "La conexion ha fallado", vbInformation, "CGV Sistemas"
        Set objConx = Nothing
        Conectar = 0
    Else
        MsgBox "Conexion exitosa... ", vbInformation, "CGV Sistemas"
        Conectar = 1
    End If

Exit Function
getThe_Error:
If Err.Number <> 0 Then
    MsgBox "Se ha producido el siguiente error: " & Err.Description
    Conectar = 0
    Set objConx = Nothing
    Err.Clear
End If

End Function

Public Function ConxSts() As Boolean
On Error GoTo getThe_Error
    
    If objConx.State = 1 Then
        ConxSts = True
    Else
        ConxSts = False
    End If

Exit Function
getThe_Error:
    MsgBox "Se ha producido el siguiente error: " & Err.Description
    Err.Clear
End Function

Public Function getRecordset(ByVal pvsQuery As String) As ADODB.Recordset

On Error GoTo getThe_Error
        
    If Not ConxSts Then
        If Not Conectar Then
            Err.Raise -1, "ObtenRecordset", "No se pudo reconectar a la BD"
        End If
    End If
    Set getRecordset = New ADODB.Recordset
    getRecordset.CursorLocation = adUseClient
    objConx.CommandTimeout = 1000
    Set getRecordset = objConx.Execute(pvsQuery)
    
Exit Function
getThe_Error:
    Set getRecordset = Nothing
    MsgBox "Se ha producido el siguiente error: " & Err.Description
    Err.Clear
End Function






