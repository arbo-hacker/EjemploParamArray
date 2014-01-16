VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form2"
   ScaleHeight     =   4785
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMDEliminar2 
      Caption         =   "E&liminar cedulas mayores o iguales a 18000000"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   3735
   End
   Begin VB.CommandButton CMDModificar2 
      Caption         =   "M&odificar cedulas mayores o iguales a 18000000"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   4320
      Width           =   3855
   End
   Begin VB.CommandButton CMDSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton CMDModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton CMDEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton CMDAñadir 
      Caption         =   "&Añadir"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "T1"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TXTCedula"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TXTNombre"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TXTSexo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TXTNotaf"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "MSFlexGrid1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "T2"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "TXTCedula2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "TXTNota"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "TXTCorte"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "MSFlexGrid2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2055
         Left            =   -74520
         TabIndex        =   22
         Top             =   1440
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3625
         _Version        =   393216
      End
      Begin VB.TextBox TXTCorte 
         Height          =   375
         Left            =   -69000
         MaxLength       =   1
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TXTNota 
         Height          =   375
         Left            =   -71760
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TXTCedula2 
         Height          =   375
         Left            =   -74520
         MaxLength       =   8
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2055
         Left            =   480
         TabIndex        =   18
         Top             =   1440
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3625
         _Version        =   393216
      End
      Begin VB.TextBox TXTNotaf 
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TXTSexo 
         Height          =   375
         Left            =   6000
         MaxLength       =   1
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TXTNombre 
         Height          =   375
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TXTCedula 
         Height          =   375
         Left            =   480
         MaxLength       =   8
         TabIndex        =   0
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Corte"
         Height          =   495
         Left            =   -69000
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Nota"
         Height          =   495
         Left            =   -71760
         TabIndex        =   20
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Cedula"
         Height          =   375
         Left            =   -74520
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Sexo"
         Height          =   495
         Left            =   6000
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   375
         Left            =   4200
         TabIndex        =   16
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Nota final"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cedula"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RS As ADODB.Recordset
Public Sql As ADODB.Command
Public DB As ADODB.Connection
Public Sub AbrirBD()
Set Sql = New ADODB.Command
Set DB = New ADODB.Connection
Set RS = New ADODB.Recordset
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\aem.mdb;Persist Security Info=False"
Sql.ActiveConnection = DB
End Sub
Private Sub CMDAñadir_Click()
Select Case SSTab1.Tab
    Case 0
        If TXTCedula.Text = "" Or TXTNotaf.Text = "" Or TXTNombre.Text = "" Or TXTSexo.Text = "" Then
            MsgBox "Debe llenar todos los datos", vbQuestion, "Información"
            Exit Sub
        Else
            Sql.CommandText = Añadir("T1", TXTCedula.Text, "#" & TXTNotaf.Text, TXTNombre.Text, TXTSexo.Text)
            Set RS = Sql.Execute
            Llenar_Flexgrid
        End If
    Case 1
        If TXTCedula2.Text = "" Or TXTNota.Text = "" Or TXTCorte.Text = "" Then
            MsgBox "Debe llenar todos los datos", vbQuestion, "Información"
            Exit Sub
        Else
            Sql.CommandText = Añadir("T2", TXTCedula2.Text, "#" & TXTNota.Text, "#" & TXTCorte.Text)
            Set RS = Sql.Execute
            Llenar_Flexgrid2
        End If
End Select
End Sub
Private Sub CMDEliminar_Click()
    Select Case SSTab1.Tab
        Case 0
            If TXTCedula.Text = "" Then
                MsgBox "Debe escribir la cedula de la persona que desea eliminar", vbQuestion, "Información"
                Exit Sub
            Else
                Sql.CommandText = Eliminar("T1", "CI", TXTCedula.Text)
                Set RS = Sql.Execute
                Llenar_Flexgrid
            End If
        Case 1
            If TXTCedula2.Text = "" Then
                MsgBox "Debe escribir la cedula de la persona que desea eliminar", vbQuestion, "Información"
                Exit Sub
            Else
                Sql.CommandText = Eliminar("T2", "CI", TXTCedula2.Text)
                Set RS = Sql.Execute
                Llenar_Flexgrid2
            End If
    End Select
End Sub
Private Sub CMDEliminar2_Click()
    Select Case SSTab1.Tab
        Case 0
            If TXTCedula.Text = "" Then
                MsgBox "Debe escribir la cedula de la persona que desea eliminar", vbQuestion, "Información"
                Exit Sub
            Else
                Sql.CommandText = Eliminar("T1", "CI", TXTCedula.Text, "and ci>='18000000'")
                Set RS = Sql.Execute
                Llenar_Flexgrid
            End If
        Case 1
            If TXTCedula2.Text = "" Then
                MsgBox "Debe escribir la cedula de la persona que desea eliminar", vbQuestion, "Información"
                Exit Sub
            Else
                Sql.CommandText = Eliminar("T2", "CI", TXTCedula2.Text, "and ci>='18000000'")
                Set RS = Sql.Execute
                Llenar_Flexgrid2
            End If
    End Select
End Sub
Private Sub CMDModificar_Click()
    Select Case SSTab1.Tab
        Case 0
            If TXTCedula.Text = "" Or TXTNotaf.Text = "" Or TXTNombre.Text = "" Or TXTSexo.Text = "" Then
                MsgBox "Debe llenar todos los datos", vbQuestion, "Información"
                Exit Sub
            Else
                Sql.CommandText = Modificar("T1", "#" & TXTNotaf.Text, TXTNombre.Text, TXTSexo.Text, TXTCedula.Text)
                Set RS = Sql.Execute
                Llenar_Flexgrid
            End If
        Case 1
            If TXTCedula2.Text = "" Or TXTNota.Text = "" Or TXTCorte.Text = "" Then
                MsgBox "Debe llenar todos los datos", vbQuestion, "Información"
                Exit Sub
            Else
                Sql.CommandText = Modificar("T2", "#" & TXTNota.Text, "#" & TXTCorte.Text, TXTCedula2.Text)
                Set RS = Sql.Execute
                Llenar_Flexgrid2
            End If
    End Select
End Sub
Private Sub CMDModificar2_Click()
    Select Case SSTab1.Tab
        Case 0
            If TXTCedula.Text = "" Or TXTNotaf.Text = "" Or TXTNombre.Text = "" Or TXTSexo.Text = "" Then
                MsgBox "Debe llenar todos los datos", vbQuestion, "Información"
                Exit Sub
            Else
                Sql.CommandText = Modificar("T1", "#" & TXTNotaf.Text, TXTNombre.Text, TXTSexo.Text, TXTCedula.Text, "and ci>='18000000'")
                Set RS = Sql.Execute
                Llenar_Flexgrid
            End If
        Case 1
            If TXTCedula2.Text = "" Or TXTNota.Text = "" Or TXTCorte.Text = "" Then
                MsgBox "Debe llenar todos los datos", vbQuestion, "Información"
                Exit Sub
            Else
                Sql.CommandText = Modificar("T2", "#" & TXTNota.Text, "#" & TXTCorte.Text, TXTCedula2.Text, "and ci>='18000000'")
                Set RS = Sql.Execute
                Llenar_Flexgrid2
            End If
    End Select
End Sub
Private Sub CMDSalir_Click()
If MsgBox("Desea salir del sistema", vbQuestion + vbYesNo, "Confirme") = vbYes Then End
End Sub
Private Function Añadir(ParamArray Elementos() As Variant) As String
        If RS.State = 1 Then RS.Close 'Si rs esta abierto, es decir, acaba de ser usado lo cierro para poder usarlo de nuevo
        RS.Open "select * from " & Elementos(0), DB, adOpenDynamic, adLockOptimistic 'Abro la tabla que esta escrita en el primer valor del array de elementos
        Dim cadena As String 'Aqui guardo los nombres de todos los campos
        Dim Campo As ADODB.Field 'Esta variable se usa para sacar los nombres de los campos
            For Each Campo In RS.Fields 'Recorrer todos los nombres de los campos en la tabla y darle el nombre a la variable campo
                cadena = cadena & Campo.Name & "," 'Aqui armo los nombres de los campos. Cadena es igual a cadena mas el nombre del campo
            Next 'Cierro el ciclo del for
            
Dim Elem As String 'Aqui construyo los valores que se le daran a los campos
Dim i As Integer 'Declaro la variable que me dara el valor de elementos

        For i = 1 To UBound(Elementos) 'Hacer desde el segundo elemento del array Elementos hasta el ultimo valor de Elementos
            If Mid(Elementos(i), 1, 1) = "#" Then 'Si la primera letra de lo que vale elementos es # entonces es un valor numerico
                If Elem = "" Then 'Si elem es igual a blanco entonces
                    Elem = "" & Mid(Elementos(i), 2, Len(Elementos(i))) & "" 'Elem es igual a lo que hay en elementos a partir de la segunda letra
                Else 'De lo contrario
                    Elem = Elem & "," & Mid(Elementos(i), 2, Len(Elementos(i))) & "" 'Elem es igual a elem mas lo que hay en elementos a partir de la segunda letra
                End If 'Cierro el if
            Else 'De lo contrario elementos es un string
                If Elem = "" Then 'Si elem es igual a blanco entonces
                    Elem = "'" & Elementos(i) & "'" 'Elem es igual a lo que hay en la variable elementos
                Else 'De lo contrario
                    Elem = Elem & ",'" & Elementos(i) & "'" 'Elem es igual a elem mas lo que hay en la variable elementos
                End If 'Cierro el if
            End If 'Cierro el if
        Next 'Cierro el ciclo del for
Añadir = "insert into " & Elementos(0) & "(" & Mid$(cadena, 1, Len(cadena) - 1) & ") values(" & Elem & ")" 'Aqui contruyo la cadena sql
End Function
Private Sub Form_Load()
AbrirBD
Call Cabecera
Call Cabecera2
Call Llenar_Flexgrid
Call Llenar_Flexgrid2
End Sub
Private Function Eliminar(ParamArray Elem() As Variant) As String
    If Mid(Elem(2), 1, 1) = "#" Then
        Elem(2) = Mid(Elem(2), 2, Len(Elem(2)))
    Else
        Elem(2) = "'" & Elem(2) & "'"
    End If
On Error GoTo elem3
        If Elem(3) <> "" Then
            Eliminar = "delete from " & Elem(0) & " where " & Elem(1) & "=" & Elem(2) & " " & Elem(3)
        End If
elem3:
        If Err.Number = 9 Then
            Eliminar = "delete from " & Elem(0) & " where " & Elem(1) & "=" & Elem(2)
        End If
End Function
Private Function Modificar(ParamArray Elementos() As Variant) As String
Dim Elem As String
Dim Campo As Field
Dim i As Integer
Dim Campo_clave As String
If RS.State = 1 Then RS.Close 'Si la tabla esta abierta la cierro
RS.Open "Select * from " & Elementos(0), DB, adOpenDynamic 'Abro la tabla cuyo nombre es igual a elementos(0)
    For Each Campo In RS.Fields 'Con este ciclo obtengo los nombres de todos los campos de una tabla
        If i >= 1 Then 'Si i es mayor o igual que 1 entonces
            If Mid(Elementos(i), 1, 1) = "#" Then 'Si la primera letra de lo que vale elementos es # entonces es un valor numerico
                If Elem = "" Then 'Si elem es igual a blanco entonces
                    Elem = Campo.Name & "=" & Mid(Elementos(i), 2, Len(Elementos(i)))  'Elem es igual a lo que hay en elementos a partir de la segunda letra
                Else 'De lo contrario
                    Elem = Elem & ", " & Campo.Name & "=" & Mid(Elementos(i), 2, Len(Elementos(i))) 'Elem es igual a elem mas lo que hay en elementos a partir de la segunda letra
                End If 'Cierro el if
            Else 'De lo contrario elementos es un string
                If Elem = "" Then 'Si elem es igual a blanco entonces
                    Elem = Campo.Name & "='" & Elementos(i) & "'"  'Elem es igual a lo que hay en la variable elementos
                Else 'De lo contrario
                    Elem = Elem & ", " & Campo.Name & "=" & "'" & Elementos(i) & "'" 'Elem es igual a elem mas lo que hay en la variable elementos
                End If 'Cierro el if
            End If 'Cierro el if
        Else 'De lo contrario
        Campo_clave = Campo.Name 'Consigo el nombre del campo clave para poder modificar un registro
        End If 'Cierro el if
        i = i + 1
    Next 'Cierro el ciclo del for
Dim Cual As String 'Determina si el valor del campo_clave es numerico o un srting
            If Mid(Elementos(i), 1, 1) = "#" Then 'Si la primera letra de lo que vale elementos es # entonces es un valor numerico
                Cual = Mid(Elementos(i), 2, Len(Elementos(i))) 'Cual es igual a lo que vale elementos a partir de la segunda letra
            Else 'De lo contrario
                Cual = "'" & Elementos(i) & "'" 'Elementos es un valor string y por eso Cual es igual a lo que vale elementos encerrado entre apostrofes
            End If 'Cierro el If
On Error GoTo Ubound_Elementos 'Si ocurre un error ir a donde dice Ubound_Elementos
                    If Elementos(i + 1) <> "" Then 'Si elementos es diferente de blanco entonces
                        Modificar = "Update " & Elementos(0) & " set " & Elem & " where " & Campo_clave & "=" & Cual & " " & Elementos(i + 1) 'Armo la cadena sql y lo concateno con el ultimo valor de elementos
                    End If 'Cierro el If
Ubound_Elementos: 'Si da error salta hasta aqui
                    If Err.Number = 9 Then 'Si ocurre el error numero 9 entonces
                        Modificar = "Update " & Elementos(0) & " set " & Elem & " where " & Campo_clave & "=" & Cual 'Armo la cadena sql
                    End If 'Cierro el If
            
End Function
Private Sub Llenar_Flexgrid2()
Dim i As Integer
MSFlexGrid2.Refresh
        If RS.State = 1 Then RS.Close
        RS.Open "select count(ci) as var from t2", DB, adOpenDynamic, adLockOptimistic
        MSFlexGrid2.Rows = RS!Var + 1
    If RS.State = 1 Then RS.Close
        RS.Open "select * from t2 order by ci asc", DB, adOpenDynamic, adLockOptimistic
        If RS.EOF = True Then Exit Sub
        RS.MoveFirst
            With MSFlexGrid2
                Do Until RS.EOF
                    i = i + 1
                    .Row = i
                    .Col = 1
                    .Text = RS!Ci
                    .Col = 2
                    .Text = RS!nota
                    .Col = 3
                    .Text = RS!corte
                    RS.MoveNext
                Loop
            End With
End Sub
Private Sub Llenar_Flexgrid()
Dim i As Integer
MSFlexGrid1.Refresh

        If RS.State = 1 Then RS.Close
        RS.Open "select count(ci) as var from t1", DB, adOpenDynamic, adLockOptimistic
        MSFlexGrid1.Rows = RS!Var + 1
    If RS.State = 1 Then RS.Close
        RS.Open "select * from t1 order by ci asc", DB, adOpenDynamic, adLockOptimistic
        If RS.EOF = True Then Exit Sub
        RS.MoveFirst
            Do Until RS.EOF
                i = i + 1
                MSFlexGrid1.Row = i
                MSFlexGrid1.Col = 1
                MSFlexGrid1.Text = RS!Ci
                MSFlexGrid1.Col = 2
                MSFlexGrid1.Text = RS!Notaf
                MSFlexGrid1.Col = 3
                MSFlexGrid1.Text = RS!Nombre
                MSFlexGrid1.Col = 4
                MSFlexGrid1.Text = RS!Sexo
                RS.MoveNext
            Loop
End Sub
Private Sub TXTCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    MsgBox "Debe escribir solo numeros", vbQuestion, "Información"
    KeyAscii = 0
End If
End Sub
Private Sub TXTCedula2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    MsgBox "Debe escribir solo numeros", vbQuestion, "Información"
    KeyAscii = 0
End If
End Sub
Private Sub TXTCorte_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    MsgBox "Debe escribir solo numeros", vbQuestion, "Información"
    KeyAscii = 0
End If
End Sub
Private Sub TXTNota_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    MsgBox "Debe escribir solo numeros", vbQuestion, "Información"
    KeyAscii = 0
End If
End Sub
Private Sub TXTNotaf_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    MsgBox "Debe escribir solo numeros", vbQuestion, "Información"
    KeyAscii = 0
End If
End Sub
Private Sub TXTSexo_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii <> Asc("f") And KeyAscii <> Asc("F") And KeyAscii <> Asc("m") And KeyAscii <> Asc("M") Then
MsgBox "Solo puede escribir sexo F o M", vbInformation, "Información"
KeyAscii = 0
End If
End Sub
Private Sub Cabecera()
Dim i As Integer
Dim x As String
MSFlexGrid1.Cols = 5
If RS.State = 1 Then RS.Close
RS.Open "select count(ci) as var from t1", DB, adOpenDynamic, adLockOptimistic
MSFlexGrid1.Rows = RS!Var + 1
MSFlexGrid1.Row = 0

    For i = 1 To 4
        
        Select Case i
            Case 1
                x = "Cedula"
            Case 2
                x = "Nota Final"
            Case 4
                x = "Nombre"
            Case 3
                x = "Sexo"
        End Select

                MSFlexGrid1.Col = i
                MSFlexGrid1.Text = x
    Next
End Sub
Private Sub Cabecera2()
Dim i As Integer
Dim x As String
If RS.State = 1 Then RS.Close
RS.Open "select count(ci) as var from t2", DB, adOpenDynamic, adLockOptimistic
MSFlexGrid2.Cols = 4
MSFlexGrid2.Rows = RS!Var + 1
MSFlexGrid2.Row = 0

    For i = 1 To 3
        
        Select Case i
            Case 1
                x = "Cedula"
            Case 2
                x = "Nota"
            Case 3
                x = "Corte"
        End Select

                MSFlexGrid2.Col = i
                MSFlexGrid2.Text = x
    Next
End Sub
