VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7656
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5976
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    CommandButton5.Caption = "Añadir"
    CommandButton6.Enabled = False
    If CommandButton1.Caption = "Productos" Then
         Call cambioBtn("Productos", "Añadir", "Modificar", "Volver")
    Else
        If CommandButton1.Caption = "Añadir" And Frame1.Caption = "Productos" Then
            Call CambioAñadir(True, False, "Productos", "Añadir")
            TextBox6.Enabled = False
            TextBox6 = numfila() + 1
        Else
            If CommandButton1.Caption = "Añadir" And Frame1.Caption = "Ventas" Then
                Call CambioAñadir(True, False, "Ventas", "Añadir")
                CommandButton7.Enabled = False
                TextBox6.Enabled = False
                TextBox6 = numfilaHoja3() + 1
            End If
        End If
    End If
End Sub

Private Sub CambioAñadir(ByVal mostrarFrame As Boolean, ByVal mostrarBtn As Boolean, ByVal valor As String, ByVal valor1 As String)
    If CommandButton1.Caption = "Añadir" And Frame1.Caption = valor Then
        Frame2.Visible = mostrarFrame
        CommandButton2.Enabled = mostrarBtn
        CommandButton3.Enabled = mostrarBtn
        CommandButton1.Enabled = mostrarBtn
        Frame2.Caption = valor1 + " " + valor
    End If
End Sub

Private Sub cambioBtn(ByVal valor As String, ByVal txtBtn1 As String, ByVal txtBtn2 As String, ByVal txtBtn3 As String)

'esta funcion sirve para cambiar los caption recibe los paramaetros necesarios para cambiar

    If CommandButton1.Caption = valor Then
        CommandButton1.Caption = txtBtn1
        CommandButton2.Caption = txtBtn2
        CommandButton3.Caption = txtBtn3
        Frame1.Caption = valor
    Else
        If CommandButton2.Caption = valor Then
            CommandButton1.Caption = txtBtn1
            CommandButton2.Caption = txtBtn2
            CommandButton3.Caption = "Eliminar"
            CommandButton7.Caption = txtBtn3
            CommandButton7.Visible = True
            Frame1.Caption = valor
        Else
            If CommandButton3.Caption = valor Or CommandButton7.Caption = valor Then
                CommandButton1.Caption = txtBtn1
                CommandButton2.Caption = txtBtn2
                CommandButton3.Caption = txtBtn3
                Frame1.Caption = "Sistema de Ventas"
                CommandButton7.Visible = False
            End If
        End If
        
    End If
End Sub

Private Sub CommandButton2_Click()
    CommandButton5.Caption = "Modificar"
    CommandButton6.Enabled = True
    If CommandButton2.Caption = "Ventas" Then
         Call cambioBtn("Ventas", "Añadir", "Modificar", "Volver")
    Else
        If CommandButton2.Caption = "Modificar" And Frame1.Caption = "Productos" Then
            Call CambioAñadir(True, False, "Productos", "Modificar")
        Else
            If CommandButton2.Caption = "Modificar" And Frame1.Caption = "Ventas" Then
            
                Call CambioAñadir(True, False, "Ventas", "Modificar")
                Label4.Caption = "Cantidad"
                'TextBox2 = "Venta " + ID
                CommandButton7.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub CommandButton3_Click()
    If CommandButton3.Caption = "Eliminar" Then
         Call CambioAñadir(True, False, "Ventas", "Eliminar")
         CommandButton7.Enabled = False
         CommandButton5.Caption = "Eliminar"
         CommandButton6.Enabled = True
    Else
        If CommandButton3.Caption = "Volver" And Frame1.Caption = "Productos" Then
            Call cambioBtn("Volver", "Productos", "Ventas", "Analisis de ventas")
        End If
    End If
End Sub

Private Sub CommandButton4_Click()
    Call limpiar
End Sub

Private Sub limpiar()
    Frame2.Visible = False
    CommandButton2.Enabled = True
    CommandButton3.Enabled = True
    CommandButton1.Enabled = True
    CommandButton7.Enabled = True
    TextBox6.Text = ""
    TextBox2.Text = ""
    TextBox3.Text = ""
    TextBox4.Text = ""
    TextBox5.Text = ""
    TextBox6.Enabled = True
End Sub
Private Sub CommandButton5_Click()
    If CommandButton5.Caption = "Añadir" Then
        Call añadir
    Else
        If CommandButton5.Caption = "Modificar" Then
            Call modificar
        Else
            If CommandButton5.Caption = "Eliminar" Then
                Call eliminar
            End If
        End If
    End If
End Sub

Private Sub añadir()
    Dim fila As Integer
    Dim ID As Integer
    fila = numfila() + 3
    ID = numfila() + 1
    
    If Frame1.Caption = "Productos" Then
        Hoja1.Cells(fila, 1) = ID
        Hoja1.Cells(fila, 2) = TextBox2.Value
        Hoja1.Cells(fila, 3) = TextBox3.Value
        Hoja2.Cells(fila, 3) = TextBox4.Value
        Hoja2.Cells(fila, 2) = ID
        Hoja2.Cells(fila, 1) = ID
        TextBox6 = ID
    Else
        If Frame1.Caption = "Ventas" Then
            Hoja3.Cells(fila, 1) = numfilaHoja3() + 1
            Hoja3.Cells(fila, 2) = TextBox2.Value
            Hoja3.Cells(fila, 3) = TextBox3.Value
            Hoja3.Cells(fila, 3) = TextBox4.Value
            TextBox6 = numfilaHoja3() + 1
        End If
    End If
    Call limpiar
End Sub

Private Sub modificar()
    Dim opcion As Integer
    Dim fila As Integer
    Dim ID As Integer
    fila = Val(TextBox6.Text) + 2
    ID = numfila() + 1
    opcion = MsgBox("¿Esta seguro que quiere modificar los datos?", vbYesNo, "Modificar usuario")
    
    If Frame1.Caption = "Productos" And opcion = 6 Then
            
            Hoja1.Cells(fila, 2) = TextBox2.Value
            Hoja1.Cells(fila, 3) = TextBox3.Value
            Hoja2.Cells(fila, 3) = TextBox4.Value
            
    Else
        If Frame1.Caption = "Ventas" And opcion = 6 Then
            
            Hoja3.Cells(fila, 2) = TextBox2.Value
            Hoja3.Cells(fila, 3) = TextBox3.Value
            Hoja3.Cells(fila, 3) = TextBox4.Value
        End If
    End If
End Sub

Private Sub eliminar()
    If Frame1.Caption = "Ventas" Then
        
    End If
End Sub

Private Sub CommandButton6_Click()
    Call buscar
End Sub

Private Sub CommandButton7_Click()
    Call cambioBtn("Volver", "Productos", "Ventas", "Analisis de ventas")
End Sub

Private Sub buscar()
    Dim numfil As Integer
    Dim numfil2 As Integer
    
    If Frame1.Caption = "Productos" Then
        'LLamamos a la funcion para obtener el numero de filas que no estan vacias'
        numfil = numfila()
        'MsgBox numfil'
        'val()-> cambia de tipo del contenido del textbox5'
        If Val(TextBox6.Text) >= 1 And Val(TextBox6.Text) <= numfil Then
            TextBox2.Text = Hoja1.Cells(Val(TextBox6.Text) + 2, 2)
            TextBox3.Text = Hoja1.Cells(Val(TextBox6.Text) + 2, 3)
            TextBox4.Text = Hoja2.Cells(Val(TextBox6.Text) + 2, 3)
            'TextBox4.Text = Hoja2.Cells(Val(TextBox5.Text) + 2, 5)
         Else
            MsgBox "No se encuentran los datos", vbExclamation, "Error"
         End If
    Else
        If Frame1.Caption = "Ventas" Then
            'LLamamos a la funcion para obtener el numero de filas que no estan vacias'
            numfil2 = numfilaHoja3()
            'MsgBox numfil'
            If Val(TextBox6.Text) >= 1 And Val(TextBox6.Text) <= numfil2 Then
                TextBox6.Text = Hoja3.Cells(Val(TextBox6.Text) + 2, 2)
                TextBox2.Text = Hoja1.Cells(Val(TextBox6.Text) + 2, 2)
                TextBox3.Text = Hoja1.Cells(Val(TextBox6.Text) + 2, 3)
                TextBox4.Text = Hoja3.Cells(Val(TextBox6.Text) + 2, 3)
                'TextBox4.Text = Hoja2.Cells(Val(TextBox5.Text) + 2, 5)
             Else
                MsgBox "No se encuentran los datos", vbExclamation, "Error"
             End If
        End If
    End If
    
    
End Sub

Private Function numfila() As Integer
    'Funcion que cuenta las celdas que estan ocupadas en fila es decir cuantas filas hay'
    Dim i As Integer
    i = 3
    
    Do While Hoja1.Cells(i, 2) <> ""
        i = i + 1
    Loop
    
    numfila = i - 3
End Function

Private Function numfilaHoja3() As Integer
    'Funcion que cuenta las celdas que estan ocupadas en fila es decir cuantas filas hay'
    Dim i As Integer
    i = 3
    
    Do While Hoja3.Cells(i, 2) <> ""
        i = i + 1
    Loop
    
    numfilaHoja3 = i - 3
End Function

