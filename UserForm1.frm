VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8664.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6624
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    If CommandButton1.Caption = "Productos" Then
         Call cambioBtn("Productos", "Añadir", "Modificar", "Volver")
    Else
        If CommandButton1.Caption = "Añadir" And Frame1.Caption = "Productos" Then
            Call CambioAñadir(True, False, "Productos", "Añadir")
        Else
            If CommandButton1.Caption = "Añadir" And Frame1.Caption = "Ventas" Then
                Call CambioAñadir(True, False, "Ventas", "Añadir")
                CommandButton7.Enabled = False
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
    If CommandButton2.Caption = "Ventas" Then
         Call cambioBtn("Ventas", "Añadir", "Modificar", "Volver")
    Else
        If CommandButton2.Caption = "Modificar" And Frame1.Caption = "Productos" Then
            Call CambioAñadir(True, False, "Productos", "Modificar")
        Else
            If CommandButton2.Caption = "Modificar" And Frame1.Caption = "Ventas" Then
                Call CambioAñadir(True, False, "Ventas", "Modificar")
                CommandButton7.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub CommandButton3_Click()
    If CommandButton3.Caption = "Eliminar" Then
         Call CambioAñadir(True, False, "Ventas", "Eliminar")
         CommandButton7.Enabled = False
    Else
        If CommandButton3.Caption = "Volver" And Frame1.Caption = "Productos" Then
            Call cambioBtn("Volver", "Productos", "Ventas", "Analisis de ventas")
        End If
    End If
    
        
End Sub

Private Sub CommandButton4_Click()
    Frame2.Visible = False
    CommandButton2.Enabled = True
    CommandButton3.Enabled = True
    CommandButton1.Enabled = True
    CommandButton7.Enabled = True
End Sub

Private Sub CommandButton7_Click()
    Call cambioBtn("Volver", "Productos", "Ventas", "Analisis de ventas")
End Sub

Private Sub Frame1_Click()

End Sub
