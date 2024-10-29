Attribute VB_Name = "ModuloInicioBallota"
Option Explicit

Public Sub CerrarFormulario(frm As UserForm)
    Unload frm
    Set frm = Nothing

End Sub

Public Sub LlenarComboBoxBalota(cbo As ComboBox)
    Dim i As Integer
    For i = 1 To 16
        cbo.AddItem i
    Next i

End Sub

Public Sub LlenarComboBoxElegirNumero(cbo As ComboBox)
    Dim i As Integer
    For i = 1 To 43
        cbo.AddItem i
    Next i

End Sub

Public Sub SugerirNumeroBalota(frm As UserForm)
    Dim numeroAleatorio As Integer
    numeroAleatorio = Int((16 - 1 + 1) * Rnd + 1) ' Genera un número entre 1 y 16
    frm.ComboBox_ElijaBalota.Value = numeroAleatorio

End Sub
 
Public Sub SugerirNumerosUnicos(frm As UserForm)
    Dim numerosElegidos As Collection
    Set numerosElegidos = New Collection
    
    Dim i As Integer
    Dim numeroAleatorio As Integer

    ' Generar 6 números aleatorios únicos del 1 al 43
    Do While numerosElegidos.Count < 6
        numeroAleatorio = Int((43 - 1 + 1) * Rnd + 1) ' Genera un número entre 1 y 43
        On Error Resume Next
        numerosElegidos.Add numeroAleatorio, CStr(numeroAleatorio)
        On Error GoTo 0
    Loop

    ' Asignar los números a cada ComboBox en el formulario pasado como parámetro
    frm.ComboBox_Numero1.Value = numerosElegidos(1)
    frm.ComboBox_Numero2.Value = numerosElegidos(2)
    frm.ComboBox_Numero3.Value = numerosElegidos(3)
    frm.ComboBox_Numero4.Value = numerosElegidos(4)
    frm.ComboBox_Numero5.Value = numerosElegidos(5)
    frm.ComboBox_Numero6.Value = numerosElegidos(6)

End Sub

Public Sub SugerirNumerosUnicosGanadores(frm As UserForm)
    Dim numerosElegidos As Collection
    Set numerosElegidos = New Collection
    
    Dim i As Integer
    Dim numeroAleatorio As Integer

    ' Generar 6 números aleatorios únicos del 1 al 43
    Do While numerosElegidos.Count < 6
        numeroAleatorio = Int((43 - 1 + 1) * Rnd + 1) ' Genera un número entre 1 y 43
        On Error Resume Next
        numerosElegidos.Add numeroAleatorio, CStr(numeroAleatorio)
        On Error GoTo 0
    Loop

    ' Asignar los números a cada ComboBox en el formulario pasado como parámetro
    frm.TextBox1.Value = numerosElegidos(1)
    frm.TextBox2.Value = numerosElegidos(2)
    frm.TextBox3.Value = numerosElegidos(3)
    frm.TextBox4.Value = numerosElegidos(4)
    frm.TextBox5.Value = numerosElegidos(5)
    frm.TextBox6.Value = numerosElegidos(6)

End Sub

Public Sub SugerirNumeroBalotaGanador(frm As UserForm)
    Dim numeroAleatorio As Integer
    numeroAleatorio = Int((16 - 1 + 1) * Rnd + 1) ' Genera un número entre 1 y 16
    frm.TextBox7.Value = numeroAleatorio

End Sub

Public Sub GuardarDatos(frm As UserForm)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Apuestas")
    
    ' Encontrar la siguiente fila vacía en la columna B
    Dim siguienteFila As Long
    siguienteFila = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1 ' End(xlUp) Encuentra fila no vacia
    
    ' Guardar el contador en B de la fila actual
    ws.Cells(siguienteFila, "B").Value = siguienteFila - 4 ' B5 es la fila 5, así que B6 es 6 y así sucesivamente

    ' Guardar los valores de los ComboBox en la misma fila
    ws.Cells(siguienteFila, "C").Value = frm.ComboBox_Numero1.Value
    ws.Cells(siguienteFila, "D").Value = frm.ComboBox_Numero2.Value
    ws.Cells(siguienteFila, "E").Value = frm.ComboBox_Numero3.Value
    ws.Cells(siguienteFila, "F").Value = frm.ComboBox_Numero4.Value
    ws.Cells(siguienteFila, "G").Value = frm.ComboBox_Numero5.Value
    ws.Cells(siguienteFila, "H").Value = frm.ComboBox_Numero6.Value
    ws.Cells(siguienteFila, "I").Value = frm.ComboBox_ElijaBalota.Value
End Sub

Public Sub ValidarYGuardarDatos(frm As UserForm)
    ' Llama a la subrutina que verifica números únicos
    If VerificarNumerosUnicos(frm) Then
        ' Si todos los números son únicos, procede a guardar
        GuardarDatos frm
        MsgBox "¡Los números son validos y se han guardado correctamente!", vbInformation
    Else
        ' Si hay números repetidos, muestra un mensaje y no guarda
        MsgBox "Hay que volver a jugar, cambia los numeros.", vbExclamation
    End If
End Sub

' Verificar si los números son únicos
Public Function VerificarNumerosUnicos(frm As UserForm) As Boolean
    Dim numeros(1 To 6) As Variant
    Dim i As Integer, j As Integer
    
    ' Almacena los valores de los ComboBox en un array
    numeros(1) = frm.ComboBox_Numero1.Value
    numeros(2) = frm.ComboBox_Numero2.Value
    numeros(3) = frm.ComboBox_Numero3.Value
    numeros(4) = frm.ComboBox_Numero4.Value
    numeros(5) = frm.ComboBox_Numero5.Value
    numeros(6) = frm.ComboBox_Numero6.Value

    ' Verifica que no haya números repetidos en el array
    For i = 1 To 6
        For j = i + 1 To 6
            If numeros(i) = numeros(j) Then
                VerificarNumerosUnicos = False ' Hay números repetidos
                Exit Function
            End If
        Next j
    Next i

    VerificarNumerosUnicos = True ' Todos los números son únicos
End Function

Public Sub BuscarGanadores(frm As UserForm)
  
   Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Apuestas")

    ' Verificar si todos los TextBox tienen valores
    If frm.TextBox1.Value = "" Or frm.TextBox2.Value = "" Or frm.TextBox3.Value = "" Or _
       frm.TextBox4.Value = "" Or frm.TextBox5.Value = "" Or frm.TextBox6.Value = "" Or _
       frm.TextBox7.Value = "" Then
        MsgBox "No hay datos ganadores, por favor obtenga los números ganadores", vbExclamation
        Exit Sub
    End If

    ' Asignar los números ganadores de los TextBox a variables
    Dim numerosGanadores(1 To 6) As Integer
    Dim balotaGanadora As Integer
    numerosGanadores(1) = CInt(frm.TextBox1.Value)
    numerosGanadores(2) = CInt(frm.TextBox2.Value)
    numerosGanadores(3) = CInt(frm.TextBox3.Value)
    numerosGanadores(4) = CInt(frm.TextBox4.Value)
    numerosGanadores(5) = CInt(frm.TextBox5.Value)
    numerosGanadores(6) = CInt(frm.TextBox6.Value)
    balotaGanadora = CInt(frm.TextBox7.Value)

    ' Contar el total de apuestas
    Dim totalApuestas As Long
    totalApuestas = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row - 4
    MsgBox "El número total de apuestas de hoy es de: " & totalApuestas, vbInformation

    ' Variables para resultados
    Dim fila As Long, aciertos As Integer, balotaCoincide As String
    Dim resultado As String
    resultado = " ID_Apuesta -        Fila   -     NumeroAciertos -    AcertoBalota" & vbCrLf

    ' Contador de ganadores
    Dim contadorGanadores As Integer
    contadorGanadores = 0

    ' Recorrer las apuestas para verificar condiciones
    For fila = 5 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        aciertos = 0
        balotaCoincide = "NO"

        ' Contar aciertos
        Dim i As Integer
        For i = 1 To 6
            If ws.Cells(fila, "C").Value = numerosGanadores(i) Or _
               ws.Cells(fila, "D").Value = numerosGanadores(i) Or _
               ws.Cells(fila, "E").Value = numerosGanadores(i) Or _
               ws.Cells(fila, "F").Value = numerosGanadores(i) Or _
               ws.Cells(fila, "G").Value = numerosGanadores(i) Or _
               ws.Cells(fila, "H").Value = numerosGanadores(i) Then
                aciertos = aciertos + 1
            End If
        Next i

        ' Verificar si la balota es correcta
        If ws.Cells(fila, "I").Value = balotaGanadora Then
            balotaCoincide = "SI"
        End If

        ' Solo agregar al resultado si hay 5 o más aciertos
        If aciertos >= 5 Then
            contadorGanadores = contadorGanadores + 1
            resultado = resultado & vbTab & ws.Cells(fila, "B").Value & "|" & vbTab & "|" & fila & vbTab & "|" & _
                        aciertos & vbTab & vbTab & "|" & balotaCoincide & vbCrLf
        End If
    Next fila

    ' Mostrar resultados o mensaje si no hubo ganadores
    If contadorGanadores > 0 Then
        MsgBox resultado, vbInformation
    Else
        MsgBox "No hubo ganadores correspondientes", vbExclamation
    End If
  
End Sub

Public Sub CerrarYGuardarExcel()
    ' Guardar y cerrar el libro
    ThisWorkbook.Save
    Application.Quit

End Sub

Public Sub AbrirFormulario()
    ' Abre el formulario llamado "NombreDelFormulario"
    JuegoBalotto.Show vbModeless
    
End Sub
