# Macros utiles de acuerdo a la ocasion

## Generar hojas de un listado

- Deebe tener la pestaña de programado activada
- esta macro funciona desde vba
- el listad de nombres debe iniciar en la celda A1

`
    Sub Generar_Hojas()
    'SE DEBE TENER EN LA COLUMNA A DESDE LA CELDA A1 EL NOMBRE DE LAS HOJAS A CREAR
    Dim c1 As Integer
    Dim d As Integer
    c1 = 1
    Sheets("Hoja1").Activate 'Si cambia el nombre de la hoja donde están los nombres, colocarla aquí.
    ActiveSheet.Range("A1").Activate
    Do While Not IsEmpty(ActiveCell)
    c1 = c1 + 1
    ActiveCell.Offset(1, 0).Activate
    Loop
    For i = 1 To c1 - 1 Step 1
        Sheets("Hoja1").Select
        Sheets.Add After:=ActiveSheet
        Sheets("Hoja1").Select
        Sheets("Hoja1").Activate
        Sheets(2).Name = Cells(i, 1)
    Next i
    d = c1 - 1
    MsgBox Prompt:="USTED HA CREADO " & d & " HOJAS", Title:="¡FELICITACIONES!"
    End Sub
`
