# Macros utiles de acuerdo a la ocasion

## Generar hojas de un listado

- Debe tener la pestaña de programado activada
- esta macro funciona desde vba
- el listad de nombres debe iniciar en la celda A1

    ```visual basic
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
    ```

[Fuente]([https://](https://youtu.be/kuR44uyi_V8?si=bf9VFir7GgZWrMCp))

## Generar hojas de un listado con un solo click

- Debe tener la pestaña de programado activada.
- esta macro funciona desde vba.
- posee una ventana para seleccionar los nombres de las hojas, el rango puede estar ubicacdo en cualquier lugar de la
  hoja activa.

    ```visual basic
        Sub crear_varias_hojas()
        Dim lista As Range
        Dim ix As Long
        Set lista = Application.InputBox(prompt:="Seleccione el nombre para las hojas", Title:="nombres para cada hoja", Type:=8)
        Application.ScreenUpdating = False
        For ix = lista.Count To 1 Step -1
        Sheets.Add.Name = lista(ix)
        Next ix
        Sheets(1).Select
        Application.ScreenUpdating = True
        End Sub
    ```

[Fuente]([https://](https://www.youtube.com/watch?v=3soatT0SGhI&ab_channel=MiltonJMorales))

## Listas desplegables auto-Actualizables

- Origen de datos en la misma hoja

  - El rango que contiene los elemento de la lista desplegable, lo convertimos en una tabla.
  - Con la tabla anteriormente creada vamos a validacion de datos y seleccionamos la lista de elementos de los que se
      conformara la lista
  - Si agregamos un nuevo elemento a la tabla, la lista desplegable se actualizara automaticamente

- Origen de datos en una hoja diferente

  - Seleccionamos los elementos que van a componer la lista y le asignamos un nombre de rango
  - Nos ubicamos en la celda que alojara nuestra lista desplegable
  - Vamos a validadicon de datos en la pestaña Datos
  - Escogemos lista y acto seguido presionamos F3, esto nos mostrara la lista de rangos disponibles
  - Seleccionamos el rango que hemos creado con los elementos de la lista
  - Al agrgar un nuevo elemento la lista se actualizara automaticamente
