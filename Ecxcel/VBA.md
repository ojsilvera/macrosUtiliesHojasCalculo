# VBA Macro to Create Sheets

Scrip de vba para excel el cual crea hojas en un libro de acuerdo a un rango de celdas determinadas, las cuales poseen
los nombres de las hojas.

``
    Sub crear_hojas()

        Dim lista As Range
        Dim ix As Long

        Set lista = Application.InputBox(prompt:="señalar rango de la lista", Title:="lista de nombres", Type:=8)

        Application.ScreenUpdating = False

        For ix = lista.Count To 1 Step -1
            Sheets.Add.Name = lista(ix)
        Next ix
            Sheets(1).Select

        Application.ScreenUpdating = True

    End Sub
``
