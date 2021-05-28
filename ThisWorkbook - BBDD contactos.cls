VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    a = MsgBox("Esta base de datos está vinculada a la base de datos FCIL. Para evitar problemas en la comunicación entre los libros se recomienda seguir las siguientes pautas.", vbExclamation)
    MsgBox ("1.  NO puede haber más de 1 correo en una celda. Esto daría un error a la hora de mandar el correo." + vbCrLf + "2.  Para añadir un nuevo contacto:" + vbCrLf + "   a.  Si la empresa existe: Insertar una fila debajo del nombre de la empresa y completar el nombre de la empresa y el correo electrónico. El resto de datos no afectan a la macro, por lo que son opcionales." + vbCrLf + "   b.  Si la empresa no existe: Insertar una fila DENTRO DE LA TABLA. De esta forma se podrán utilizar los filtros correctamente." + vbCrLf + "3.  Para quitar un contacto seleccionar la fila del contacto y eliminar." + vbCrLf + "4.  Bajo ningún concepto combinar celdas. Esto impediría filtrar los datos de la tabla y dificultaría su manipulación." + vbCrLf + vbCrLf + "IMPORTANTE: Si se modifica la tabla pulsar el botón ACTUALIZAR FORMATO.")

End Sub
