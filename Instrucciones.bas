Attribute VB_Name = "Instrucciones"
Sub Instrucciones()

    a = MsgBox("Esta base de datos est� vinculada a la base de datos FCIL. Para evitar problemas en la comunicaci�n entre los libros se recomienda seguir las siguientes pautas.", vbExclamation)
    MsgBox ("1.  NO puede haber m�s de 1 correo en una celda. Esto dar�a un error a la hora de mandar el correo." + vbCrLf + "2.  Para a�adir un nuevo contacto:" + vbCrLf + "   a.  Si la empresa existe: Insertar una fila debajo del nombre de la empresa y completar el nombre de la empresa y el correo electr�nico. El resto de datos no afectan a la macro, por lo que son opcionales." + vbCrLf + "   b.  Si la empresa no existe: Insertar una fila DENTRO DE LA TABLA. De esta forma se podr�n utilizar los filtros correctamente." + vbCrLf + "3.  Para quitar un contacto seleccionar la fila del contacto y eliminar." + vbCrLf + "4.  Bajo ning�n concepto combinar celdas. Esto impedir�a filtrar los datos de la tabla y dificultar�a su manipulaci�n." + vbCrLf + vbCrLf + "IMPORTANTE: Si se modifica la tabla pulsar el bot�n ACTUALIZAR FORMATO.")

End Sub


                                                                                                                                                                                                                                                                                                                                                                                                    v=� �	��!  �R��                  