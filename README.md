# EN45545
 
 Control de versiones para los módulos de automatización para la Base de Datos de Fuegos y Humos.
 
-------------------------------------------------- BBDD F&S - VERSIÓN 1 ---------------------------------------------------
 
 - CheckStatus: Módulo para checkear el estado de los Test Reports individuales de cada ensayo y globales de cada Part Number.
 
 - EmailGen: Generación de emails automáticamente para las lineas en las que algún certificado haya expirado o expire en los próximos 6 meses. 
   Los correos se generan con un lista con información todos los Part Numbers que necesitan nuevos certificados para cada proveedor.
 
-------------------------------------------------- BBDD F&S - VERSIÓN 2.1 -------------------------------------------------

- Vinculación con la BBDD de Contactos.

- Añadido código en "ThisWorkbook" para generar una MsgBox al iniciar la BBDD. Corre el módulo CheckStatus.
 
- EmailGen: Función EXPORT_DATA en módulo EmailGen registra información en BBDD de Pedidos.
 
- PartNameCheck: Código creado a demanda para comprobar inconsistencias entre los Part Names con el mismo Part Number.

----------------------------------------------- BBDD Contactos - VERSIÓN 1 -----------------------------------------------

- Hoja1: Código para organizar automáticamente la información de contacto de la BBDD de contactos.

- Añadido código en "ThisWorkbook": Se generan dos mensajes al iniciar la BBDD de contactos con instrucciones de como añadir información en la BBDD.

------------------------------------------------- BBDD Pedidos - VERSIÓN 1 -----------------------------------------------

- EnCurso_OK_PorArchivar: Mueve las líneas resueltas desde la hoja "EN CURSO" a "OK".

------------------------------------------------- BBDD Pedidos - VERSIÓN 2 -----------------------------------------------
Se adapta al nuevo formato prescindiendo de los campos innecesarios. Nuevos módulos añadidos:

- EnCurso_OK_PorArchivar (Actualizado): Añadido condicional para archivar las líneas en "POR ARCHIVAR". Errores corregidos.

- PorArchivar_Archivados: Nuevo código para mover las líneas resueltas de "POR ARCHIVAR" a "ARCHIVADOS".

- Temp_EnCurso: Nuevo código para mover las líneas resueltas de "Temp" a "EN CURSO". 
  En "Temp" se registran las líneas con la información necesaria para poder hacer el seguimiento de los pedidos de certificados. 
  Estas líneas se registran mediante el código del módulo EmailGen.

------------------------------------------------- BBDD Pedidos - VERSIÓN 3 -----------------------------------------------
------------------------------------------------ BBDD Pedidos - VERSIÓN 3.1 ----------------------------------------------
- Estandarización del código:
	
	- Application.ScreenUpdating = False/True: Desactivación de la actualización de pantalla para evitar parpadeos.
	
	- SheetName = ActiveSheet.Name: Se registra el nombre de la hoja en la que se ejecuta el código en una variable.

- Reorganización del código para evitar errores.










