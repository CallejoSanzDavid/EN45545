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

-------------------------------------------------- BBDD F&S - VERSIÓN 2.1 -------------------------------------------------

- Cambios mínimos: Comentarios.

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

------------------------------------------------ BBDD Pedidos - VERSIÓN 3.2 ----------------------------------------------

- Cambio del nombre de los módulo para mejor identificación.

- PorArchivar_Archivados: Se busca la línea del Part Number archivado y marca la línea como OK en "EN CURSO".

--------------------------------------------------- BBDD F&S - VERSIÓN 3 --------------------------------------------------

- Eliminación del mensaje al inicio al haber recibido observaciones de los usuarios de que al abrir el archivo han actualizado la Info sin querer 
  y han tenido que esperar a que terminara el proceso para consultar información.
  
- Cambiadas de posición las columnas “Comments / Remarks” (CB) y “Manufacturer Declaration Date” (CA) para mantener el formato con la plantilla original.
  Esto es necesario para la nueva tool de corrección de FCILs.

- CheckStatus: Modificación masiva de código.
	- Correcciones en el código de “COMPROBAR CADUCIDAD”. Error de rellenado del estado, no tenía en cuenta la fecha de las declaraciones de conformidad en el estado global.
	- Función Contar_Elem_DB() eliminada por ser innecesaria.
	- Función Contar_Elem() eliminada por ser innecesaria.
	- Comparador_Fechas() para optimizar el código.
	- Añadida función StatusGlobal() para optimizar el código.
	- Optimización del código para reducir tiempos de procesado.
	- Corrección de error en caso de que se active la macro con una celda pulsada fuera de la tabla.

- EmailGen: Eliminada función ContarElem() por ser innecesaria.

-------------------------------------------------- BBDD F&S - VERSIÓN 3.1 -------------------------------------------------
Cambios en CheckStatus.

- CheckStatus: Reestructuración de código.

- Nuevo módulo: ClearFormat_Test: Limpia los rangos de celdas para facilitar el testeo del código

-------------------------------------------------- BBDD F&S - VERSIÓN 3.2 -------------------------------------------------
Cambios en EmailGen.

- EmailGen:
	- Unificadas las funciones Alarmas() y AlarmasX() ya que tenían una estructura muy parecida y es innecesario tener dos funciones que hacen lo mismo.
	- Líneas de código reorganizadas para evitar repeticiones innecesarias.
	- Optimización del código para mayor velocidad de computación.

-------------------------------------------------- BBDD F&S - VERSIÓN 3.3 -------------------------------------------------
Cambios en CheckStatus.

- Corrección de error de funcionamiento en la función que busca la información de contacto.

--------------------------------------------------- BBDD F&S - VERSIÓN 4 --------------------------------------------------
Cambio de la filosofía de programación. Funciones cortas y eficientes.
Todos los módulos renombrados para facilitar su identificación.

-EmailGen:
	- Cambio de los nombre de los módulos para ser más descriptivos.
	- Estandarización del código mediante la línea: SheetName = ActiveSheet.Name.
	- Optimizado código de función Alarmas mediante Select Case.
	- Cambiados los rangos de busqueda para los identificadores de filas y columnas. Estandarización y optimización.
	- Los correos solo añaden información de los los materiales cuyos certificados están expirados o a punto de expirar.
	- Public funtion de filtros.
	- Cambio de modo de programación: Funciones cortas y eficientes.

- Filters: Nuevo módulo estándar para aplicar filtros.

- GlobalEntities: Inicialización de entidades globales.

--------------------------------------------------- BBDD F&S - VERSIÓN 5 --------------------------------------------------
Adaptación del código InfoProveedores_SAP de “BOM Check REACH format Data Base V1.1.xlsm”.

- Añadida nueva hoja "Ranking Status" para simplificar el código de colores.

- Variables en “GlobalEntities” organizadas dentro de las funciones en donde se inicializan.
- En “GlobalEntities” solo se declaran las variables que se inicializan en varios módulos.

- Renombrado de funciones a Inglés con el formato “Function_Name”.
- Traducidos los comentarios en Inglés.

- CheckStatus:
	- Nueva función: NoDate(). Función que se invoca cuando no hay fecha para un ensayo. <----------- Eliminada
	- Nueva función: CountersCheck(). Checkea y actualiza el estado de los contadores de los bucles.
	- Cambio de Nombres de variables para hacer el código más comprensible. Por ejemplo: G por GlobalStatusj.
	- Eliminada cadena status(6,1) por ser innecesaria. Por lo que podemos prescindir del valor k.
	- Check_Contacts reestructurado.
	- Eliminadas variables innecesarias.
	
- LocatePositions:
	- Error corregido.
	












