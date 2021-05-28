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

----------------------------------------------- BBDD Pedidos - VERSIÓN 1 ----------------------------------------------

- 
