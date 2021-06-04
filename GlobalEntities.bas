Attribute VB_Name = "GlobalEntities"
Option Explicit
    
    Public CPmaili As Integer
    'Localizador de materiales - ComplexPartNumber
    Public nombi As Integer
    '-----------------LocatePositions-----------------
    'Valores - Email_Gen
    Public nombre As String
    Public material As String
    Public Valj As Integer
    Public nproducto As String
    Public manufacturer As String
    'Variables Split - Email_Gen
    Public auxname() As String
    Public Auxsplit As String
    'Variables Split - SpanishModule
    Public statusES() As String
    'Variables Split - Alarmas
    Public auxday() As String
    Public daynum As Integer
    
    'Información del correo
    Public OutApp As Object
    Public OutMail As Object
    Public EncabezadoEN As String
    Public EncabezadoES As String
    Public InfoENRW As String
    Public AuxENRW As String
    Public InfoESRW As String
    Public AuxESRW As String
    Public InfoEN As String
    Public InfoES As String
    Public FinalInfoEN As String
    Public FinalInfoES As String
    Public DespedidaEN As String
    Public DespedidaES As String
    Public Firma As String
    Public Separacion As String
    Public Destinatario As String
    
    'Valores devueltos por funciones
    Public validacion As Integer
    Public NoContact As Integer
    Public statusmin As Integer
    
    'Contadores
    Public ncorreos As Integer
    Public nsincontacto As Integer
    Public nexport As Integer
    
    'Marcadores o flags
    Public lasterror As Integer
    Public stat As Integer
    
    'Select Case Option
    Public Case_Option As String
    
    'Valores auxiliares.
    Public expstatus As String
    Public auxstatus As Integer
    
    
'-----------------SAP_InfoProveedores-----------------
    'BaseProveedores
    'Contador de bucle
    Public m As Integer
    
    Public mailj As Integer
    Public c As Range
    Public a As Integer
    'Localizador de filas - Contacto de proveedores
    Public InicioConti As Integer
    Public ContarDBi As Integer
    'Localizadores de columnas y su valor - Contacto de proveedores
    Public supplierj As Integer
    Public supplier As String
    Public merakj As Integer
    Public tlfnoj As Integer
    Public countryj As Integer
    Public languagej As Integer
    
    'LocateSupplier
    Public linea As Integer
    Public InfoUpdated As Integer
    
    'ME2M_SAP_SUPPLIER_CONTACT
    'Nombre de hojas
    Public DBContactSheetName As String
    
    'Marcadores o flags
    Public saveflag As Integer
    
    Public mark As Integer
    Public meraki As Integer
    Public EndDBi As Integer
    Public contacto As String
    Public language As String
    Public obj As Object
       
'-----------------Variables comunes-----------------
       
    'Contador de bucle
    'Public i As Integer 'Delete?
    
    
    
    





















