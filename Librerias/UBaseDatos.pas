//-----------------------------------------------------------------------------
//                           Delphi Runtime Library
//                 Copyright(c) 2021 Bismarck Sierra Ibarra
//                       bismarcksierra@gmail.com
//-----------------------------------------------------------------------------
unit UBaseDatos;

interface

uses
  Classes, FireDAC.Comp.Client, FireDAC.stan.def, FireDAC.Phys.MSSQL, SysUtils, FireDAC.stan.Param, FireDAC.Dapt,
  FireDAC.VCLUI.Wait, FireDAC.stan.Async, FireDAC.Phys.FB, Data.DB, FireDAC.Phys.MySQL, System.Diagnostics,
  System.Variants, UBitacora, FireDAC.Comp.ScriptCommands, FireDAC.Comp.Script, FireDAC.UI.Intf, FireDAC.FMXUI.Wait;

const
  FORMATO_FECHA_FIREBIRD   = 'mm/dd/yyyy';
  FORMATO_FECHA_SQL_SERVER = 'yyyy/mm/dd';
  FORMATO_FECHA_MY_SQL     = 'dd/mm/yyyy';

  FECHA_HORA_FIREBIRD = 'CURRENT_TIMESTAMP';

  VALOR_SI = 'S';
  VALOR_NO = 'N';

type
  TTipoConexion = (tcFirebird, tcSqlServer, tcMySql);

  TCampoValor = class
  private
    FValor : Variant;
  public
    constructor Create(const AValor : Variant);
    property Valor : Variant read FValor;
  end;

  TProcSalidaScript = procedure(ATextoSalida : string) of object;

  /// <summary> Permite realizar consultas e incluye un Datasource para conectarse a componentes visuales </summary>
  /// <summary> Oculta el componente de consulta para que pueda ser cambiada sin problema </summary>
  TConsulta = class(TFDQuery)
  private
    FDataSource : TDataSource;
    function GetFuenteDatos : TDataSource;
    function GetFetchRecordCount : integer;
  public
    destructor Destroy; override;

    /// <summary> Modifica o agrega si no existe un parámetro de la consulta</summary>
    procedure AgregarParametro(ANombreParametro : string; AValor : Variant); overload;

    /// <summary> Obtiene el valor de un campo de tipo cadena por su índice </summary>
    function ObtenerCampoCadena(AIndice : integer) : string;

    /// <summary> Obtiene el valor de un campo de tipo entero por su índice </summary>
    function ObtenerCampoEntero(AIndice : integer) : integer;

    /// <summary> Obtiene el valor de un campo de tipo flotante por su índice </summary>
    function ObtenerCampoFlotante(AIndice : integer) : double;

    /// <summary> Obtiene el valor de un campo de tipo Variant por su índice </summary>
    function ObtenerCampoVariante(AIndice : integer) : Variant;

    /// <summary> Obtiene el valor de un campo de tipo cadena por su nombre </summary>
    function ObtenerCampoNombreCadena(sCampo : string) : string;

    /// <summary> Obtiene el valor de un campo de tipo entero por su nombre </summary>
    function ObtenerCampoNombreEntero(sCampo : string) : integer;

    /// <summary> Obtiene el valor de un campo de tipo flotante por su nombre </summary>
    function ObtenerCampoNombreFlotante(sCampo : string) : real;

    /// <summary> Libera la fuente de datos ligada a la consulta </summary>
    procedure LiberarFuenteDatos;

    /// <summary> Limpia todos los parámetros de la consulta </summary>
    procedure LimpiarParametros;

    /// <summary> Avanza al siguiente registro del conjunto de datos </summary>
    procedure Siguiente;

    // <summary> Indica si se llegó al fin del conjunto de datos </summary>
    function EsFinDatos : boolean;

    /// <summary> Regresa el conteo REAL de registros, activando el FetchAll </summary>
    property NumeroRegistros : integer read GetFetchRecordCount;

    /// <summary> Componente para conectar con componentes visuales </summary>
    property FuenteDatos : TDataSource read GetFuenteDatos;
  end;

  /// <summary> Clase que encapsula las operaciones de la base de datos para facilitar su uso</summary>
  /// <remarks> Bismarck Sierra Ibarra - 2017-04-01 </remarks>
  TBaseDatos = class
  private
    /// <summary> FLog </summary>
    FBitacora : TBitacora;
    /// <summary> Tipo de conexión </summary>
    FTipoConexion : TTipoConexion;
    /// <summary> Nombre de la base de datos </summary>
    FBaseDatos : string;
    /// <summary> Componente de conexión a la base de datos </summary>
    FConexion : TFDConnection;
    /// <summary> Indica si se encuentra conectado actualmente a la base de datos </summary>
    FConectado : boolean;
    /// <summary> Contraseña para acceder a la base de datos </summary>
    FContrasenia : string;
    /// <summary> Listado que almacena las consultas activas que no se han liberado </summary>
    FListaConsultas : TStringList;
    /// <summary> Puerto para conectarse a la base de datos </summary>
    FPuerto : string;
    /// <summary> Procedimiento para reportar la salida de ejecución de un script SQL </summary>
    FScriptProcedimiento : TProcSalidaScript;
    /// <summary> Si está encendida solo reportara errores al ejecutar un script, de foram contraria reporta todo </summary>
    FScriptSoloErrores : Boolean;
    /// <summary> Nombre del servidor </summary>
    FServidor : string;
    /// <summary> Tiempo de ejecución de la última sentencia SQL </summary>
    FTiempoEjecucion : int64;
    /// <summary> Usuario para acceder a la base de datos </summary>
    FUsuario : string;
    /// <summary> Último error registrado </summary>
    FUltimoError : string;

    /// <summary> Escribe un mensaje en el Log </summary>
    procedure EscribirBitacora(AModulo : string; AMensaje : string;
      ATipoRegistro : TTipoRegistroBitacora = tbExcepcion);

    /// <summary> Establece la conexión con la base de datos o se desconecta </summary>
    procedure setConectado(const Value : boolean);

    procedure FDScriptSpoolPut(AEngine : TFDScript; const AMessage : string; AKind : TFDScriptOutputKind);
  public
    /// <summary> Crea los objetos privados necesarios </summary>
    constructor Create;

    /// <summary> Destruye los objetos privados </summary>
    destructor Destroy; override;

    /// <summary> Inicia una nueva transacción </summary>
    /// <returns> Verdadero si se tuvo éxito, falso de forma contraria </returns>
    function ComenzarTransacion : boolean;

    /// <summary> Aplica la transacción en curso</summary>
    /// <returns> Verdadero si se tuvo éxito, falso de forma contraria </returns>
    function AplicarTransaccion : boolean;

    /// <summary> Conecta con una base de datos </summary>
    /// <param name="pConnectionString"> Cadena de conexión a la base de datos </param>
    /// <returns> Verdadero si se conectó, falso si falló </returns>
    function Conectar(ACadenaConexion : string) : boolean; overload;

    /// <summary> Conecta con una base de datos</summary>
    /// <param name="AServidor"> Nombre del servidor que aloja la base de datos </param>
    /// <param name="ABaseDatos"> Nombre de la base de datos </param>
    /// <param name="AUsuario"> Nombre de usuario para conectar a la base de datos </param>
    /// <param name="AContrasenia"> Contraseña sin encriptar para cocnectarse a la base de datos </param>
    /// <param name="APuerto"> Contraseña sin encriptar para cocnectarse a la base de datos </param>
    /// <returns> Verdadero si se conectó, falso si falló </returns>
    function Conectar(AServidor, ABaseDatos, AUsuario, AContrasenia, APuerto : string) : boolean; overload;

    /// <summary> Conecta a la base de datos con los valores asginados previamente de las campos privados </summary>
    /// <returns> Verdadero si se conectó, falso si falló </returns>
    function Conectar : boolean; overload;

    /// <summary> Genera una consulta a la base de datos (solamente sentecias Select)</summary>
    /// <param name="Asentencia"> Sentencia con la consulta SQL a ejecutar </param>
    /// <param name="ANombre"> Nombre a asignar a la consulta </param>
    /// <param name="AReintentos"> Indica si se reintetará conectar a la base de datos en caso de falla </param>
    /// <returns> Instancia tipo TXDBQuery con el dataset resultado de la consulta</returns>
    function CrearConsulta(ASentencia : string; ANombre : string = ''; AReintentos : boolean = true) : TConsulta;

    /// <summary> Crea una consulta sin abrirla, útil para crear consulta con parámetros</ summary>
    /// <summary> se abre hasta que se asignan los valores  </summary>
    /// <param name="ASentencia"> Sentencia con la consulta SQL a ejecutar </param>
    /// <param name="ANombre"> Nombre a asignar a la consulta </param>
    /// <returns> Instancia tipo TXDBQuery sin abrir </returns>
    function CrearConsultaCerrada(ASentencia : string; ANombre : string = '') : TConsulta;

    /// <summary> Llena un listado con el resultado de una sentencia SQL</summary>
    /// <param name="Asentencia"> Sentencia con la consulta SQL a ejecutar </param>
    /// <param name="ALista"> Objeto instanciado </param>
    /// <returns> verdadero si no hubo error, falso de forma contraria</returns>
    function GenerarListaSQL(ASentencia : string; ALista : TStringList) : boolean;

    /// <summary> Libera un listado llenado con GenerarListaSQL</summary>
    /// <param name="ALista"> Objeto instanciado </param>
    procedure LiberarListaSQL(var ALista : TStringList);

    /// <summary> Llena un ComboBox con el contenido de una consulta</summary>
    /// <param name="ACombo"> ComboBox a llenar </param>
    /// <param name="ASentencia"> Sentencia SQL </param>
    /// <param name="ASeleccionar"> Indica si se debe seleccionar el primer elemento del combo </param>
    /// <param name="APrimeraLinea"> Inserta un texto en el primer elemento del combo </param>
    // function LlenarCombo(var ACombo : TComboBox; ASentencia : string; ASeleccionar : boolean = true;
    // APrimeraLinea : string = '') : boolean;

    /// <summary> Libera los objetos que se utilizaron al llenar un combo con la función LlenarCombo </summary>
    // procedure LiberarCombo(var ACombo : TComboBox);

    /// <summary> Libera un objeto tipo consulta y lo elimina de la lista de consultas activas</summary>
    /// <param name="AConsulta"> Consulta que se desea liberar </param>
    procedure LiberarConsulta(var AConsulta : TConsulta); overload;

    /// <summary> Libera una consulta por su nombre y lo elimina de la lista de consultas activas</summary>
    /// <param name="ANombre"> Nombre de la consulta que se desea liberar </param>
    procedure LiberarConsulta(ANombre : string); overload;

    /// <summary> Desconecta la base de datos </summary>
    /// <returns> Verdadero si se conectó, falso si falló </returns>
    procedure Desconectar;

    /// <summary> Deshace la transacción en curso</summary>
    /// <returns> Verdadero si se tuvo éxito, falso de forma contraria </returns>
    function DeshacerTransaccion : boolean;

    /// <summary> Ejecuta la sentencia SQL (Insert, Delete, Update) </summary>
    /// <param name="ASentencia"> Sentencia SQL a ejecutar </param>
    /// <param name="AReintentos"> Indica si debe hacer un reintento en caso de fallo</param>
    /// <returns> -1 si ocurrió una falla, en caso contrario regresa el número de registros afectados </returns>
    function EjecutarSentencia(ASentencia : string; AReintentos : boolean = true) : integer;

    /// <summary> Ejecuta un conjunto de sentencias SQL </summary>
    /// <param name="AScript"> Sentencias SQL a ejecutar </param>
    /// <param name="AReintentos"> Indica si debe hacer un reintento en caso de fallo</param>
    /// <returns> -1 si ocurrió una falla, en caso contrario regresa el número de registros afectados </returns>
    function EjecutarScript(AScript : TStringList; AProcedimientoSalida : TProcSalidaScript;
      ASoloErrores : boolean = false; AReintentos : boolean = true) : boolean;

    /// <summary> Obtiene un consecutivo para una tabla determinada </summary>
    function ObtenerNuevaClave(ATabla : string) : integer;

    /// <summary> Recupera un query por su nombre </summary>
    /// <param name="ANombre"> Nombre del query que se quiere recuperar  </param>
    /// <returns> Una instancia del query que se encontró o nulo si no se encontró</returns>
    function RecuperarConsultaNombre(ANombre : string) : TConsulta;

    /// <summary> Verifica si existe una tabla </summary>
    /// <param name="ANombre"> Nombre de la tabla que se requiere verificar </param>
    /// <returns> Verdadero si existe, falso de forma contraria </returns>
    function ExisteTabla(ANombre : string) : boolean;

    /// <summary> Base de datos a la que se conectará </summary>
    property BaseDatos : string read FBaseDatos write FBaseDatos;
    /// <summary> Indica el estatus de la conexión </summary>
    property Conectado : boolean read FConectado write setConectado;
    /// <summary> Contraseña para conectar a la base de datos </summary>
    property Contrasenia : string read FContrasenia write FContrasenia;
    /// <summary> Objeto Log para reportar errores </summary>
    property Log : TBitacora read FBitacora write FBitacora;
    /// <summary> Puerto de la base de datos a la que se conectará </summary>
    property Puerto : string read FPuerto write FPuerto;
    /// <summary> Nombre de del servidor al que se conectará </summary>
    property Servidor : string read FServidor write FServidor;
    /// <summary> Tiempo de ejecución de la última sentencia SQL </summary>
    property TiempoEjecucion : int64 read FTiempoEjecucion;
    /// <summary> Tipo de conexión: SQL Server (tcnSqlServer) o MySQL (tcnMySQL) </summary>
    property TipoConexion : TTipoConexion read FTipoConexion write FTipoConexion;
    /// <summary> Último error registrado <7summary>
    property UltimoError : string read FUltimoError;
    /// <summary> Nombre de usuario para conectarse a la base de datos </summary>
    property Usuario : string read FUsuario write FUsuario;
  end;

var
  GBaseDatos : TBaseDatos;

implementation

const
  PUERTO_DEFAULT = '3050';

  { TXposDataBase }

function TBaseDatos.Conectar : boolean;
begin
  result := false;
  FConexion.LoginPrompt := false;

  try
    with FConexion do
    begin
      Params.Clear;
      Connected := false;
      case FTipoConexion of
        tcFirebird :
          begin
            DriverName := 'FB';
            if (Trim(FPuerto) = EmptyStr) then
              FPuerto := PUERTO_DEFAULT;
            Params.DriverID := 'FB';
            Params.Database := FServidor + '/' + FPuerto + ':' + FBaseDatos;
            Params.UserName := FUsuario;
            Params.Password := FContrasenia;
            Params.Values['SQLDialect'] := '3';
            Params.Values['Protocol'] := 'TCP/IP';
            Params.Values['BlobSize'] := '-1';
            Params.Values['CommitRetain'] := 'false';
            Params.Values['WaitOnLocks'] := 'True';
            Params.Values['IsolationLevel'] := 'ReadCommited';
            Params.Values['EnableBCD'] := 'True';
            Params.Values['CharacterSet'] := 'ISO8859_1';
            LoginPrompt := false;
          end;
        tcSqlServer :
          begin
            DriverName := 'MSSQL';
            Params.Values['Port'] := FPuerto;
            Params.Values['Database'] := FBaseDatos;
            Params.Values['Server'] := FServidor;

            if (FUsuario <> '') then
            begin
              Params.Values['OSAuthent'] := 'No';
              Params.Values['User_Name'] := FUsuario;
              Params.Values['Password'] := FContrasenia;
            end
            else
              Params.Values['OSAuthent'] := 'Yes';
          end;
        tcMySql :
          begin
            DriverName := 'MYSQL';
            Params.Values['Database'] := FBaseDatos;
            Params.Values['Server'] := FServidor;
            Params.Values['User_Name'] := FUsuario;
            Params.Values['Password'] := FContrasenia;
          end;
      end;
      Connected := true;
      FConectado := Connected;
    end;
  except
    on e : Exception do
    begin
      FUltimoError := 'Ocurrió un error al conectarse a la base de datos. ' + e.Message;
      EscribirBitacora(UnitName + '.connect', 'Error al conectar a la base de datos. ' + e.Message, tbExcepcion);
      exit;
    end;
  end;
  result := FConexion.Connected;
end;

function TBaseDatos.Conectar(ACadenaConexion : string) : boolean;
begin
  result := false;
  try
    with FConexion do
    begin
      Connected := false;
      if (FTipoConexion = tcSqlServer) then
        ConnectionString := ACadenaConexion;
      Connected := true;
      FConectado := Connected;
    end;
  except
    on e : Exception do
    begin
      FUltimoError := 'Ocurrió un error al conectarse a la base de datos. ' + e.Message;
      EscribirBitacora(UnitName + '.connect', 'Error al conectar a la base de datos. ' + e.Message, tbExcepcion);
      exit;
    end;
  end;
  result := FConexion.Connected;
end;

constructor TBaseDatos.Create;
begin
  inherited;
  FConectado := false;
  FTipoConexion := tcFirebird;
  FConexion := TFDConnection.Create(nil);
  FListaConsultas := TStringList.Create;
end;

// procedure TBaseDatos.LiberarCombo(var ACombo : TComboBox);
// begin
// // Libera el stringList asociado al combo
// TStringList(ACombo.Tag).Free;
// ACombo.Clear;
// ACombo.Tag := 0;
// end;

procedure TBaseDatos.LiberarConsulta(ANombre : string);
var
  LIndice  : integer;
  LSql     : string;
  qryQuery : TConsulta;
begin
  LIndice := FListaConsultas.IndexOf(ANombre);
  if (LIndice <> - 1) then
  begin
    qryQuery := TConsulta(FListaConsultas.Objects[LIndice]);

    LSql := qryQuery.SQL.Text;

    qryQuery.Close;
    qryQuery.Free;
    FListaConsultas.Delete(LIndice);
  end;
end;

procedure TBaseDatos.LiberarListaSQL(var ALista : TStringList);
var
  I      : integer;
  LValor : TCampoValor;
begin
  if (ALista = nil) then
    exit;

  for I := 0 to Pred(ALista.Count) do
  begin
    LValor := TCampoValor(ALista.Objects[I]);
    LValor.Free;
    ALista.Objects[I] := nil;
  end;
  FreeAndNil(ALista);
end;

// function TBaseDatos.LlenarCombo(var ACombo : TComboBox; ASentencia : string; ASeleccionar : boolean;
// APrimeraLinea : string) : boolean;
// var
// qryConsulta : TConsulta;
// LLista      : TStringList;
// LPos        : integer;
// begin
// result := true;
// ACombo.Items.BeginUpdate;
// ACombo.Clear;
// if (APrimeraLinea <> EmptyStr) then
// ACombo.Items.AddObject(APrimeraLinea, TObject(0));
//
// qryConsulta := CrearConsulta(ASentencia);
// with qryConsulta do
// begin
// if (qryConsulta = nil) then
// exit;
//
// // Si ya había un StringList asociado al combo, lo libera
// if (ACombo.Tag > 0) then
// TStringList(ACombo.Tag).Free;
//
// // if not (Fields[1].DataType in [ftInteger, ftSmallint]) then
// // begin
// LLista := TStringList.Create;
// ACombo.Tag := integer(LLista);
// // end;
//
// while (not EsFinDatos) do
// begin
// if (Fields[1].DataType in [ftInteger, ftSmallint]) then
// ACombo.AddItem(CampoCadena(0), TObject(CampoEntero(1)))
// else
// begin
// LPos := LLista.Add(CampoCadena(1));
// ACombo.Items.AddObject(CampoCadena(0), TObject(LLista.Strings[LPos]));
// end;
// Siguiente;
// end;
// ACombo.Items.EndUpdate;
// if (ASeleccionar) then
// ACombo.ItemIndex := 0
// else
// ACombo.ItemIndex := - 1;
// end;
// LiberarConsulta(qryConsulta);
// Result := true;
// end;

function TBaseDatos.AplicarTransaccion : boolean;
begin
  result := false;
  try
    if (FConexion.InTransaction) then
      FConexion.Commit
    else
    begin
      FUltimoError := 'Ocurrió un error al intentar cerrar la transacción en la base de datos';
      raise Exception.Create('No existe una transacción abierta');
    end;
  except
    on e : Exception do
    begin
      FUltimoError := 'Ocurrió un error al intentar cerrar la transacción en la base de datos. ' + e.Message;
      EscribirBitacora(UnitName + '.commitTransaction', e.Message, tbExcepcion);
      exit;
    end;
  end;
  result := true;
end;

function TBaseDatos.Conectar(AServidor, ABaseDatos, AUsuario, AContrasenia, APuerto : string) : boolean;
begin
  FServidor := AServidor;
  FBaseDatos := ABaseDatos;
  FUsuario := AUsuario;
  FContrasenia := AContrasenia;
  FPuerto := APuerto;
  result := Conectar;
end;

function TBaseDatos.CrearConsulta(ASentencia : string; ANombre : string = ''; AReintentos : boolean = true) : TConsulta;
var
  LIndice    : integer;
  stpChronom : TStopWatch;
begin
  result := nil;

  if (not FConectado) then
  begin
    FUltimoError := 'Ocurrió un error al acceder a una base de datos desconectada';
    EscribirBitacora(UnitName + '.createQuery', 'Error no se encuentra conectado a la base de datos: ' + ASentencia,
      tbExcepcion);
    exit;
  end;

  if (ANombre <> '') then
  begin
    // Busca el nombre de la consulta, si la encuentra la recupera en lugar de crearla
    LIndice := FListaConsultas.IndexOf(ANombre);
    if (LIndice >= 0) then
    begin
      result := TConsulta(FListaConsultas.Objects[LIndice]);
      result.Close;
    end;
  end;

  // Si no se encontró la consulta, la crea
  if (result = nil) then
  begin
    result := TConsulta.Create(nil);

    if (ANombre <> '') then
      result.name := ANombre
    else
      result.name := 'qry' + IntToStr(FListaConsultas.Count + 1);
    result.Connection := FConexion;
    FListaConsultas.AddObject(result.name, TObject(result));
  end;

  // Crea el cronómetro
  stpChronom := TStopWatch.StartNew;

  try
    // Inicia el conteo de tiempo
    stpChronom.Start;

    result.SQL.Clear;
    result.SQL.Add(ASentencia);
    result.Open;

    // Detiene el conteo de tiempo
    stpChronom.Stop;

  except
    on e : Exception do
    begin
      result.Free;
      result := nil;
      FUltimoError := 'Ocurrió un error al intentar consultar la base de datos. ' + e.Message;

      stpChronom.Stop;

      EscribirBitacora(UnitName + '.createQuery', 'Error al ejecutar la consulta: ' + ASentencia + '. ' + e.Message,
        tbExcepcion);

      // Si se desconectó, intenta conectarse de nuevo y volver a ejecutar la consulta
      if (AReintentos) and (not FConexion.Connected) and (Conectar) then
        result := CrearConsulta(ASentencia, ANombre, false)
      else
      begin
        // Genera un mensaje crítico indicando que la estructura es incorrecta
        if (FConexion.Connected) then
          EscribirBitacora(UnitName + '.createQuery', 'Error al ejecutar la consulta: ' + ASentencia + '. ' + e.Message,
            tbExcepcion);

        result.Free;
        FListaConsultas.Delete(FListaConsultas.Count - 1);
      end;
    end;
  end;
end;

destructor TBaseDatos.Destroy;
var
  I                     : integer;
  qryQuery              : TConsulta;
  LConsultasNoLiberadas : string;
begin
  FConexion.Connected := false;
  FConexion.Free;

  // if (DebugHook <> 0) then
  // begin
  // Verifica si hay consultas no liberadas
  for I := FListaConsultas.Count - 1 downto 0 do
  begin
    qryQuery := TConsulta(FListaConsultas.Objects[I]);
    LConsultasNoLiberadas := LConsultasNoLiberadas + IntToStr(I + 1) + ' - ' + qryQuery.SQL.Text + #10#13;
    // Libera la consulta
    LiberarConsulta(qryQuery);
    // end;
    // if (LConsultasNoLiberadas <> EmptyStr) then
    // ShowMessage('Consultas no liberadas: ' + #10#13 + LConsultasNoLiberadas);
  end;

  FListaConsultas.Free;
  inherited;
end;

procedure TBaseDatos.Desconectar;
begin
  FConexion.Connected := false;
  FConectado := false;
end;

function TBaseDatos.ComenzarTransacion : boolean;
begin
  result := false;
  try
    if (not FConexion.InTransaction) then
      FConexion.StartTransaction
    else
    begin
      FUltimoError := 'Ocurrió un error al iniciar la transacción en la base de datos';
      EscribirBitacora(UnitName + '.beginTransaction',
        'Ocurrió un error al iniciar una transacción, ya existe una abierta', tbExcepcion);
      raise Exception.Create(UnitName + ': Ya existe una transacción abierta');
    end;
  except
    on e : Exception do
    begin
      FUltimoError := 'Ocurrió un error al iniciar la transacción en la base de datos. ' + e.Message;
      EscribirBitacora(UnitName + '.beginTransaction', e.Message, tbExcepcion);
      exit;
    end;
  end;
  result := true;
end;

procedure TBaseDatos.LiberarConsulta(var AConsulta : TConsulta);
var
  LIndice : integer;
  LNombre : string;
  LSql    : string;
begin
  if (AConsulta <> nil) then
  begin

    LIndice := FListaConsultas.IndexOf(AConsulta.name);
    if (LIndice >= 0) then
    begin
      LNombre := AConsulta.name;
      LSql := AConsulta.SQL.Text;
      FListaConsultas.Delete(LIndice);
    end;

    AConsulta.Close;
    AConsulta.Free;
    AConsulta := nil;
  end;
end;

function TBaseDatos.EjecutarSentencia(ASentencia : string; AReintentos : boolean = true) : integer;
var
  LSql        : TConsulta;
  LCronometro : TStopWatch;
begin
  result := - 1;

  if (not FConectado) then
  begin
    FUltimoError := 'Ocurrió un error al acceder a una base de datos desconectada';
    EscribirBitacora(UnitName + '.executeSentence', 'Error no se encuentra conectado a la base de datos');
    exit;
  end;

  // Crea el cronómetro
  LCronometro := TStopWatch.StartNew;

  LSql := TConsulta.Create(nil);
  try
    // Inicia el conteo de tiempo
    LCronometro.Start;

    LSql.Connection := FConexion;
    LSql.SQL.Add(ASentencia);
    LSql.ExecSQL;
    result := LSql.RowsAffected;

    // Detiene el conteo de tiempo
    LCronometro.Stop;
    FTiempoEjecucion := LCronometro.ElapsedMilliseconds;
  except
    on e : Exception do
    begin
      FUltimoError := 'Ocurrió un error al intentar afectar a la base de datos. ' + e.Message;

      // Detiene el conteo de tiempo
      LCronometro.Stop;
      FTiempoEjecucion := LCronometro.ElapsedMilliseconds;

      EscribirBitacora(UnitName + '.executeSentence', 'Error al ejecutar la sentencia. ' + e.Message, tbExcepcion);

      // Intenta conectarse de nuevo a la base de datos en caso de desconexión
      if (AReintentos) and (not FConexion.Connected) and (Conectar) then
        result := EjecutarSentencia(ASentencia, false);
    end;
  end;
  LSql.Free;
end;

function TBaseDatos.ObtenerNuevaClave(ATabla : string) : integer;
var
  qryConsulta : TConsulta;
  LSql        : string;
begin
  result := - 1;

  LSql := 'SELECT clave FROM obtener_nueva_clave(' + QuotedStr(ATabla) + ')';
  qryConsulta := CrearConsulta(LSql);
  with qryConsulta do
  begin
    if (qryConsulta = nil) then
    begin
      // TODO Mensaje Error
      exit;
    end;
    result := ObtenerCampoNombreEntero('clave');
  end;
  LiberarConsulta(qryConsulta);
end;

function TBaseDatos.ExisteTabla(ANombre : string) : boolean;
var
  LSql     : string;
  qryQuery : TConsulta;
begin
  result := false;

  if (FTipoConexion = tcSqlServer) then
    LSql := 'SELECT OBJECT_ID(' + QuotedStr(ANombre) + ') AS table_id'
  else
    LSql := 'SELECT RDB$RELATION_NAME AS table_id FROM RDB$RELATIONS WHERE UPPER(RDB$RELATION_NAME) = ' +
      QuotedStr(UpperCase(ANombre));
  qryQuery := CrearConsulta(LSql);

  if (qryQuery = nil) then
    exit;

  result := qryQuery.FieldByName('table_id').AsString <> EmptyStr;

  LiberarConsulta(qryQuery);
end;

function TBaseDatos.RecuperarConsultaNombre(ANombre : string) : TConsulta;
var
  LPosicion : integer;
begin
  LPosicion := FListaConsultas.IndexOf(ANombre);
  if (LPosicion = - 1) then
    result := nil
  else
    result := TConsulta(FListaConsultas.Objects[LPosicion]);
end;

function TBaseDatos.CrearConsultaCerrada(ASentencia : string; ANombre : string = '') : TConsulta;
var
  LIndice : integer;
begin
  result := nil;

  if (not FConectado) then
  begin
    FUltimoError := 'Ocurrió un error al acceder a una base de datos desconectada';
    EscribirBitacora(UnitName + '.prepareQuery', 'Error no se encuentra conectado a la base de datos: ' + ASentencia,
      tbExcepcion);
    exit;
  end;

  if (ANombre <> '') then
  begin
    // Busca el nombre de la consulta, si la encuentra la recupera en lugar de crearla
    LIndice := FListaConsultas.IndexOf(ANombre);
    if (LIndice >= 0) then
    begin
      result := TConsulta(FListaConsultas.Objects[LIndice]);
      result.Close;
    end;
  end;

  // Si no se encontró la consulta, la crea
  if (result = nil) then
  begin
    result := TConsulta.Create(nil);

    if (ANombre <> '') then
      result.name := ANombre
    else
      result.name := 'qry' + IntToStr(FListaConsultas.Count + 1);
    result.Connection := FConexion;
    FListaConsultas.AddObject(result.name, TObject(result));
  end;

  result.SQL.Clear;
  result.SQL.Add(ASentencia);
end;

function TBaseDatos.DeshacerTransaccion : boolean;
begin
  result := false;
  try
    if (FConexion.InTransaction) then
      FConexion.Rollback
    else
    begin
      FUltimoError := 'Ocurrió un error al intentar deshacer la transacción';
      raise Exception.Create(UnitName + ': No existe una transacción abierta');
    end;
  except
    on e : Exception do
    begin
      FUltimoError := 'Ocurrió un error al intentar deshacer la transacción. ' + e.Message;
      EscribirBitacora(UnitName + '.rollbackTransaction', e.Message, tbExcepcion);
      exit;
    end;
  end;
  result := true;
end;

procedure TBaseDatos.setConectado(const Value : boolean);
begin
  if (Value) then
    Conectar
  else
    Desconectar;
end;

function TBaseDatos.EjecutarScript(AScript : TStringList; AProcedimientoSalida : TProcSalidaScript;
  ASoloErrores : boolean = false; AReintentos : boolean = true) : boolean;
var
  LSql        : TFDScript;
  LCronometro : TStopWatch;
begin
  result := false;

  if (not FConectado) then
  begin
    FUltimoError := 'Ocurrió un error al acceder a una base de datos desconectada';
    EscribirBitacora(UnitName + '.executeSentence', 'Error no se encuentra conectado a la base de datos');
    exit;
  end;

  // Crea el cronómetro
  LCronometro := TStopWatch.StartNew;

  FScriptProcedimiento := AProcedimientoSalida;
  FScriptSoloErrores := ASoloErrores;

  LSql := TFDScript.Create(nil);
  LSql.ScriptOptions.SpoolOutput := smOnAppend;
  LSql.OnSpoolPut := FDScriptSpoolPut;

  try
    // Inicia el conteo de tiempo
    LCronometro.Start;

    LSql.Connection := FConexion;
    LSql.SQLScripts.Clear;
    LSql.SQLScripts.Add;
    LSql.SQLScripts[0].SQL.Assign(AScript);
    LSql.ValidateAll;
    LSql.ExecuteAll;

    // Detiene el conteo de tiempo
    LCronometro.Stop;
    FTiempoEjecucion := LCronometro.ElapsedMilliseconds;
  except
    on e : Exception do
    begin
      FUltimoError := 'Ocurrió un error al intentar afectar a la base de datos. ' + e.Message;

      // Detiene el conteo de tiempo
      LCronometro.Stop;
      FTiempoEjecucion := LCronometro.ElapsedMilliseconds;

      EscribirBitacora(UnitName + '.executeSentence', 'Error al ejecutar la sentencia. ' + e.Message, tbExcepcion);

      // Intenta conectarse de nuevo a la base de datos en caso de desconexión
      if (AReintentos) and (not FConexion.Connected) and (Conectar) then
        result := EjecutarScript(AScript, AProcedimientoSalida, false);

      LSql.Free;
      exit;
    end;
  end;
  LSql.Free;

  result := true;
end;

procedure TBaseDatos.EscribirBitacora(AModulo : string; AMensaje : string;
  ATipoRegistro : TTipoRegistroBitacora = tbExcepcion);
begin
  if (Assigned(FBitacora)) then
    FBitacora.Escribir(AModulo, AMensaje, ATipoRegistro);
end;

procedure TBaseDatos.FDScriptSpoolPut(AEngine : TFDScript; const AMessage : string; AKind : TFDScriptOutputKind);
begin
  if (Assigned(FScriptProcedimiento)) then
    if(not FScriptSoloErrores) or  (AKind = TFDScriptOutputKind.soError) then
      FScriptProcedimiento(AMessage);

  if (AKind = soError) then
    EscribirBitacora(UnitName + '.FDScriptSpoolPut', AMessage, tbExcepcion);
end;

function TBaseDatos.GenerarListaSQL(ASentencia : string; ALista : TStringList) : boolean;
var
  qryConsulta : TConsulta;
begin
  result := false;
  ALista.Clear;

  qryConsulta := CrearConsulta(ASentencia);
  with qryConsulta do
  begin
    if (qryConsulta = nil) then
      exit;

    while (not EsFinDatos) do
    begin
      if (FieldCount > 1) then
        ALista.AddObject(ObtenerCampoCadena(0), TObject(TCampoValor.Create(ObtenerCampoVariante(1))))
      else
      begin
        ALista.Add(ObtenerCampoCadena(0));
      end;
      Siguiente;
    end;
    // ACombo.Items.EndUpdate;
    // if (ASeleccionar) then
    // ACombo.ItemIndex := 0
    // else
    // ACombo.ItemIndex := - 1;
  end;
  LiberarConsulta(qryConsulta);
  result := true;
end;

{ TDBQuery }

destructor TConsulta.Destroy;
begin
  // Libera su datasource
  LiberarFuenteDatos;
  inherited;
end;

function TConsulta.EsFinDatos : boolean;
begin
  result := Eof;
end;

procedure TConsulta.LiberarFuenteDatos;
begin
  if (FDataSource <> nil) then
    FDataSource.Free;
end;

function TConsulta.GetFetchRecordCount : integer;
begin
  FetchAll;
  result := inherited RecordCount;
end;

function TConsulta.GetFuenteDatos : TDataSource;
begin
  if (FDataSource = nil) then
  begin
    FDataSource := TDataSource.Create(nil);
    FDataSource.DataSet := self;
  end;
  result := FDataSource;
end;

procedure TConsulta.LimpiarParametros;
begin
  self.Params.Clear;
end;

procedure TConsulta.Siguiente;
begin
  Next;
end;

procedure TConsulta.AgregarParametro(ANombreParametro : string; AValor : Variant);
var
  prmParam : TFDParam;
begin
  prmParam := self.FindParam(ANombreParametro);

  if (prmParam = nil) then
    prmParam := self.Params.Add;

  prmParam.name := ANombreParametro;

  prmParam.Value := AValor;
end;

function TConsulta.ObtenerCampoNombreEntero(sCampo : string) : integer;
begin
  result := FieldByName(sCampo).AsInteger;
end;

function TConsulta.ObtenerCampoCadena(AIndice : integer) : string;
begin
  result := Fields[AIndice].AsString;
end;

function TConsulta.ObtenerCampoEntero(AIndice : integer) : integer;
begin
  result := Fields[AIndice].AsInteger;
end;

function TConsulta.ObtenerCampoFlotante(AIndice : integer) : double;
begin
  result := Fields[AIndice].AsFloat;
end;

function TConsulta.ObtenerCampoNombreCadena(sCampo : string) : string;
begin
  result := FieldByName(sCampo).AsString;
end;

function TConsulta.ObtenerCampoNombreFlotante(sCampo : string) : real;
begin
  result := FieldByName(sCampo).AsFloat;
end;

function TConsulta.ObtenerCampoVariante(AIndice : integer) : Variant;
begin
  result := Fields[AIndice].AsVariant;
end;

{ TCampoValor }

constructor TCampoValor.Create(const AValor : Variant);
begin
  FValor := AValor;
end;

end.
