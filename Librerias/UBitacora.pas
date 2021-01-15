//-----------------------------------------------------------------------------
//                           Delphi Runtime Library
//                 Copyright(c) 2021 Bismarck Sierra Ibarra
//                       bismarcksierra@gmail.com
//-----------------------------------------------------------------------------
unit UBitacora;

interface

uses
  System.SysUtils, System.StrUtils, System.SyncObjs, System.Classes;

type
  /// <summary> Enumeración del tipo de registro en la bitácora</summary>
  TTipoRegistroBitacora = (tbEncabezado, tbExcepcion, tbNormal, tbDetalle);

  /// <summary> Enumeración del tipo de registro en la bitácora</summary>
  TModoBitacora = (mbNormal, mbDetalle, mbSoloErrores);

  /// <summary> Enumeración del tipo de registro en la bitácora</summary>
  TVigenciaBitacora = (vbMes, vbDia, vbInfinito);

  /// <summary> Clase para generar los registros del sistema </summary>
  /// <remarks> Bismarck Sierra Ibarra 2018-11-26</remarks>
  TBitacora = class
  private
    FFechaBitacora      : TDate;
    FNombreArchivo      : string;
    FModoBitacora       : TModoBitacora;
    FTiempoVidaBitacora : TVigenciaBitacora;

    /// <summary> Escribe una entrada en la bitácora </summary>
    /// <param name="AModulo"> Nombre del módulo de donde se manda llamar la bitácora </param>
    /// <param name="AMensaje"> Mensaje que se desea registrar </param>
    /// <param name="ATipo"> Tipo de registro que se desea hacer </param>
    /// <returns> Verdadero si se escribió correctamente, false de forma contraria </returns>
    function Guardar(AModulo : string; AMensaje : string; ATipo : TTipoRegistroBitacora) : boolean;

    /// <summary> Establece el nombre del archivo de acuerdo a la vigencia  </summary>
    procedure EstablecerNombreBitacora;
  public
    /// <summary> constructor para crear la bitácora con el nombre de la aplicación y extensión .log </summary>
    /// <param name="ATiempoVida"> Especifica si la bitácora se llevará por día, mes o infito </param>
    constructor Create(ATiempoVida : TVigenciaBitacora); overload;

    /// <summary> constuctor para crear la bitácora con el nombre, extensión y ruta especificada </summary>
    /// <param name="ANombreArchivo"> Nombre del archivo incluyendo ruta </param>
    constructor Create(ANombreArchivo : string); overload;

    /// <summary> Metodo público para guardar un registro en la bitácora</summary>
    /// <param name="AModulo"> Nombre del módulo de donde se manda llamar la bitácora </param>
    /// <param name="AMensaje"> Mensaje que se desea registrar </param>
    /// <param name="ATipo"> Tipo de registro que se desea hacer </param>
    /// <returns> Verdadero si se escribió correctamente, false de forma contraria </returns>
    function Escribir(AModulo : string; AMensaje : string; ATipo : TTipoRegistroBitacora = tbNormal) : boolean;

    /// <summary> Recupera los modos de bitácora para desplegarlos en un combo o listbox</summary>
    /// <returns> Listado con los modos de pago en texto y la enumeración del modo de bitácora como objeto </returns>
    function RecuperarModosBitacora : TStrings;

    /// <summary> Especifica el modo de registro de la bitácora, por default es mbNormal </summary>
    property ModoBitacora : TModoBitacora read FModoBitacora write FModoBitacora;
  end;

const
  TModoBitacoraDescripcion : array [TModoBitacora] of string = ('Normal', 'Detalle', 'Solo errrores');

var
  GBitacora : TBitacora;

implementation

uses
  UUtilerias;
{ TLog }

constructor TBitacora.Create(ATiempoVida : TVigenciaBitacora);
begin
  FModoBitacora := mbNormal;
  FTiempoVidaBitacora := ATiempoVida;

  EstablecerNombreBitacora;
end;

procedure TBitacora.EstablecerNombreBitacora;
const
  EXTENSION_LOG  = '.log';
  DIRECTORIO_LOG = 'Log\';
var
  LDirectorio          : string;
  LDirectorioAgrupador : string;
  LTiempoVida          : string;
begin
  FFechaBitacora := Date;
  FNombreArchivo := TUtilerias.AplicacionObtenerNombre;
  FNombreArchivo := TUtilerias.ArchivoObtenerNombreSinExt(FNombreArchivo);
  case FTiempoVidaBitacora of
    vbMes :
      begin
        LTiempoVida := FormatDateTime('yyyymm', Date);
        // Agrupa por año
        LDirectorioAgrupador := FormatDateTime('yyyy', Date) + '\';
      end;
    vbDia :
      begin
        LTiempoVida := FormatDateTime('yyyymmdd', Date);
        // Agrupa por mes
        LDirectorioAgrupador := FormatDateTime('yyyymm', Date) + '\';
      end;
    vbInfinito :
      begin
        LTiempoVida := '';
        LDirectorioAgrupador := '';
      end;
  end;
  LDirectorio := TUtilerias.AplicacionObtenerRuta + DIRECTORIO_LOG + LDirectorioAgrupador;
  ForceDirectories(LDirectorio);
  FNombreArchivo := LDirectorio + FNombreArchivo + LTiempoVida + EXTENSION_LOG;
end;

function TBitacora.Guardar(AModulo, AMensaje : string; ATipo : TTipoRegistroBitacora) : boolean;
const
  ANCHO_ENCABEZADO = 100;
var
  LFecha    : string;
  LlArchivo : TextFile;
begin
  Result := true;
  try
    AssignFile(LlArchivo, FNombreArchivo);
    // Lectura/escritura
    FileMode := fmOpenWrite;
    if (not FileExists(FNombreArchivo)) then
      // Crea el archivo
      Rewrite(LlArchivo)
    else
      // Abre el archivo para agregar
      Reset(LlArchivo);

    Append(LlArchivo);

    LFecha := FormatDateTime('dd/mm/yyyy hh:nn:ss', Now);
    case ATipo of
      tbDetalle :
        Writeln(LlArchivo, LFecha + ' [' + AModulo + '] ' + AMensaje);
      tbEncabezado :
        begin
          Writeln(LlArchivo, StringOfchar('-', ANCHO_ENCABEZADO));
          Writeln(LlArchivo, TUtilerias.CadenaAlinearCentrar(AMensaje, ANCHO_ENCABEZADO));
          Writeln(LlArchivo, TUtilerias.CadenaAlinearCentrar(TUtilerias.AplicacionObtenerNombre + ' ver. ' +
            TUtilerias.AplicacionObtenerVersion + ' - ' + LFecha, ANCHO_ENCABEZADO));
          Writeln(LlArchivo, StringOfchar('-', ANCHO_ENCABEZADO));
        end;
      tbExcepcion :
        Writeln(LlArchivo, '***EXCEPCION*** ' + LFecha + ' [' + AModulo + '] ' + AMensaje);
      tbNormal :
        Writeln(LlArchivo, LFecha + ' [' + AModulo + '] ' + AMensaje);
    end;
    CloseFile(LlArchivo);
  except
    result := false;
  end;
end;

function TBitacora.RecuperarModosBitacora : TStrings;
var
  I : TModoBitacora;
  LTexto : string;
begin
  Result := TStringList.Create;
  for I := low(TModoBitacora) to high(TModoBitacora) do
  begin
    LTexto := TModoBitacoraDescripcion[I];
    Result.AddObject(LTexto, TObject(I));
  end;
end;

constructor TBitacora.Create(ANombreArchivo : string);
begin
  FNombreArchivo := ANombreArchivo;

  // Si se especifica el nombre, se guardará siempre bajo ese nombre, no se renovará el archivo por fecha
  FTiempoVidaBitacora := vbInfinito;
  FModoBitacora := mbNormal;
end;

function TBitacora.Escribir(AModulo : string; AMensaje : string; ATipo : TTipoRegistroBitacora = tbNormal) : boolean;
begin
  result := true;

  if (Self = nil) then
    exit;

  // si cambia la fecha y se debe cambiar el nombre del archivo
  if (FTiempoVidaBitacora = TVigenciaBitacora.vbMes) and
    (FormatDateTime('yyyymm', FFechaBitacora) < FormatDateTime('yyyymm', Date)) then
    EstablecerNombreBitacora
  else if (FTiempoVidaBitacora = TVigenciaBitacora.vbDia) and
    (FormatDateTime('yyyymmdd', FFechaBitacora) < FormatDateTime('yyyymmd', Date)) then
    EstablecerNombreBitacora;

  // Determina si se va a guardar en archivo o se ignorará
  case ATipo of
    tbDetalle :
      if (FModoBitacora = mbDetalle) then
        result := Guardar(AModulo, AMensaje, ATipo);
    tbEncabezado, tbExcepcion :
      result := Guardar(AModulo, AMensaje, ATipo);
    tbNormal :
      if (FModoBitacora <> mbSoloErrores) then
        result := Guardar(AModulo, AMensaje, ATipo);
  else
    result := false;
  end;
end;

end.
