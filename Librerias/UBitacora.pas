unit UBitacora;

interface

uses
  System.SysUtils, System.StrUtils;

type
  /// <summary> Enumeraci�n del tipo de registro en la bit�cora</summary>
  TTipoRegistroBitacora = (tlNormal, tlExcepcion, tlEncabezado);

  /// <summary> Clase para generar los registros del sistema </summary>
  /// <remarks> Bismarck Sierra Ibarra 2018-11-26</remarks>
  TBitacora = class
  private
    FNombreArchivo : string;
  public
    /// <summary> Crear� la bit�cora con el nombre de la aplicaci�n y extensi�n .log </summary>
    constructor Create; overload;

    /// <summary> Crear� la bit�cora con el nombre, extensi�n y ruta especificada </summary>
    /// <param name="ANombreArchivo"> Nombre del archivo incluyendo ruta </param>
    constructor Create(ANombreArchivo : string); overload;

    /// <summary> Escribe una entrada en la bit�cora </summary>
    /// <param name="AModulo"> Nombre del m�dulo de donde se manda llamar la bit�cora </param>
    /// <param name="AMensaje"> Mensaje que se desea registrar </param>
    /// <param name="ATipo"> Tipo de registro que se desea hacer </param>
    /// <returns> Verdadero si se escribi� correctamente, false de forma contraria </returns>
    function Escribir(AModulo : string; AMensaje : string; ATipo : TTipoRegistroBitacora = tlNormal) : boolean;
  end;

var
  GBitacora : TBitacora;

implementation

uses
  UUtilerias;
{ TLog }

constructor TBitacora.Create;
const
  EXTENSION_LOG = '.log';
begin
  FNombreArchivo := TUtilerias.AplicacionObtenerNombre;
  FNombreArchivo := TUtilerias.ArchivoObtenerNombreSinExt(FNombreArchivo);
  FNombreArchivo := TUtilerias.AplicacionObtenerRuta + FNombreArchivo + FormatDateTime('yyyymm', Date) + EXTENSION_LOG;
end;

constructor TBitacora.Create(ANombreArchivo : string);
begin
  FNombreArchivo := ANombreArchivo;
end;

function TBitacora.Escribir(AModulo : string; AMensaje : string; ATipo : TTipoRegistroBitacora = tlNormal) : boolean;
const
  ANCHO_ENCABEZADO = 86;
var
  LlArchivo : TextFile;
  LFecha    : string;
begin
  result := true;

  if (Self = nil) then
    exit;

  // No se requiere instanciar la variable para llamar las funciones gen�ricas

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

    LFecha := FormatDateTime(FormatSettings.ShortDateFormat + ' ' + FormatSettings.ShortTimeFormat, Now);
    case ATipo of
      tlNormal :
        Writeln(LlArchivo, LFecha + ' ' + AMensaje);
      tlExcepcion :
        Writeln(LlArchivo, '***EXCEPCION*** ' + LFecha + ' [' + AModulo + '] ' + AMensaje);
      tlEncabezado :
        begin
          Writeln(LlArchivo, StringOfchar('-', ANCHO_ENCABEZADO));
          Writeln(LlArchivo, TUtilerias.CadenaAlinearCentrar(AMensaje, ANCHO_ENCABEZADO));
          Writeln(LlArchivo, TUtilerias.CadenaAlinearCentrar(TUtilerias.AplicacionObtenerNombre + ' ver. ' +
            TUtilerias.AplicacionObtenerVersion + ' - ' + LFecha, ANCHO_ENCABEZADO));
          Writeln(LlArchivo, StringOfchar('-', ANCHO_ENCABEZADO));
        end;
    end;
    CloseFile(LlArchivo);
  except
    result := false;
  end;
end;

end.