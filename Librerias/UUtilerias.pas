Unit UUtilerias;

Interface

Uses
  System.SysUtils, Winapi.Windows, System.Classes, WinSvc, WinSock, System.Win.ComObj;

Const

  ESTADO_SERVICIO_DETENIDO   = 'Detenido';
  ESTADO_SERVICIO_DETENIENDO = 'Deteniendo';
  ESTADO_SERVICIO_EN_EJECUCION = 'En ejecución';
  ESTADO_SERVICIO_INICIANDO = 'Iniciando';
  ESTADO_SERVICIO_NO_DISPONIBLE = 'No disponible';
  ESTADO_SERVICIO_PAUSADO = 'Pausado';

Type
  TProcSalida = Procedure(ATextoSalida : String) Of Object;

  /// <summary> Clase con métodos genéricos </summary>
  /// <remarks> Bismarck Sierra Ibarra 2019-01-10</remarks>
  TUtilerias = Class
  Private
  Public
    /// <summary> Elimina archivos o directorios </summary>
    /// <param name="AArchivos"> Archivo, directorio o filtro de archivos a eliminar </param>
    /// <param name="ARecursivo"> Indica si se eliminará de forma recursiva en caso de ser directorio </param>
    /// <returns> Verdadero si se eliminó exitosamente, falso de forma contraria </returns>
    Class Function ArchivosEliminar(AArchivos : String; ARecursivo : boolean = false) : boolean;

    /// <summary> Obtiene el nombre de un archivo sin su extensión </summary>
    /// <param name="AArchivo"> Archivo al que se desea quitar la extensión </param>
    /// <returns> Nombre de la aplicación sin extensión y sin ruta </returns>
    Class Function ArchivoObtenerNombreSinExt(AArchivo : String) : String;

    /// <summary> Obtiene el nombre de la aplicación </summary>
    /// <returns> Nombre de la aplicación sin ruta </returns>
    Class Function AplicacionObtenerNombre : String;

    /// <summary> Obtiene el nombre y ruta de la aplicación </summary>
    /// <returns> Nombre de la aplicación con su ruta</returns>
    Class Function AplicacionObtenerNombreRuta : String;

    /// <summary> Obtiene la ruta donde se encuentra la aplicación </summary>
    /// <returns> Ruta de la aplicación con diagonal invertida al final </returns>
    Class Function AplicacionObtenerRuta : String;

    /// <summary> Obtiene la versión de la aplicación </summary>
    /// <returns> Versión de la aplicación </returns>
    Class Function AplicacionObtenerVersion : String;

    /// <summary> Centra un texto en una cadena de cierta longitud</summary>
    /// <param name="ACadena"> Cadena a centrar </param>
    /// <param name="AAncho"> Ancho de la cadena final </param>
    /// <returns> Cadena centrada </returns>
    Class Function CadenaAlinearCentrar(ACadena : String; AAncho : integer) : String;

    /// <summary> Obtiene el serial del disco duro </summary>
    /// <returns> Número de serie de la unidad c:\ </returns>
    Class Function DiscoDuroObtenerSerial : String;

    /// <summary> Valida si un correo electrónico es correcto </summary>
    /// <param name="ACorreo"> Correo electrónico a validar </param>
    /// <returns> Verdadero si es correcto, falso en caso contrario </returns>
    Class Function CorreoValidar(ACorreo : String) : boolean;

    /// <summary> Abre un puerto en el cortafuegos de Windows</summary>
    /// <param name="ANombreEntrada"> Nombre de la regla de entrada del cortafuegos</param>
    /// <param name="ANumeroPuerto"> Puerto que se desea abrir </param>
    /// <param name="AProtocoloTCP"> Indica si es portocolo TCP </param>
    /// <param name="AAmbitoLocal"> Indica si la regla es local o pública </param>
    /// <param name="AActivar"> Es pecifica si se desea activar la regla</param>
    Class Procedure CortaFuegosAbrirPuerto(ANombreEntrada : String; ANumeroPuerto : integer; AProtocoloTCP : boolean;
      AAmbitoLocal : boolean; AActivar : boolean);

    /// <summary> Obtiene el nombre de la computadora </summary>
    /// <returns> Nombre del equipoo </returns>
    Class Function NombrePCObtener : String;

    /// <summary> Ejecuta un comando de MSDOS y envía los resultados de consola como un parámetro a un procedimiento</summary>
    /// <param name="AComando"> Comando que se desea ejecutar </param>
    /// <param name="AParametros"> Parametros que se le enviarán al comando </param>
    /// <param name="AProcedimiento"> Procedimiento que se llamará cada vez que se genere una salida</param>
    Class Procedure MsdosEjecutar(Const AComando, AParametros : String; AProcedimiento : TProcSalida);

    /// <summary> Detiene un servicio de Windows </summary>
    /// <param name="ANombreServicio"> Nombre del servicio a detener </param>
    /// <returns> Verdadero si se logró detener el servicio  </returns>
    Class Function ServicioDetener(ANombreServicio : String) : boolean;

    /// <summary> Inicia un servicio de Windows </summary>
    /// <param name="ANombreServicio"> Nombre del servicio a iniciar </param>
    /// <returns> Verdadero si se logró iniciar el servicio  </returns>
    Class Function ServicioIniciar(ANombreServicio : String) : boolean;

    /// <summary> Obtiene el estado de un servicio de Windows </summary>
    /// <param name="ANombreServicio"> Nombre del servicio a verificar </param>
    /// <returns> Estado del servicio:  </returns>
    Class Function ServicioObtenerEstado(ANombreServicio : String) : String;
  End;

Implementation

{ TUtilerias }

Class Function TUtilerias.CadenaAlinearCentrar(ACadena : String; AAncho : integer) : String;
Var
  LEspacios : integer;
Begin
  Result := ACadena.Trim;
  LEspacios := (AAncho - Length(Result)) Div 2;
  Result := StringOfChar(' ', LEspacios) + Result + StringOfChar(' ', LEspacios);
End;

Class Function TUtilerias.AplicacionObtenerNombre : String;
Begin
  Result := ExtractFileName(AplicacionObtenerNombreRuta);
  Result := Copy(Result, 1, Length(Result) - Length(ExtractFileExt(Result)));
End;

Class Function TUtilerias.AplicacionObtenerNombreRuta : String;
Var
  LNombre : Array [0 .. MAX_PATH] Of char;
Begin
  FillChar(LNombre, sizeof(LNombre), #0);
  GetModuleFileName(hInstance, LNombre, sizeof(LNombre));
  Result := LNombre;
End;

Class Function TUtilerias.AplicacionObtenerRuta : String;
Begin
  Result := ExtractFilePath(AplicacionObtenerNombreRuta);
End;

Class Function TUtilerias.AplicacionObtenerVersion : String;
Var
  LNombre          : Array [0 .. MAX_PATH] Of char;
  LInformacion     : PChar;
  LTamanio         : DWORD;
  LTamanioBuffer   : DWORD;
  LConsultaVersion : String;
  LValor           : Pointer;
Begin
  LValor := Nil;
  FillChar(LNombre, sizeof(LNombre), #0);
  GetModuleFileName(hInstance, LNombre, sizeof(LNombre));

  LTamanioBuffer := GetFileVersionInfoSize(LNombre, LTamanioBuffer);
  If LTamanioBuffer > 0 Then
  Begin
    LInformacion := AllocMem(LTamanioBuffer);
    Try
      GetFileVersionInfo(LNombre, 0, LTamanioBuffer, LInformacion);

      VerQueryValue(LInformacion, '\VarFileInfo\Translation', LValor, LTamanio);
      If LTamanio > 0 Then
      Begin
        LConsultaVersion := Format('\StringFileInfo\%.4x%.4x\%s', [LoWord(integer(LValor^)), HiWord(integer(LValor^)),
          'FileVersion']);

        If VerQueryValue(LInformacion, PChar(LConsultaVersion), LValor, LTamanio) Then
          Result := StrPas(PChar(LValor));
      End;
    Finally
      FreeMem(LInformacion, LTamanioBuffer);
    End;
  End;
End;

Class Function TUtilerias.ArchivoObtenerNombreSinExt(AArchivo : String) : String;
Begin
  Result := ExtractFileName(AArchivo);
  Result := Copy(Result, 1, Length(Result) - Length(ExtractFileExt(Result)));
End;

Class Function TUtilerias.DiscoDuroObtenerSerial : String;
Var
  a, b      : DWORD;
  Buffer    : Array [0 .. MAX_PATH] Of char;
  SerialNum : DWORD;
Begin
  Result := 'NULL';
  If GetVolumeInformation(PChar(ExtractFileDrive(ParamStr(0))), Buffer, sizeof(Buffer), @SerialNum, a, b, Nil, 0) Then
    Result := Format('%8.8X', [SerialNum]);
End;

Class Function TUtilerias.CorreoValidar(ACorreo : String) : boolean;
Const
  // Caracteres válidos en un "átomo"
  ATOM_CHARS = [#33 .. #255] - ['(', ')', '<', '>', '@', ',', ';', ':', '\', '/', '"', '.', '[', ']', #127];
  // Caracteres válidos en una "cadena-entrecomillada"
  QUOTED_STRING_CHARS = [#0 .. #255] - ['"', #13, '\'];
  // Caracteres válidos en un subdominio
  LETTERS         = ['A' .. 'Z', 'a' .. 'z'];
  LETTERS_DIGITS  = ['0' .. '9', 'A' .. 'Z', 'a' .. 'z'];
  SUBDOMAIN_CHARS = ['-', '0' .. '9', 'A' .. 'Z', 'a' .. 'z'];

Type
  States = (STATE_BEGIN, STATE_ATOM, STATE_QTEXT, STATE_QCHAR, STATE_QUOTE, STATE_LOCAL_PERIOD,
    STATE_EXPECTING_SUBDOMAIN, STATE_SUBDOMAIN, STATE_HYPHEN);
Var
  State             : States;
  i, n, iSubdomains : integer;
  c                 : char;
Begin
  State := STATE_BEGIN;
  n := Length(ACorreo);
  i := 1;
  iSubdomains := 1;
  While (i <= n) Do
  Begin
    c := ACorreo[i];
    Case State Of
      STATE_BEGIN :
        If CharInSet(c, ATOM_CHARS) Then
          State := STATE_ATOM
        Else If c = '"' Then
          State := STATE_QTEXT
        Else
          break;
      STATE_ATOM :
        If c = '@' Then
          State := STATE_EXPECTING_SUBDOMAIN
        Else If c = '.' Then
          State := STATE_LOCAL_PERIOD
        Else If Not CharInSet(c, ATOM_CHARS) Then
          break;
      STATE_QTEXT :
        If c = '\' Then
          State := STATE_QCHAR
        Else If c = '"' Then
          State := STATE_QUOTE
        Else If Not CharInSet(c, QUOTED_STRING_CHARS) Then
          break;
      STATE_QCHAR :
        State := STATE_QTEXT;
      STATE_QUOTE :
        If c = '@' Then
          State := STATE_EXPECTING_SUBDOMAIN
        Else If c = '.' Then
          State := STATE_LOCAL_PERIOD
        Else
          break;
      STATE_LOCAL_PERIOD :
        If CharInSet(c, ATOM_CHARS) Then
          State := STATE_ATOM
        Else If c = '"' Then
          State := STATE_QTEXT
        Else
          break;
      STATE_EXPECTING_SUBDOMAIN :
        If CharInSet(c, LETTERS_DIGITS) Then
          State := STATE_SUBDOMAIN
        Else
          break;
      STATE_SUBDOMAIN :
        If c = '.' Then
        Begin
          Inc(iSubdomains);
          State := STATE_EXPECTING_SUBDOMAIN
        End
        Else If c = '-' Then
          State := STATE_HYPHEN
        Else If Not CharInSet(c, LETTERS_DIGITS) Then
          break;
      STATE_HYPHEN :
        If CharInSet(c, LETTERS_DIGITS) Then
          State := STATE_SUBDOMAIN
        Else If c <> '-' Then
          break;
    End;
    Inc(i);
  End;
  If i <= n Then
    Result := false
  Else
    Result := (State = STATE_SUBDOMAIN) And (iSubdomains >= 2);

  // si sCorreo esta vacio regresa true
  If ACorreo = '' Then
    Result := true;
End;

Class Function TUtilerias.ArchivosEliminar(AArchivos : String; ARecursivo : boolean = false) : boolean;
Var
  srcBusqueda : TSearchRec;
Begin
  Result := true;
  If System.SysUtils.FindFirst(AArchivos, faAnyFile, srcBusqueda) <> 0 Then
    exit;

  Repeat
    If (srcBusqueda.Attr And faDirectory <> 0) Then
    Begin
      If (srcBusqueda.Name <> '.') And (srcBusqueda.Name <> '..') Then
        If (ARecursivo) Then
        Begin
          ArchivosEliminar(ExtractFilePath(AArchivos) + srcBusqueda.Name + '\' + ExtractFileName(AArchivos),
            ARecursivo);
          System.SysUtils.RemoveDir(ExtractFilePath(AArchivos) + srcBusqueda.Name);
        End;
    End
    Else
      Try
        System.SysUtils.DeleteFile(ExtractFilePath(AArchivos) + srcBusqueda.Name);
      Except
        Result := false;
      End;
  Until System.SysUtils.FindNext(srcBusqueda) <> 0;
  System.SysUtils.FindClose(srcBusqueda);
End;

Class Procedure TUtilerias.MsdosEjecutar(Const AComando, AParametros : String; AProcedimiento : TProcSalida);
Const
  CReadBuffer = 2400;
Var
  saSecurity : TSecurityAttributes;
  hRead      : THandle;
  hWrite     : THandle;
  suiStartup : TStartupInfo;
  piProcess  : TProcessInformation;
  pBuffer    : Array [0 .. CReadBuffer] Of AnsiChar;
  dBuffer    : Array [0 .. CReadBuffer] Of AnsiChar;
  dRead      : DWORD;
  dRunning   : DWORD;
  dAvailable : DWORD;
Begin
  saSecurity.nLength := sizeof(TSecurityAttributes);
  saSecurity.bInheritHandle := true;
  saSecurity.lpSecurityDescriptor := Nil;
  If CreatePipe(hRead, hWrite, @saSecurity, 0) Then
    Try
      FillChar(suiStartup, sizeof(TStartupInfo), #0);
      suiStartup.cb := sizeof(TStartupInfo);
      suiStartup.hStdInput := hRead;
      suiStartup.hStdOutput := hWrite;
      suiStartup.hStdError := hWrite;
      suiStartup.dwFlags := STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW;
      suiStartup.wShowWindow := SW_HIDE;
      If CreateProcess(Nil, PChar(AComando + ' ' + AParametros), @saSecurity, @saSecurity, true, NORMAL_PRIORITY_CLASS,
        Nil, Nil, suiStartup, piProcess) Then
        Try
          Repeat
            dRunning := WaitForSingleObject(piProcess.hProcess, 100);
            PeekNamedPipe(hRead, Nil, 0, Nil, @dAvailable, Nil);
            If (dAvailable > 0) Then
              Repeat
                dRead := 0;
                ReadFile(hRead, pBuffer[0], CReadBuffer, dRead, Nil);
                pBuffer[dRead] := #0;
                OemToCharA(pBuffer, dBuffer);
                AProcedimiento(String(pBuffer));
              Until (dRead < CReadBuffer);
          Until (dRunning <> WAIT_TIMEOUT);
        Finally
          CloseHandle(piProcess.hProcess);
          CloseHandle(piProcess.hThread);
        End;
    Finally
      CloseHandle(hRead);
      CloseHandle(hWrite);
    End;
End;

Class Function TUtilerias.ServicioObtenerEstado(ANombreServicio : String) : String;
Var
  schm, schs : SC_Handle;
  ss         : TServiceStatus;
  dwStat     : DWORD;
  LEquipo    : String;
Begin
  dwStat := 0;
  LEquipo := NombrePCObtener;

  schm := OpenSCManager(PWideChar(LEquipo), Nil, SC_MANAGER_CONNECT);
  If (schm > 0) Then
  Begin
    schs := OpenService(schm, PWideChar(ANombreServicio), SERVICE_QUERY_STATUS);
    If (schs > 0) Then
    Begin
      If (QueryServiceStatus(schs, ss)) Then
      Begin
        dwStat := ss.dwCurrentState;
      End;
      CloseServiceHandle(schs);
    End;
    CloseServiceHandle(schm);
  End;
  Case dwStat Of
    0 :
      Result := ESTADO_SERVICIO_NO_DISPONIBLE;
    1 :
      Result := ESTADO_SERVICIO_DETENIDO;
    2 :
      Result := ESTADO_SERVICIO_INICIANDO;
    3 :
      Result := ESTADO_SERVICIO_DETENIENDO;
    4 :
      Result := ESTADO_SERVICIO_EN_EJECUCION;
    7 :
      Result := ESTADO_SERVICIO_PAUSADO;
  Else
    Result := IntToStr(dwStat);
  End;
End;

Class Function TUtilerias.NombrePCObtener : String;
Var
  LNombre     : AnsiString;
  DatosSocket : WSAData;
Begin
  WSAStartup($0101, DatosSocket);
  SetLength(LNombre, MAX_PATH);
  gethostname(PAnsiChar(LNombre), MAX_PATH);
  SetLength(LNombre, StrLen(PAnsiChar(LNombre)));
  Result := String(LNombre);
End;

Class Procedure TUtilerias.CortaFuegosAbrirPuerto(ANombreEntrada : String; ANumeroPuerto : integer;
  AProtocoloTCP : boolean; AAmbitoLocal : boolean; AActivar : boolean);
Const
  NET_FW_PROFILE_DOMAIN     = 0;
  NET_FW_PROFILE_STANDARD   = 1;
  NET_FW_IP_VERSION_ANY     = 2;
  NET_FW_IP_PROTOCOL_UDP    = 17;
  NET_FW_IP_PROTOCOL_TCP    = 6;
  NET_FW_SCOPE_ALL          = 0;
  NET_FW_SCOPE_LOCAL_SUBNET = 1;
Var
  fwMgr, objPuerto : OleVariant;
  perfil           : OleVariant;
Begin
  fwMgr := CreateOLEObject('HNetCfg.FwMgr');
  perfil := fwMgr.LocalPolicy.CurrentProfile;
  objPuerto := CreateOLEObject('HNetCfg.FWOpenPort');
  objPuerto.Name := ANombreEntrada;
  If AProtocoloTCP Then
    objPuerto.Protocol := NET_FW_IP_PROTOCOL_TCP
  Else
    objPuerto.Protocol := NET_FW_IP_PROTOCOL_UDP;
  objPuerto.Port := ANumeroPuerto;
  If AAmbitoLocal Then
    objPuerto.Scope := NET_FW_SCOPE_LOCAL_SUBNET
  Else
    objPuerto.Scope := NET_FW_SCOPE_ALL;
  objPuerto.Enabled := true;
  perfil.GloballyOpenPorts.Add(objPuerto);
End;

Class Function TUtilerias.ServicioIniciar(ANombreServicio : String) : boolean;
Var
  schm, schs : SC_Handle;
  ss         : TServiceStatus;
  psTemp     : PChar;
  dwChkP     : DWord;
  LEquipo    : String;
Begin
  LEquipo := '\\127.0.0.1';
  ss.dwCurrentState := 1;
  schm := OpenSCManager(PChar(LEquipo), SERVICES_ACTIVE_DATABASE, SC_MANAGER_CONNECT);

  If (schm > 0) Then
  Begin

    schs := OpenService(schm, PChar(ANombreServicio), SERVICE_START Or SERVICE_QUERY_STATUS);

    If (schs > 0) Then
    Begin
      psTemp := Nil;
      If (StartService(schs, 0, psTemp)) Then
        If (QueryServiceStatus(schs, ss)) Then
          While (SERVICE_RUNNING <> ss.dwCurrentState) Do
          Begin
            dwChkP := ss.dwCheckPoint;
            Sleep(ss.dwWaitHint);
            If (Not QueryServiceStatus(schs, ss)) Then
              Break;
            If (ss.dwCheckPoint < dwChkP) Then
              Break;
          End;
      CloseServiceHandle(schs);
    End;
    CloseServiceHandle(schm);
  End;

  Result := SERVICE_RUNNING = ss.dwCurrentState;
End;

Class Function TUtilerias.ServicioDetener(ANombreServicio : String) : boolean;
Var
  schm, schs : SC_Handle;
  ss         : TServiceStatus;
  dwChkP     : DWord;
  LEquipo    : String;
Begin

  schm := OpenSCManager(PChar(LEquipo), Nil, SC_MANAGER_CONNECT);

  If (schm > 0) Then
  Begin

    schs := OpenService(schm, PChar(ANombreServicio), SERVICE_STOP Or SERVICE_QUERY_STATUS);
    If (schs > 0) Then
    Begin
      If (ControlService(schs, SERVICE_CONTROL_STOP, ss)) Then
        If (QueryServiceStatus(schs, ss)) Then
          While (SERVICE_STOPPED <> ss.dwCurrentState) Do
          Begin
            dwChkP := ss.dwCheckPoint;
            Sleep(ss.dwWaitHint);
            If (Not QueryServiceStatus(schs, ss)) Then
              Break;
            If (ss.dwCheckPoint < dwChkP) Then
              Break;
          End;
      CloseServiceHandle(schs);
    End;

    CloseServiceHandle(schm);
  End;

  Result := SERVICE_STOPPED = ss.dwCurrentState;
End;

End.
