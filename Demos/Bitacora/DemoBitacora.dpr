program DemoBitacora;

uses
  Vcl.Forms,
  UfrmDemoBitacora in 'UfrmDemoBitacora.pas' {Form1},
  UBitacora in '..\..\Librerias\UBitacora.pas',
  UUtilerias in '..\..\Librerias\UUtilerias.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
