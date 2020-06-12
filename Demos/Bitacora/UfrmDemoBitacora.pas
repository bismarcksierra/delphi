unit UfrmDemoBitacora;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, UBitacora;

type
  TForm1 = class(TForm)
    btnIniciarProceso: TBitBtn;
    btnProcesoNormal: TBitBtn;
    btnError: TBitBtn;
    btnDetalle: TBitBtn;
    cmbModoBitacora: TComboBox;
    procedure btnIniciarProcesoClick(Sender : TObject);
    procedure FormCreate(Sender : TObject);
    procedure btnProcesoNormalClick(Sender : TObject);
    procedure btnErrorClick(Sender : TObject);
    procedure btnDetalleClick(Sender: TObject);
    procedure cmbModoBitacoraChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1 : TForm1;

implementation

{$R *.dfm}

procedure TForm1.btnIniciarProcesoClick(Sender : TObject);
begin
  GBitacora.Escribir(UnitName, 'Iniciando', TTipoRegistroBitacora.tbEncabezado);
end;

procedure TForm1.btnProcesoNormalClick(Sender : TObject);
begin
  GBitacora.Escribir(UnitName, 'Cargando variables', TTipoRegistroBitacora.tbNormal);
end;

procedure TForm1.btnErrorClick(Sender : TObject);
var
  LOperacion : integer;
begin
  try
    LOperacion := 1;
    LOperacion := 2 div (LOperacion - 1);
    Caption := FloatToStr(LOperacion);
  except
    on e : exception do
      GBitacora.Escribir(UnitName, 'Error: ' + e.Message, TTipoRegistroBitacora.tbExcepcion);
  end;
end;

procedure TForm1.btnDetalleClick(Sender: TObject);
begin
  GBitacora.Escribir(UnitName, 'Datos sensibles', TTipoRegistroBitacora.tbDetalle);
end;

procedure TForm1.cmbModoBitacoraChange(Sender: TObject);
begin
  GBitacora.ModoBitacora := TModoBitacora(cmbModoBitacora.Items.Objects[cmbModoBitacora.ItemIndex])
end;

procedure TForm1.FormCreate(Sender : TObject);
var
  LModosBitacora : TStrings;
begin
  GBitacora := TBitacora.Create(TVigenciaBitacora.vbMes);
  // GBitacora := TBitacora.Create('C:\Log\Nombre.log');

  LModosBitacora:= GBitacora.RecuperarModosBitacora;
  cmbModoBitacora.Items.Assign(LModosBitacora);
  cmbModoBitacora.ItemIndex := 0;
  LModosBitacora.Free;
end;

end.
