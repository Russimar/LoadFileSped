unit uPrincipal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, ACBrBase, ACBrSpedPisCofins,
  Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.ComCtrls, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS,
  FireDAC.Phys.Intf, FireDAC.DApt.Intf, Data.DB, Vcl.Grids, Vcl.DBGrids,
  SMDBGrid, FireDAC.Comp.DataSet, FireDAC.Comp.Client, cxGraphics, cxControls,
  cxLookAndFeels, cxLookAndFeelPainters, cxContainer, cxEdit, dxSkinsCore,
  dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee,
  dxSkinDarkRoom, dxSkinDarkSide, dxSkinDevExpressDarkStyle,
  dxSkinDevExpressStyle, dxSkinFoggy, dxSkinGlassOceans, dxSkinHighContrast,
  dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky,
  dxSkinMcSkin, dxSkinMetropolis, dxSkinMetropolisDark, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black,
  dxSkinOffice2010Blue, dxSkinOffice2010Silver, dxSkinOffice2013DarkGray,
  dxSkinOffice2013LightGray, dxSkinOffice2013White, dxSkinOffice2016Colorful,
  dxSkinOffice2016Dark, dxSkinPumpkin, dxSkinSeven, dxSkinSevenClassic,
  dxSkinSharp, dxSkinSharpPlus, dxSkinSilver, dxSkinSpringTime, dxSkinStardust,
  dxSkinSummer2008, dxSkinTheAsphaltWorld, dxSkinsDefaultPainters,
  dxSkinValentine, dxSkinVisualStudio2013Blue, dxSkinVisualStudio2013Dark,
  dxSkinVisualStudio2013Light, dxSkinVS2010, dxSkinWhiteprint,
  dxSkinXmas2008Blue, cxTextEdit, cxCurrencyEdit, Vcl.DBCtrls;
type
  TC100 = record
    Data : TDate;
    Valor : Real;
    CodigoParticipante : String;
    Tipo : String;
    EntSai : String;
    Chave : String;
  end;

type
  TForm1 = class(TForm)
    pnlTop: TPanel;
    pnlPrincipal: TPanel;
    OpenDialog1: TOpenDialog;
    edtFileName: TEdit;
    Label1: TLabel;
    SpeedButton2: TSpeedButton;
    btnGerar: TBitBtn;
    ProgressBar1: TProgressBar;
    MemDados: TFDMemTable;
    MemDadosDataEmissao: TDateField;
    MemDadosValorCupom: TFloatField;
    MemDadosConta: TStringField;
    MemDadosTipo: TStringField;
    SMDBGrid1: TSMDBGrid;
    dsPadrao: TDataSource;
    memParticipante: TFDMemTable;
    memParticipanteCodigo: TStringField;
    memParticipanteNome: TStringField;
    MemDadosCodigoParticipante: TStringField;
    MemDadosEntSai: TStringField;
    MemDadosNomeParticipante: TStringField;
    MemDadosChave: TStringField;
    pnlMensagem: TPanel;
    pnlBotton: TPanel;
    MemDadosValorEntrada: TFloatField;
    MemDadosValorSaida: TFloatField;
    MemDadosValorGeral: TFloatField;
    MemDadosValorTotalCupom: TAggregateField;
    MemDadosValorTotalSaida: TAggregateField;
    MemDadosValorTotalEntrada: TAggregateField;
    txtEntrada: TDBText;
    Label2: TLabel;
    Label3: TLabel;
    txtSaida: TDBText;
    Label5: TLabel;
    txtCupom: TDBText;
    procedure ImportarArquivo;
    procedure SpeedButton2Click(Sender: TObject);
    procedure MemDadosCalcFields(DataSet: TDataSet);
    procedure btnGerarClick(Sender: TObject);
  private
    { Private declarations }
   function NumeroDeLinhasTXT(FilePath:String): Integer;
   function ValidaData(aValue : String) : Boolean;
   function StrZero(Valor : string; Quant : Integer): string;
  public
    { Public declarations }
    procedure GerarArquivoCobol;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.btnGerarClick(Sender: TObject);
begin
  GerarArquivoCobol;
end;

procedure TForm1.GerarArquivoCobol;
var
  ArquivoCobol : TextFile;
  linha : String;
  CaminhoArquivo : String;
  xData, xValor, xConta, xTipo, xParticipante : String;
begin
  MemDados.DisableControls;
  try
    ProgressBar1.Min := 0;
    ProgressBar1.Max := MemDados.RecordCount;
    CaminhoArquivo := ExtractFilePath(Application.ExeName) + 'SPDCOBOL.TXT';
    AssignFile(ArquivoCobol,CaminhoArquivo);
    Rewrite(ArquivoCobol);
    MemDados.First;
    while not MemDados.Eof do
    begin
      ProgressBar1.Position := ProgressBar1.Position + 1;
      xValor := StringReplace(FormatFloat('0.00', MemDadosValorGeral.AsFloat),',','',[rfReplaceAll]);
      xValor := StrZero(xValor,11);
      xConta := MemDadosConta.AsString;
      xData := FormatDateTime('ddmmyyyy', MemDadosDataEmissao.AsDateTime);
      xTipo := MemDadosEntSai.AsString;
      xParticipante := memParticipanteNome.AsString;
      linha := xData +
               copy(xConta,1,20) +
               copy(xValor,1,11) +
               copy(xTipo,1,1) +
               copy(xParticipante,1,50);
      Writeln(ArquivoCobol, linha);
      MemDados.Next;
    end;



  finally
    MemDados.EnableControls;
    pnlMensagem.Caption := 'Arquivo Gerado com Sucesso!';
  end;
end;

procedure TForm1.ImportarArquivo;
var
  ArquivoSped : TextFile;
  Contador, I : Integer;
  Linha : String;
  Registro : String;
  Gravou : Boolean;
  x, dia, mes, ano : String;
  C100 : TC100;
  function MontaValor : String;
  begin
    Result := '';
    i := pos('|', Linha);
    if i = 0 then
      i := Length(Linha) + 1;
    Result := Trim(Copy(Linha, 1, i - 1));
    Delete(Linha, 1, i);
  end;
begin
  MemDados.CreateDataSet;
  MemDados.EmptyDataSet;

  memParticipante.CreateDataSet;
  memParticipante.EmptyDataSet;

  AssignFile(ArquivoSped, OpenDialog1.FileName);
  try
    Reset(ArquivoSped);
    Readln(ArquivoSped, Linha);
    while not Eoln(ArquivoSped) do
    begin
      ProgressBar1.Position := ProgressBar1.Position + 1;
      Registro := copy(Linha,2, 4);
      if registro = '0150' then
      begin
        MontaValor;
        MontaValor;
        memParticipante.Append;
        memParticipanteCodigo.AsString := MontaValor;
        memParticipanteNome.AsString := MontaValor;
        memParticipante.Post;
      end;

      if Registro = 'C100' then
      begin
        begin
          try
            MontaValor; //Tira pipe
            MontaValor; //Tira pipe
            x := MontaValor;
            if x = '0' then
              c100.EntSai := 'E';
            if x = '1' then
              c100.EntSai := 'S';
            MontaValor;
            C100.CodigoParticipante := MontaValor;
            C100.Tipo := MontaValor;
            Contador := 1;
            while Contador < 4 do
            begin
             MontaValor;
             inc(Contador);
            end;
            C100.Chave := MontaValor;
            x := MontaValor;
            dia := Copy(x,1,2);
            mes := Copy(x,3,2);
            ano := Copy(x,5,4);
            x := (dia + '/' + mes + '/' + ano);
            if ValidaData(x) then
            begin
              C100.Data := StrToDate(x);
            end
            else
              Continue;
          except
            on E : Exception do
            begin
              ShowMessage('Erro na chave :' + C100.Chave + #13#10 + e.Message);
            end;
          end;
        end;
      end;
      if Registro = 'C170' then
      begin
        try
        MemDados.Append;
          MemDadosDataEmissao.AsDateTime := C100.Data;
          MemDadosTipo.AsString := C100.Tipo;
          MemDadosCodigoParticipante.AsString := C100.CodigoParticipante;
          if memParticipante.Locate('codigo',C100.CodigoParticipante,[loCaseInsensitive]) then
            MemDadosNomeParticipante.AsString := memParticipanteNome.AsString;
          MemDadosEntSai.AsString := C100.EntSai;
          MemDadosChave.AsString := C100.Chave;
          Contador := 1;
          while Contador < 8 do
          begin
           MontaValor;
           inc(Contador);
          end;
          if c100.EntSai = 'E' then
            MemDadosValorEntrada.AsFloat := StrToFloat(MontaValor)
          else
            MemDadosValorSaida.AsFloat := StrToFloat(MontaValor);

          Contador := 1;
          while Contador < 30 do
          begin
           MontaValor;
           inc(Contador);
          end;
          MemDadosConta.AsString := MontaValor;
          MemDados.Post;
        except
          on E : Exception do
          begin
            ShowMessage('Erro na chave :' + C100.Chave + #13#10 + e.Message);
          end;
        end;
      end;
      if Registro = 'C175' then
      begin
        try
          MemDados.Append;
          MemDadosDataEmissao.AsDateTime := C100.Data;
          MemDadosTipo.AsString := C100.Tipo;
          MemDadosChave.AsString := C100.Chave;
          MemDadosCodigoParticipante.AsString := C100.CodigoParticipante;
          if memParticipante.Locate('codigo',C100.CodigoParticipante,[loCaseInsensitive]) then
            MemDadosNomeParticipante.AsString := memParticipanteNome.AsString;
          MemDadosEntSai.AsString := C100.EntSai;
          Contador := 1;
          while Contador < 4 do
          begin
           MontaValor;
           inc(Contador);
          end;
          MemDadosValorCupom.AsFloat := StrToFloat(MontaValor);
          Contador := 1;
          while Contador < 14 do
          begin
           MontaValor;
           inc(Contador);
          end;
          MemDadosConta.AsString := MontaValor;
          MemDados.Post;
        except
          on E : Exception do
          begin
            ShowMessage('Erro na chave :' + C100.Chave + #13#10 + e.Message);
          end;
        end;
      end;

      Readln(ArquivoSped, Linha);
      inc(Contador);
    end;
  finally
    CloseFile(ArquivoSped);
  end;


end;

procedure TForm1.MemDadosCalcFields(DataSet: TDataSet);
begin
  MemDadosValorGeral.AsFloat := MemDadosValorCupom.AsFloat + MemDadosValorEntrada.AsFloat + MemDadosValorSaida.AsFloat;
end;

function TForm1.NumeroDeLinhasTXT(FilePath: String): Integer;
var
  aList : TStringList;
begin
  if FileExists(FilePath) then
  begin
    aList := TStringList.Create;
    try
      aList.LoadFromFile(FilePath);
      Result := aList.Count;
    finally
      FreeAndNil(aList);
    end;
  end
  else
    Result := 0;
end;

procedure TForm1.SpeedButton2Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
  begin
    pnlMensagem.Caption := 'Aguarde...Lendo Arquivo Texto';
    try
      pnlMensagem.Update;
      edtFileName.Text := OpenDialog1.FileName;
      edtFileName.Update;
      ProgressBar1.Min := 0;
      ProgressBar1.Max := NumeroDeLinhasTXT(edtFileName.Text);
      ImportarArquivo;
    except
      pnlMensagem.Caption := 'Erro ao ler o arquivo';
    end;
    pnlMensagem.Caption := 'Arquivo Gerado!';
    ProgressBar1.Position := 0;
  end;
end;

function TForm1.StrZero(Valor: string; Quant: Integer): string;
begin
  Result := Valor;
  Quant := Quant - Length(Result);
  if Quant > 0 then
    Result := StringOfChar('0', Quant) + Result;
end;

function TForm1.ValidaData(aValue: String): Boolean;
var
  Data : TDate;
begin
  try
    Data := StrToDate(aValue);
    Result := True;
  except
    Result := False;
  end;
end;

end.
