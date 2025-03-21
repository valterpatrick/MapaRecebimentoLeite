{ ****************************************************************************** }
{ Projeto: Componentes ACBr }
{ Biblioteca de componentes Delphi para geração de arquivos de mapa de }
{ recebimento de leite. }
{ }
{ Direitos Autorais Reservados (c) 2023 Valter Patrick Silva Ferreira }
{ }
{ Esta biblioteca é distribuída na expectativa de que seja útil, porém, SEM }
{ NENHUMA GARANTIA; nem mesmo a garantia implícita de COMERCIABILIDADE OU }
{ ADEQUAÇÃO A UMA FINALIDADE ESPECÍFICA. Consulte a Licença Pública Geral Menor }
{ do GNU para mais detalhes. (Arquivo LICENÇA.TXT ou LICENSE.TXT) }
{ }
{ ****************************************************************************** }

{ ******************************************************************************
  |* Histórico
  |*
  |* 24/08/2023: Valter Patrick Silva Ferreira
  |*  - Criação e distribuição da Primeira Versão
  ****************************************************************************** }

unit ACBrMapaRecLeite;

interface

uses Classes, SysUtils, Contnrs, StrUtils, DateUtils, Math, ComObj, XMLDoc, XMLIntf;

const
  REGISTRO100 = '100|NM_DECLARANTE|CD_CNPJ|CD_IE|DS_EMAIL|';
  REGISTRO200 = '200|CD_PRODUTOR_IE|DT_RECEBIMENTO|QT_LITROS|CD_PLACA|';
  REGISTRO300 = '300|CD_PRODUTOR_IE|DT_NF|NR_NF|CD_SERIE|CD_CHAVE|FL_RESPONSABILIDADE|QT_LITROS|VR_TOTAL_NF|VR_MERCADORIA|VR_FRETE|VR_BC|VR_INCENTIVO|VR_DEDUCOES|VR_ICMS|';
  REGISTRO400 = '400|CD_PRODUTOR_IE|CD_PRODUTOR_CPF|NM_PRODUTOR|';
  REGISTROFIM = '999|FIM|';
  PIPE = '|';

type
  TACBrRespFrete = (rfLaticinio, rfProdutor);
  TACBrModoArquivo = (maProducao, maHomologacao);
  TACBrTipoArquivo = (taTXT, taXML, taExcel);

  TDeclarante = class
  private
    FNomeDeclarante: String;
    FCNPJ: String;
    FInscEstadual: String;
    FEmail: String;
  public
    property NomeDeclarante: String read FNomeDeclarante write FNomeDeclarante;
    property CNPJ: String read FCNPJ write FCNPJ;
    property InscEstadual: String read FInscEstadual write FInscEstadual;
    property Email: String read FEmail write FEmail;
  end;

  TRecebimentosLeite = class
  private
    FProdutorInscEstadual: String;
    FDataRecebimento: TDate;
    FQuantLitros: Double;
    FPlaca: String;
  public
    property ProdutorInscEstadual: String read FProdutorInscEstadual write FProdutorInscEstadual;
    property DataRecebimento: TDate read FDataRecebimento write FDataRecebimento;
    property QuantLitros: Double read FQuantLitros write FQuantLitros;
    property Placa: String read FPlaca write FPlaca;
  end;

  TRecebimentosLeiteLista = class(TObjectList)
  protected
    procedure SetObject(Index: Integer; Item: TRecebimentosLeite);
    function GetObject(Index: Integer): TRecebimentosLeite;
    procedure Insert(Index: Integer; Obj: TRecebimentosLeite);
  public
    function New: TRecebimentosLeite;
    property Objects[Index: Integer]: TRecebimentosLeite read GetObject; default;
  end;

  TNotasFiscais = class
  private
    FProdutorInscEstadual: String;
    FDataNFe: TDate;
    FNumero: Integer;
    FSerie: String;
    FChave: String;
    FResponsabilidadeFrete: TACBrRespFrete;
    FQuantLitros: Double;
    FTotalNFe: Currency;
    FValorMercadorias: Currency;
    FValorFrete: Currency;
    FValorBase: Currency;
    FValorIncentivo: Currency;
    FValorDeducoes: Currency;
    FValorICMS: Currency;
  public
    property ProdutorInscEstadual: String read FProdutorInscEstadual write FProdutorInscEstadual;
    property DataNFe: TDate read FDataNFe write FDataNFe;
    property Numero: Integer read FNumero write FNumero;
    property Serie: String read FSerie write FSerie;
    property Chave: String read FChave write FChave;
    property ResponsabilidadeFrete: TACBrRespFrete read FResponsabilidadeFrete write FResponsabilidadeFrete;
    property QuantLitros: Double read FQuantLitros write FQuantLitros;
    property TotalNFe: Currency read FTotalNFe write FTotalNFe;
    property ValorMercadorias: Currency read FValorMercadorias write FValorMercadorias;
    property ValorFrete: Currency read FValorFrete write FValorFrete;
    property ValorBase: Currency read FValorBase write FValorBase;
    property ValorIncentivo: Currency read FValorIncentivo write FValorIncentivo;
    property ValorDeducoes: Currency read FValorDeducoes write FValorDeducoes;
    property ValorICMS: Currency read FValorICMS write FValorICMS;
  end;

  TNotasFiscaisLista = class(TObjectList)
  protected
    procedure SetObject(Index: Integer; Item: TNotasFiscais);
    function GetObject(Index: Integer): TNotasFiscais;
    procedure Insert(Index: Integer; Obj: TNotasFiscais);
  public
    function New: TNotasFiscais;
    property Objects[Index: Integer]: TNotasFiscais read GetObject; default;
  end;

  TProdutores = class
  private
    FProdutorInscEstadual: String;
    FProdutorCPF: String;
    FProdutorNome: String;
  public
    property ProdutorInscEstadual: String read FProdutorInscEstadual write FProdutorInscEstadual;
    property ProdutorCPF: String read FProdutorCPF write FProdutorCPF;
    property ProdutorNome: String read FProdutorNome write FProdutorNome;
  end;

  TProdutoresLista = class(TObjectList)
  protected
    procedure SetObject(Index: Integer; Item: TProdutores);
    function GetObject(Index: Integer): TProdutores;
    procedure Insert(Index: Integer; Obj: TProdutores);
  public
    function New: TProdutores;
    property Objects[Index: Integer]: TProdutores read GetObject; default;
  end;

  TACBrMapaRecLeite = class(TComponent)
  private
    FModoArquivo: TACBrModoArquivo;
    FTipoArquivo: TACBrTipoArquivo;
    FAnoReferencia: Integer;
    FMesReferencia: Integer;
    FNomeArquivo: String;
    FDirArquivo: String;
    Arquivo: TextFile;
    ExcApp: Variant;
    XMLDoc: TXMLDocument;
    NodeRaiz: IXMLNode;
    FDeclarante: TDeclarante;
    FRecebimentosLeiteLista: TRecebimentosLeiteLista;
    FNotasFiscaisLista: TNotasFiscaisLista;
    FProdutoresLista: TProdutoresLista;
    FRejeicoes: TStringList;
    FValidarRegistrosAntesGerar: Boolean;
    FResponsabilidadeFretePadrao: TACBrRespFrete;
    procedure WriteRecord(Registro, Linha: String);
    procedure GeraCabecalhoXML;
    procedure GeraPeriodoXML;
    procedure GeraRegistroFIM;
    procedure GeraDeclarante;
    procedure GeraRecebimentosLeite;
    procedure GeraNotasFiscais;
    procedure GeraProdutores;
    procedure ValidarDeclarante;
    procedure ValidarRecebimentosLeite;
    procedure ValidarNotasFiscais;
    procedure ValidarProdutores;
    function GetDataRefIni: TDate;
    function GetDataRefFim: TDate;
    function DateTimeToDate(Data: TDateTime): TDate;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property Declarante: TDeclarante read FDeclarante write FDeclarante;
    property RecebimentosLeiteLista: TRecebimentosLeiteLista read FRecebimentosLeiteLista write FRecebimentosLeiteLista;
    property NotasFiscaisLista: TNotasFiscaisLista read FNotasFiscaisLista write FNotasFiscaisLista;
    property ProdutoresLista: TProdutoresLista read FProdutoresLista write FProdutoresLista;
    property ModoArquivo: TACBrModoArquivo read FModoArquivo write FModoArquivo;
    property TipoArquivo: TACBrTipoArquivo read FTipoArquivo write FTipoArquivo;
    property ResponsabilidadeFretePadrao: TACBrRespFrete read FResponsabilidadeFretePadrao write FResponsabilidadeFretePadrao;
    property Rejeicoes: TStringList read FRejeicoes;
    property NomeArquivo: String read FNomeArquivo;
    property DirArquivo: String read FDirArquivo write FDirArquivo;
    property AnoReferencia: Integer read FAnoReferencia write FAnoReferencia;
    property MesReferencia: Integer read FMesReferencia write FMesReferencia;
    property ValidarRegistrosAntesGerar: Boolean read FValidarRegistrosAntesGerar write FValidarRegistrosAntesGerar;
    property DataIniRef: TDate read GetDataRefIni;
    property DataFimRef: TDate read GetDataRefFim;
    procedure LimparRegistros;
    function GeraNomeArquivo(ModoArq: TACBrModoArquivo; InscEstadual: String; AnoRef, MesRef: Integer): String;
    function GeraExtensao(FTipoArquivo: TACBrTipoArquivo): String;
    function GerarArquivo: String;
    function ValidarRegistros: Boolean;
    function RespFreteTipoToStr(resp: TACBrRespFrete): String;
    function RespFreteTipoToStrDesc(resp: TACBrRespFrete): String;
    function ModoArquivoTipoToStr(modArq: TACBrModoArquivo): String;
    function TipoArquivoTipoToStr(TipoArq: TACBrTipoArquivo): String;
  end;

procedure Register;

implementation

function LenghtNativeString(const AString: String): Integer;
begin
{$IFDEF FPC}
  Result := UTF8Length(AString);
{$ELSE}
  Result := Length(AString);
{$ENDIF}
end;

function PadRight(const AString: String; const nLen: Integer;
  const Caracter: Char): String;
var
  Tam: Integer;
begin
  Tam := LenghtNativeString(AString);
  if Tam < nLen then
    Result := AString + StringOfChar(Caracter, (nLen - Tam))
  else
    Result := LeftStr(AString, nLen);
end;

{ -----------------------------------------------------------------------------
  Completa <AString> com <Caracter> a esquerda, até o tamanho <nLen>, Alinhando
  a <AString> a Direita. Se <AString> for maior que <nLen>, ela será truncada
  ---------------------------------------------------------------------------- }
function PadLeft(const AString: String; const nLen: Integer;
  const Caracter: Char): String;
var
  Tam: Integer;
begin
  Tam := LenghtNativeString(AString);
  if Tam < nLen then
    Result := StringOfChar(Caracter, (nLen - Tam)) + AString
  else
    Result := LeftStr(AString, nLen); // RightStr(AString,nLen) ;
end;

function TiraPontos(Str: string): string;
var
  i, Count: Integer;
begin
  SetLength(Result, Length(Str));
  Count := 0;
  for i := 1 to Length(Str) do
  begin
    if not CharInSet(Str[i], ['/', ',', '-', '.', ')', '(', ' ']) then
    begin
      inc(Count);
      Result[Count] := Str[i];
    end;
  end;
  SetLength(Result, Count);
end;

procedure Register;
begin
  RegisterComponents('ACBrMapaRecLeite', [TACBrMapaRecLeite]);
end;

{ TACBrMapaRecLeite }

constructor TACBrMapaRecLeite.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FModoArquivo := maProducao;
  FTipoArquivo := taTXT;
  FRejeicoes := TStringList.Create;
  FValidarRegistrosAntesGerar := True;
  FResponsabilidadeFretePadrao := rfLaticinio;

  FDeclarante := TDeclarante.Create;
  FRecebimentosLeiteLista := TRecebimentosLeiteLista.Create;
  FNotasFiscaisLista := TNotasFiscaisLista.Create;
  FProdutoresLista := TProdutoresLista.Create;
end;

function TACBrMapaRecLeite.DateTimeToDate(Data: TDateTime): TDate;
begin
  Result := EncodeDate(YearOf(Data), MonthOf(Data), DayOf(Data));
end;

destructor TACBrMapaRecLeite.Destroy;
begin
  FRejeicoes.Free;
  FDeclarante.Free;
  FRecebimentosLeiteLista.Free;
  FNotasFiscaisLista.Free;
  FProdutoresLista.Free;
  inherited Destroy;
end;

procedure TACBrMapaRecLeite.GeraCabecalhoXML;
begin
  NodeRaiz := XMLDoc.AddChild('operacoesRecebimentoLeite');
end;

procedure TACBrMapaRecLeite.GeraRecebimentosLeite;
var
  Linha, i, j, k: Integer;
  vRegistro: String;
  NodeMapaRecebimentos, NodeProdutor, NodeDadosProdutor, NodeNotasFiscais, NodeNota, NodeMovimentacaoDiaria, NodeDia: IXMLNode;
begin
  if FTipoArquivo = taTXT then
  begin
    if FRecebimentosLeiteLista.Count > 0 then
      WriteRecord('200', REGISTRO200);
    for i := 0 to FRecebimentosLeiteLista.Count - 1 do
    begin
      vRegistro := '211|';
      vRegistro := vRegistro + PadLeft(TiraPontos(Trim(FRecebimentosLeiteLista[i].FProdutorInscEstadual)), 13, '0') + PIPE;
      vRegistro := vRegistro + FormatDateTime('dd/mm/yyyy', FRecebimentosLeiteLista[i].FDataRecebimento) + PIPE;
      vRegistro := vRegistro + FormatFloat('0.00', FRecebimentosLeiteLista[i].FQuantLitros) + PIPE;
      vRegistro := vRegistro + PadRight(TiraPontos(Trim(FRecebimentosLeiteLista[i].FPlaca)), 7, ' ') + PIPE;

      WriteRecord('211', vRegistro);
    end;
    if FRecebimentosLeiteLista.Count > 0 then
      WriteRecord('299', '299|' + FRecebimentosLeiteLista.Count.ToString + '|');
  end
  else if FTipoArquivo = taExcel then
  begin
    ExcApp.Workbooks[1].Sheets[2].Name := 'Recebimentos-Leite';

    // Cabeçalho
    ExcApp.Workbooks[1].Sheets[2].Range['A1', 'D' + (FRecebimentosLeiteLista.Count + 1).ToString].Font.Name := 'Calibri'; // Fonte
    ExcApp.Workbooks[1].Sheets[2].Range['A1', 'D' + (FRecebimentosLeiteLista.Count + 1).ToString].Font.Size := 11; // Tamanho da Fonte
    ExcApp.Workbooks[1].Sheets[2].Range['A2', 'D' + (FRecebimentosLeiteLista.Count + 1).ToString].NumberFormat := AnsiChar('@');

    ExcApp.Workbooks[1].Sheets[2].Cells[1, 1] := 'CD_PRODUTOR_IE';
    ExcApp.Workbooks[1].Sheets[2].Range['A1', 'A' + (FRecebimentosLeiteLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[2].Range['A1', 'A' + (FRecebimentosLeiteLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[2].Range['A1', 'A' + (FRecebimentosLeiteLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[2].Cells[1, 2] := 'DT_RECEBIMENTO';
    ExcApp.Workbooks[1].Sheets[2].Range['B1', 'B' + (FRecebimentosLeiteLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[2].Range['B1', 'B' + (FRecebimentosLeiteLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[2].Range['B1', 'B' + (FRecebimentosLeiteLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }
    ExcApp.Workbooks[1].Sheets[2].Range['B1', 'B' + (FRecebimentosLeiteLista.Count + 1).ToString].NumberFormat := 'dd/mm/aaaa';

    ExcApp.Workbooks[1].Sheets[2].Cells[1, 3] := 'QT_LITROS';
    ExcApp.Workbooks[1].Sheets[2].Range['C1', 'C' + (FRecebimentosLeiteLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[2].Range['C1', 'C' + (FRecebimentosLeiteLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[2].Range['C1', 'C' + (FRecebimentosLeiteLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[2].Cells[1, 4] := 'CD_PLACA';
    ExcApp.Workbooks[1].Sheets[2].Range['D1', 'D' + (FRecebimentosLeiteLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[2].Range['D1', 'D' + (FRecebimentosLeiteLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[2].Range['D1', 'D' + (FRecebimentosLeiteLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    // Dados
    Linha := 2;
    for i := 0 to FRecebimentosLeiteLista.Count - 1 do
    begin
      ExcApp.Workbooks[1].Sheets[2].Cells[Linha, 1] := PadLeft(TiraPontos(Trim(FRecebimentosLeiteLista[i].FProdutorInscEstadual)), 13, '0');
      ExcApp.Workbooks[1].Sheets[2].Cells[Linha, 2] := FormatDateTime('dd/mm/yyyy', FRecebimentosLeiteLista[i].FDataRecebimento);
      ExcApp.Workbooks[1].Sheets[2].Cells[Linha, 3] := FormatFloat('0.00', FRecebimentosLeiteLista[i].FQuantLitros);
      ExcApp.Workbooks[1].Sheets[2].Cells[Linha, 4] := TiraPontos(Trim(FRecebimentosLeiteLista[i].FPlaca));

      Linha := Linha + 1;
    end;
  end
  else if FTipoArquivo = taXML then
  begin
    NodeMapaRecebimentos := NodeRaiz.AddChild('mapaRecebimentos');
    for i := 0 to FProdutoresLista.Count - 1 do
    begin
      NodeProdutor := NodeMapaRecebimentos.AddChild('produtor');
      NodeProdutor.Attributes['IE'] := PadLeft(TiraPontos(Trim(FProdutoresLista[i].FProdutorInscEstadual)), 13, '0');

      NodeDadosProdutor := NodeProdutor.AddChild('dadosProdutor');
      NodeDadosProdutor.ChildValues['cd_produtor_ie'] := PadLeft(TiraPontos(Trim(FProdutoresLista[i].FProdutorInscEstadual)), 13, '0');
      NodeDadosProdutor.ChildValues['cd_produtor_cpf'] := PadLeft(TiraPontos(Trim(FProdutoresLista[i].FProdutorCPF)), 11, '0');
      NodeDadosProdutor.ChildValues['nm_produtor'] := Trim(FProdutoresLista[i].FProdutorNome);

      NodeNotasFiscais := NodeProdutor.AddChild('notasFiscais');
      for j := 0 to FNotasFiscaisLista.Count - 1 do
      begin
        if TiraPontos(Trim(FNotasFiscaisLista[j].FProdutorInscEstadual)) = TiraPontos(Trim(FProdutoresLista[i].FProdutorInscEstadual)) then
        begin
          NodeNota := NodeNotasFiscais.AddChild('nota');
          NodeNota.ChildValues['dt_nf'] := FormatDateTime('dd/mm/yyyy', FNotasFiscaisLista[j].FDataNFe);
          NodeNota.ChildValues['nr_nf'] := PadLeft(IntToStr(FNotasFiscaisLista[j].FNumero), 9, '0');
          NodeNota.ChildValues['cd_serie'] := Trim(FNotasFiscaisLista[j].FSerie);
          NodeNota.ChildValues['cd_chave'] := PadLeft(TiraPontos(Trim(FNotasFiscaisLista[j].FChave)), 44, '0');
          NodeNota.ChildValues['fl_responsabilidade'] := RespFreteTipoToStr(FNotasFiscaisLista[j].FResponsabilidadeFrete);
          NodeNota.ChildValues['qt_litros'] := FormatFloat('0.00', FNotasFiscaisLista[j].FQuantLitros);
          NodeNota.ChildValues['vr_total_nf'] := FormatFloat('0.00', FNotasFiscaisLista[j].FTotalNFe);
          NodeNota.ChildValues['vr_mercadoria'] := FormatFloat('0.00', FNotasFiscaisLista[j].FValorMercadorias);
          NodeNota.ChildValues['vr_frete'] := FormatFloat('0.00', FNotasFiscaisLista[j].FValorFrete);
          NodeNota.ChildValues['vr_bc'] := FormatFloat('0.00', FNotasFiscaisLista[j].FValorBase);
          NodeNota.ChildValues['vr_incentivo'] := FormatFloat('0.00', FNotasFiscaisLista[j].FValorIncentivo);
          NodeNota.ChildValues['vr_deducoes'] := FormatFloat('0.00', FNotasFiscaisLista[j].FValorDeducoes);
          NodeNota.ChildValues['vr_icms'] := FormatFloat('0.00', FNotasFiscaisLista[j].FValorICMS);
        end;
      end;
      NodeMovimentacaoDiaria := NodeProdutor.AddChild('movimentacaoDiaria');
      for k := 0 to FRecebimentosLeiteLista.Count - 1 do
      begin
        if TiraPontos(Trim(FRecebimentosLeiteLista[k].FProdutorInscEstadual)) = TiraPontos(Trim(FProdutoresLista[i].FProdutorInscEstadual)) then
        begin
          NodeDia := NodeMovimentacaoDiaria.AddChild('d' + PadLeft(IntToStr(DayOf(FRecebimentosLeiteLista[k].DataRecebimento)), 2, '0'));
          NodeDia.ChildValues['dt_recebimento'] := FormatDateTime('dd/mm/yyyy', FRecebimentosLeiteLista[k].FDataRecebimento);
          NodeDia.ChildValues['qt_litros'] := FormatFloat('0.00', FRecebimentosLeiteLista[k].FQuantLitros);
          NodeDia.ChildValues['cd_placa'] := TiraPontos(Trim(FRecebimentosLeiteLista[k].FPlaca));
        end;
      end;
    end;
  end;
end;

procedure TACBrMapaRecLeite.GeraRegistroFIM;
begin
  WriteRecord('999', REGISTROFIM);
end;

procedure TACBrMapaRecLeite.GeraPeriodoXML;
var
  NodePeriodo: IXMLNode;
begin
  if FTipoArquivo = taXML then
  begin
    NodePeriodo := NodeRaiz.AddChild('periodo');
    NodePeriodo.ChildValues['inicio'] := FormatDateTime('dd/mm/yyyy', GetDataRefIni);
    NodePeriodo.ChildValues['fim'] := FormatDateTime('dd/mm/yyyy', GetDataRefFim);
  end;
end;

procedure TACBrMapaRecLeite.GeraDeclarante;
var
  vRegistro: String;
  NodeDeclarante: IXMLNode;
begin
  if FTipoArquivo = taTXT then
  begin
    WriteRecord('100', REGISTRO100);
    vRegistro := '111|';
    vRegistro := vRegistro + Trim(FDeclarante.FNomeDeclarante) + PIPE; // Não especificado o tamanho do campo
    vRegistro := vRegistro + PadLeft(TiraPontos(Trim(FDeclarante.FCNPJ)), 14, '0') + PIPE;
    vRegistro := vRegistro + PadLeft(TiraPontos(Trim(FDeclarante.FInscEstadual)), 13, '0') + PIPE;
    vRegistro := vRegistro + Trim(FDeclarante.FEmail) + PIPE; // Não especificado o tamanho do campo

    WriteRecord('111', vRegistro);
  end
  else if FTipoArquivo = taXML then
  begin
    NodeDeclarante := NodeRaiz.AddChild('dadosDeclarante');
    NodeDeclarante.ChildValues['nm_declarante'] := Trim(FDeclarante.FNomeDeclarante);
    NodeDeclarante.ChildValues['cd_cnpj'] := PadLeft(TiraPontos(Trim(FDeclarante.FCNPJ)), 14, '0');
    NodeDeclarante.ChildValues['cd_ie'] := PadLeft(TiraPontos(Trim(FDeclarante.FInscEstadual)), 13, '0');
    NodeDeclarante.ChildValues['ds_email'] := Trim(FDeclarante.FEmail);
  end;
end;

function TACBrMapaRecLeite.GeraNomeArquivo(ModoArq: TACBrModoArquivo; InscEstadual: String; AnoRef, MesRef: Integer): String;
begin
  Result := 'MRL_' + IfThen(ModoArq = maProducao, 'P_01_', 'H_01_') + PadLeft(TiraPontos(Trim(InscEstadual)), 13, '0') + '_' +
    PadLeft(IntToStr(AnoRef), 4, '0') + PadLeft(IntToStr(MesRef), 2, '0');
end;

function TACBrMapaRecLeite.GeraExtensao(FTipoArquivo: TACBrTipoArquivo): String;
begin
  if FTipoArquivo = taTXT then
    Result := '.txt'
  else if FTipoArquivo = taExcel then
    Result := '.xlsx'
  else if FTipoArquivo = taXML then
    Result := '.xml'
  else
    Result := '.dat';
end;

procedure TACBrMapaRecLeite.GeraNotasFiscais;
var
  vRegistro: String;
  Linha, i: Integer;
begin
  if FTipoArquivo = taTXT then
  begin
    if FNotasFiscaisLista.Count > 0 then
      WriteRecord('300', REGISTRO300);
    for i := 0 to FNotasFiscaisLista.Count - 1 do
    begin
      vRegistro := '311|';
      vRegistro := vRegistro + PadLeft(TiraPontos(Trim(FNotasFiscaisLista[i].FProdutorInscEstadual)), 13, '0') + PIPE;
      vRegistro := vRegistro + FormatDateTime('dd/mm/yyyy', FNotasFiscaisLista[i].FDataNFe) + PIPE;
      vRegistro := vRegistro + PadLeft(IntToStr(FNotasFiscaisLista[i].FNumero), 9, '0') + PIPE;
      vRegistro := vRegistro + Trim(FNotasFiscaisLista[i].FSerie) + PIPE;
      vRegistro := vRegistro + PadLeft(TiraPontos(Trim(FNotasFiscaisLista[i].FChave)), 44, '0') + PIPE;
      vRegistro := vRegistro + RespFreteTipoToStr(FNotasFiscaisLista[i].FResponsabilidadeFrete) + PIPE;
      vRegistro := vRegistro + FormatFloat('0.00', FNotasFiscaisLista[i].FQuantLitros) + PIPE;
      vRegistro := vRegistro + FormatFloat('0.00', FNotasFiscaisLista[i].FTotalNFe) + PIPE;
      vRegistro := vRegistro + FormatFloat('0.00', FNotasFiscaisLista[i].FValorMercadorias) + PIPE;
      vRegistro := vRegistro + FormatFloat('0.00', FNotasFiscaisLista[i].FValorFrete) + PIPE;
      vRegistro := vRegistro + FormatFloat('0.00', FNotasFiscaisLista[i].FValorBase) + PIPE;
      vRegistro := vRegistro + FormatFloat('0.00', FNotasFiscaisLista[i].FValorIncentivo) + PIPE;
      vRegistro := vRegistro + FormatFloat('0.00', FNotasFiscaisLista[i].FValorDeducoes) + PIPE;
      vRegistro := vRegistro + FormatFloat('0.00', FNotasFiscaisLista[i].FValorICMS) + PIPE;

      WriteRecord('311', vRegistro);
    end;
    if FNotasFiscaisLista.Count > 0 then
      WriteRecord('399', '399|' + FNotasFiscaisLista.Count.ToString + '|');
  end
  else if FTipoArquivo = taExcel then
  begin
    ExcApp.Workbooks[1].Sheets[3].Name := 'Notas Fiscais - Globais';

    // Cabeçalho
    ExcApp.Workbooks[1].Sheets[3].Range['A1', 'N' + (FNotasFiscaisLista.Count + 1).ToString].Font.Name := 'Calibri'; // Fonte
    ExcApp.Workbooks[1].Sheets[3].Range['A1', 'N' + (FNotasFiscaisLista.Count + 1).ToString].Font.Size := 11; // Tamanho da Fonte
    ExcApp.Workbooks[1].Sheets[3].Range['A2', 'N' + (FNotasFiscaisLista.Count + 1).ToString].NumberFormat := AnsiChar('@');

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 1] := 'CD_PRODUTOR_IE';
    ExcApp.Workbooks[1].Sheets[3].Range['A1', 'A' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[3].Range['A1', 'A' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['A1', 'A' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 2] := 'DT_NF';
    ExcApp.Workbooks[1].Sheets[3].Range['B1', 'B' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[3].Range['B1', 'B' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['B1', 'B' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }
    ExcApp.Workbooks[1].Sheets[3].Range['B1', 'B' + (FNotasFiscaisLista.Count + 1).ToString].NumberFormat := 'dd/mm/aaaa';

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 3] := 'NR_NF';
    ExcApp.Workbooks[1].Sheets[3].Range['C1', 'C' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[3].Range['C1', 'C' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['C1', 'C' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 4] := 'CD_SERIE';
    ExcApp.Workbooks[1].Sheets[3].Range['D1', 'D' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[3].Range['D1', 'D' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['D1', 'D' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 5] := 'CD_CHAVE';
    ExcApp.Workbooks[1].Sheets[3].Range['E1', 'E' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 100;
    ExcApp.Workbooks[1].Sheets[3].Range['E1', 'E' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['E1', 'E' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 6] := 'FL_RESPONSABILIDADE';
    ExcApp.Workbooks[1].Sheets[3].Range['F1', 'F' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 50;
    ExcApp.Workbooks[1].Sheets[3].Range['F1', 'F' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['F1', 'F' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 7] := 'QT_LITROS';
    ExcApp.Workbooks[1].Sheets[3].Range['G1', 'G' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[3].Range['G1', 'G' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['G1', 'G' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 8] := 'VR_TOTAL_NF';
    ExcApp.Workbooks[1].Sheets[3].Range['H1', 'H' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[3].Range['H1', 'H' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['H1', 'H' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 9] := 'VR_MERCADORIA';
    ExcApp.Workbooks[1].Sheets[3].Range['I1', 'I' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[3].Range['I1', 'I' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['I1', 'I' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 10] := 'VR_FRETE';
    ExcApp.Workbooks[1].Sheets[3].Range['J1', 'J' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[3].Range['J1', 'J' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['J1', 'J' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 11] := 'VR_BC';
    ExcApp.Workbooks[1].Sheets[3].Range['K1', 'K' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[3].Range['K1', 'K' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['K1', 'K' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 12] := 'VR_INCENTIVO';
    ExcApp.Workbooks[1].Sheets[3].Range['L1', 'L' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[3].Range['L1', 'L' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['L1', 'L' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 13] := 'VR_DEDUCOES';
    ExcApp.Workbooks[1].Sheets[3].Range['M1', 'M' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[3].Range['M1', 'M' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['M1', 'M' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[3].Cells[1, 14] := 'VR_ICMS';
    ExcApp.Workbooks[1].Sheets[3].Range['N1', 'N' + (FNotasFiscaisLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[3].Range['N1', 'N' + (FNotasFiscaisLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[3].Range['N1', 'N' + (FNotasFiscaisLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    // Dados
    Linha := 2;
    for i := 0 to FNotasFiscaisLista.Count - 1 do
    begin
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 1] := PadLeft(TiraPontos(Trim(FNotasFiscaisLista[i].FProdutorInscEstadual)), 13, '0');
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 2] := FormatDateTime('dd/mm/yyyy', FNotasFiscaisLista[i].FDataNFe);
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 3] := PadLeft(IntToStr(FNotasFiscaisLista[i].FNumero), 9, '0');
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 4] := Trim(FNotasFiscaisLista[i].FSerie);
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 5] := PadLeft(TiraPontos(Trim(FNotasFiscaisLista[i].FChave)), 44, '0');
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 6] := RespFreteTipoToStr(FNotasFiscaisLista[i].FResponsabilidadeFrete);
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 7] := FormatFloat('0.00', FNotasFiscaisLista[i].FQuantLitros);
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 8] := FormatFloat('0.00', FNotasFiscaisLista[i].FTotalNFe);
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 9] := FormatFloat('0.00', FNotasFiscaisLista[i].FValorMercadorias);
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 10] := FormatFloat('0.00', FNotasFiscaisLista[i].FValorFrete);
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 11] := FormatFloat('0.00', FNotasFiscaisLista[i].FValorBase);
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 12] := FormatFloat('0.00', FNotasFiscaisLista[i].FValorIncentivo);
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 13] := FormatFloat('0.00', FNotasFiscaisLista[i].FValorDeducoes);
      ExcApp.Workbooks[1].Sheets[3].Cells[Linha, 14] := FormatFloat('0.00', FNotasFiscaisLista[i].FValorICMS);

      Linha := Linha + 1;
    end;
  end;
end;

procedure TACBrMapaRecLeite.GeraProdutores;
var
  Linha, i: Integer;
  vRegistro: String;
  NodeProdutoresCadastro, NodeProdutor: IXMLNode;
begin
  if FTipoArquivo = taTXT then
  begin
    if FProdutoresLista.Count > 0 then
      WriteRecord('400', REGISTRO400);
    for i := 0 to FProdutoresLista.Count - 1 do
    begin
      vRegistro := '411|';
      vRegistro := vRegistro + PadLeft(TiraPontos(Trim(FProdutoresLista[i].FProdutorInscEstadual)), 13, '0') + PIPE;
      vRegistro := vRegistro + PadLeft(TiraPontos(Trim(FProdutoresLista[i].FProdutorCPF)), 11, '0') + PIPE;
      vRegistro := vRegistro + Trim(FProdutoresLista[i].FProdutorNome) + PIPE;

      WriteRecord('411', vRegistro);
    end;
    if FProdutoresLista.Count > 0 then
      WriteRecord('499', '499|' + FProdutoresLista.Count.ToString + '|');
  end
  else if FTipoArquivo = taExcel then
  begin
    ExcApp.Workbooks[1].Sheets[1].Name := 'Produtores';

    // Cabeçalho
    ExcApp.Workbooks[1].Sheets[1].Range['A1', 'C' + (FProdutoresLista.Count + 1).ToString].Font.Name := 'Calibri'; // Fonte
    ExcApp.Workbooks[1].Sheets[1].Range['A1', 'C' + (FProdutoresLista.Count + 1).ToString].Font.Size := 11; // Tamanho da Fonte
    ExcApp.Workbooks[1].Sheets[1].Range['A2', 'C' + (FProdutoresLista.Count + 1).ToString].NumberFormat := AnsiChar('@'); // Formato Texto

    ExcApp.Workbooks[1].Sheets[1].Cells[1, 1] := 'CD_PRODUTOR_IE';
    ExcApp.Workbooks[1].Sheets[1].Range['A1', 'A' + (FProdutoresLista.Count + 1).ToString].ColumnWidth := 20;
    ExcApp.Workbooks[1].Sheets[1].Range['A1', 'A' + (FProdutoresLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[1].Range['A1', 'A' + (FProdutoresLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[1].Cells[1, 2] := 'CD_PRODUTOR_CPF';
    ExcApp.Workbooks[1].Sheets[1].Range['B1', 'B' + (FProdutoresLista.Count + 1).ToString].ColumnWidth := 30;
    ExcApp.Workbooks[1].Sheets[1].Range['B1', 'B' + (FProdutoresLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[1].Range['B1', 'B' + (FProdutoresLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    ExcApp.Workbooks[1].Sheets[1].Cells[1, 3] := 'NM_PRODUTOR';
    ExcApp.Workbooks[1].Sheets[1].Range['C1', 'C' + (FProdutoresLista.Count + 1).ToString].ColumnWidth := 80;
    ExcApp.Workbooks[1].Sheets[1].Range['C1', 'C' + (FProdutoresLista.Count + 1).ToString].VerticalAlignment := 2; { VerticalAlignment: 1 = Top, 2 = Center, 3 - Bottom }
    ExcApp.Workbooks[1].Sheets[1].Range['C1', 'C' + (FProdutoresLista.Count + 1).ToString].HorizontalAlignment := 5; { HorizontalAlignment: 3 = Center, 4 = Right, 5 - Left }

    // Dados
    Linha := 2;
    for i := 0 to FProdutoresLista.Count - 1 do
    begin
      ExcApp.Workbooks[1].Sheets[1].Cells[Linha, 1] := PadLeft(TiraPontos(Trim(FProdutoresLista[i].FProdutorInscEstadual)), 13, '0');
      ExcApp.Workbooks[1].Sheets[1].Cells[Linha, 2] := PadLeft(TiraPontos(Trim(FProdutoresLista[i].FProdutorCPF)), 14, '0');
      ExcApp.Workbooks[1].Sheets[1].Cells[Linha, 3] := Trim(FProdutoresLista[i].FProdutorNome);

      Linha := Linha + 1;
    end;
  end
  else if FTipoArquivo = taXML then
  begin
    NodeProdutoresCadastro := NodeRaiz.AddChild('produtoresCadastro');
    for i := 0 to FProdutoresLista.Count - 1 do
    begin
      NodeProdutor := NodeProdutoresCadastro.AddChild('produtor');
      NodeProdutor.ChildValues['cd_produtor_ie'] := PadLeft(TiraPontos(Trim(FProdutoresLista[i].FProdutorInscEstadual)), 13, '0');
      NodeProdutor.ChildValues['cd_produtor_cpf'] := PadLeft(TiraPontos(Trim(FProdutoresLista[i].FProdutorCPF)), 11, '0');
      NodeProdutor.ChildValues['nm_produtor'] := Trim(FProdutoresLista[i].FProdutorNome);
    end;
  end;
end;

function TACBrMapaRecLeite.GerarArquivo: String;
var
  NomeArq: String;
begin
  Result := '';
  FNomeArquivo := '';
  if not DirectoryExists(ExtractFileDir(DirArquivo)) then
    ForceDirectories(DirArquivo);

  if not DirectoryExists(ExtractFileDir(DirArquivo)) then
    raise Exception.Create('Diretório inválido: "' + DirArquivo + '".');

  if FValidarRegistrosAntesGerar and not ValidarRegistros then
    raise Exception.Create('Há rejeições nos registros, corrija para prosseguir.');

  NomeArq := ExtractFileDir(DirArquivo) + PathDelim + GeraNomeArquivo(FModoArquivo, FDeclarante.FInscEstadual, FAnoReferencia, FMesReferencia) + GeraExtensao(FTipoArquivo);

  FDirArquivo := ExtractFileDir(NomeArq);
  FNomeArquivo := ExtractFileName(NomeArq);
  Result := NomeArq;

  if FTipoArquivo = taTXT then
  begin
    AssignFile(Arquivo, NomeArq);
    Rewrite(Arquivo);
    try
      GeraDeclarante;
      GeraRecebimentosLeite;
      GeraNotasFiscais;
      GeraProdutores;
      GeraRegistroFIM;
    finally
      CloseFile(Arquivo);
      LimparRegistros;
    end;
  end
  else if FTipoArquivo = taExcel then
  begin
    ExcApp := CreateOleObject('Excel.Application');
    ExcApp.Visible := False; // Não necessita abrir o excel, apenas salva.
    ExcApp.Caption := FNomeArquivo;
    ExcApp.Workbooks.Add;
    ExcApp.Workbooks[1].Sheets.Add;
    ExcApp.Workbooks[1].Sheets.Add;
    try
      GeraProdutores;
      GeraRecebimentosLeite;
      GeraNotasFiscais;
      ExcApp.Workbooks[1].SaveAs(Result);
    finally
      ExcApp.Quit;
      ExcApp := varEmpty;
      LimparRegistros;
    end;
  end
  else if FTipoArquivo = taXML then
  begin
    XMLDoc := TXMLDocument.Create(Self);
    XMLDoc.Active := True;
    XMLDoc.Version := '1.0';
    XMLDoc.Encoding := 'ISO-8859-1';
    try
      GeraCabecalhoXML;
      GeraDeclarante;
      GeraPeriodoXML;
      GeraProdutores;
      GeraRecebimentosLeite;
      XMLDoc.SaveToFile(Result);
    finally
      XMLDoc.Free;
    end;
  end;
end;

function TACBrMapaRecLeite.GetDataRefFim: TDate;
begin
  Result := IncMonth(EncodeDate(FAnoReferencia, FMesReferencia, 1)) - 1;
end;

function TACBrMapaRecLeite.GetDataRefIni: TDate;
begin
  Result := EncodeDate(FAnoReferencia, FMesReferencia, 1);
end;

procedure TACBrMapaRecLeite.LimparRegistros;
begin
  FDeclarante.FNomeDeclarante := '';
  FDeclarante.FCNPJ := '';
  FDeclarante.FInscEstadual := '';
  FDeclarante.FEmail := '';
  FRejeicoes.Clear;
  FProdutoresLista.Clear;
  FNotasFiscaisLista.Clear;
  FRecebimentosLeiteLista.Clear;
end;

function TACBrMapaRecLeite.ModoArquivoTipoToStr(modArq: TACBrModoArquivo): String;
begin
  if modArq = maProducao then
    Result := 'Produção'
  else if modArq = maHomologacao then
    Result := 'Homologação'
  else
    Result := '';
end;

function TACBrMapaRecLeite.RespFreteTipoToStr(resp: TACBrRespFrete): String;
begin
  if (resp = rfLaticinio) or (resp = rfProdutor) then
  begin
    if (resp = rfLaticinio) then
      Result := 'L'
    else
      Result := 'P';
  end
  else if (FResponsabilidadeFretePadrao = rfLaticinio) or (FResponsabilidadeFretePadrao = rfProdutor) then
  begin
    if (FResponsabilidadeFretePadrao = rfLaticinio) then
      Result := 'L'
    else
      Result := 'P';
  end
  else
    Result := '';
end;

function TACBrMapaRecLeite.RespFreteTipoToStrDesc(resp: TACBrRespFrete): String;
begin
  if resp = rfLaticinio then
    Result := 'Laticínio'
  else if resp = rfProdutor then
    Result := 'Produtor'
  else
    Result := '';
end;

function TACBrMapaRecLeite.TipoArquivoTipoToStr(TipoArq: TACBrTipoArquivo): String;
begin
  if TipoArq = taTXT then
    Result := 'TXT'
  else if TipoArq = taXML then
    Result := 'XML'
  else if TipoArq = taExcel then
    Result := 'Excel'
  else
    Result := '';
end;

function TACBrMapaRecLeite.ValidarRegistros: Boolean;
begin
  if (FModoArquivo <> maProducao) and (FModoArquivo <> maHomologacao) then
    FRejeicoes.Add('Arquivo| Modo Arquivo não informado|');

  if (GetDataRefIni >= Date) or (GetDataRefFim >= Date) then
    FRejeicoes.Add('Arquivo| Período Referência Maior que a Data Atual|');

  if (YearOf(FAnoReferencia) = YearOf(Date)) and (MonthOf(FMesReferencia) = MonthOf(Date)) then
    FRejeicoes.Add('Arquivo| Período Referência igual ao Mês Atual|');

  if FTipoArquivo = taTXT then
    ValidarDeclarante;
  ValidarProdutores;
  ValidarRecebimentosLeite;
  ValidarNotasFiscais;

  Result := (FRejeicoes.Count = 0);
end;

procedure TACBrMapaRecLeite.ValidarDeclarante;
begin
  if FDeclarante.FNomeDeclarante.Trim = '' then
    FRejeicoes.Add('Registro 111| Nome Declarante não informado|');

  if (FDeclarante.FCNPJ.Trim = '') then
    FRejeicoes.Add('Registro 111| CNPJ Declarante não informado|');
  if (FDeclarante.FCNPJ.Trim <> '') and (TiraPontos(FDeclarante.FCNPJ.Trim) = '') or (TiraPontos(FDeclarante.FCNPJ.Trim).Length <> 14) then
    FRejeicoes.Add('Registro 111| CNPJ Declarante informado incorretamente|');

  if FDeclarante.FInscEstadual.Trim = '' then
    FRejeicoes.Add('Registro 111| Inscrição Estadual Declarante não informado|');
  if (UpperCase(FDeclarante.FInscEstadual.Trim) <> 'ISENTO') and ((TiraPontos(FDeclarante.FInscEstadual.Trim) = '') or (StrToFloatDef(TiraPontos(FDeclarante.FInscEstadual.Trim), 1) <= 0)) then
    FRejeicoes.Add('Registro 111| Inscrição Estadual Declarante informado incorretamente|');

  if FDeclarante.FEmail.Trim = '' then
    FRejeicoes.Add('Registro 111| Email Declarante não informado|');
  if (FDeclarante.FEmail.Trim <> '') and (Pos('@', FDeclarante.FEmail.Trim) = 0) then
    FRejeicoes.Add('Registro 111| Email Declarante informado incorretamente|');
end;

procedure TACBrMapaRecLeite.ValidarRecebimentosLeite;
var
  i: Integer;

  function CabecalhoMsg: String;
  begin
    if TipoArquivo = taTXT then
      Result := 'Registro 211'
    else if TipoArquivo = taExcel then
      Result := 'Recebimentos-Leite'
    else if TipoArquivo = taXML then
      Result := 'movimentacaoDiaria';
  end;

begin
  for i := 0 to FRecebimentosLeiteLista.Count - 1 do
  begin
    if FRecebimentosLeiteLista[i].FProdutorInscEstadual.Trim = '' then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Inscrição Estadual Produtor não informado|');
    if (UpperCase(FRecebimentosLeiteLista[i].FProdutorInscEstadual.Trim) <> 'ISENTO') and
      ((TiraPontos(FRecebimentosLeiteLista[i].FProdutorInscEstadual.Trim) = '') or (StrToFloatDef(TiraPontos(FRecebimentosLeiteLista[i].FProdutorInscEstadual.Trim), 1) <= 0)) then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Inscrição Estadual Produtor informado incorretamente|');

    if (FRecebimentosLeiteLista[i].FDataRecebimento < GetDataRefIni) or (DateTimeToDate(FRecebimentosLeiteLista[i].FDataRecebimento) > GetDataRefFim) then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Data Recebimento fora do período referência|');

    if FRecebimentosLeiteLista[i].FQuantLitros <= 0 then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Quantidade de Litros inválida|');

    if FRecebimentosLeiteLista[i].FPlaca.Trim = '' then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Placa não informada|');
  end;
  if FRecebimentosLeiteLista.Count = 0 then
    FRejeicoes.Add(CabecalhoMsg + '| Não há registros de recebimentos de leite informados|');
end;

procedure TACBrMapaRecLeite.ValidarNotasFiscais;
var
  i: Integer;

  function CabecalhoMsg: String;
  begin
    if TipoArquivo = taTXT then
      Result := 'Registro 311'
    else if TipoArquivo = taExcel then
      Result := 'Notas Fiscais - Globais'
    else if TipoArquivo = taXML then
      Result := 'notasFiscais';
  end;

begin
  for i := 0 to FNotasFiscaisLista.Count - 1 do
  begin
    if FNotasFiscaisLista[i].FProdutorInscEstadual.Trim = '' then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Inscrição Estadual Produtor não informado|');
    if (UpperCase(FNotasFiscaisLista[i].FProdutorInscEstadual.Trim) <> 'ISENTO') and ((TiraPontos(FNotasFiscaisLista[i].FProdutorInscEstadual.Trim) = '') or (StrToFloatDef(TiraPontos(FNotasFiscaisLista[i].FProdutorInscEstadual.Trim), 1) <= 0)) then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Inscrição Estadual Produtor informado incorretamente|');

    if (FNotasFiscaisLista[i].FDataNFe < GetDataRefIni) or (DateTimeToDate(FNotasFiscaisLista[i].FDataNFe) > GetDataRefFim) then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Data NFe fora do período referência|');

    if FNotasFiscaisLista[i].FNumero <= 0 then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Número Nota não informado|');

    if FNotasFiscaisLista[i].FSerie.Trim = '' then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Série não informada|');

    if TiraPontos(FNotasFiscaisLista[i].FChave.Trim) = '' then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Chave de Acesso não informada|');
    if (TiraPontos(FNotasFiscaisLista[i].FChave.Trim).Length <> 44) then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Chave de Acesso informada incorretamente|');

    if (FNotasFiscaisLista[i].FResponsabilidadeFrete <> rfLaticinio) and (FNotasFiscaisLista[i].FResponsabilidadeFrete <> rfProdutor) then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Responsabilidade não informada|');

    if FNotasFiscaisLista[i].FQuantLitros <= 0 then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Quantidade de Litros inválida|');

    if FNotasFiscaisLista[i].FTotalNFe <= 0 then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Total NFe inválido|');

    if FNotasFiscaisLista[i].FValorMercadorias <= 0 then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Valor Mercadoria inválida|');

    if FNotasFiscaisLista[i].FValorFrete < 0 then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Frete inválido|');

    if FNotasFiscaisLista[i].FValorBase < 0 then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Valor Base inválido|');

    if FNotasFiscaisLista[i].FValorIncentivo < 0 then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Valor Incentivo inválido|');

    if FNotasFiscaisLista[i].FValorDeducoes < 0 then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Valor Deduções inválido|');

    if FNotasFiscaisLista[i].FValorICMS < 0 then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Valor ICMS inválido|');
  end;
  if FNotasFiscaisLista.Count = 0 then
    FRejeicoes.Add(CabecalhoMsg + '| Não há registros de notas fiscais informados|');
end;

procedure TACBrMapaRecLeite.ValidarProdutores;
var
  i: Integer;
  function CabecalhoMsg: String;
  begin
    if TipoArquivo = taTXT then
      Result := 'Registro 411'
    else if TipoArquivo = taExcel then
      Result := 'Produtores'
    else if TipoArquivo = taXML then
      Result := 'produtoresCadastro';
  end;

begin
  for i := 0 to FProdutoresLista.Count - 1 do
  begin
    if FProdutoresLista[i].FProdutorInscEstadual.Trim = '' then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Inscrição Estadual Produtor não informado|');
    if (UpperCase(FProdutoresLista[i].FProdutorInscEstadual.Trim) <> 'ISENTO') and ((TiraPontos(FProdutoresLista[i].FProdutorInscEstadual.Trim) = '') or (StrToFloatDef(TiraPontos(FProdutoresLista[i].FProdutorInscEstadual.Trim), 1) <= 0)) then
      FRejeicoes.Add(CabecalhoMsg + '| Registro [' + (i + 1).ToString + ']| Inscrição Estadual Produtor informado incorretamente|');

    if (FProdutoresLista[i].FProdutorCPF.Trim = '') then
      FRejeicoes.Add(CabecalhoMsg + '| CPF Produtor não informado|');
    if (TiraPontos(FProdutoresLista[i].FProdutorCPF.Trim) = '') or (TiraPontos(FProdutoresLista[i].FProdutorCPF.Trim).Length <> 11) then
      FRejeicoes.Add(CabecalhoMsg + '| CPF Produtor informado incorretamente|');

    if FProdutoresLista[i].FProdutorNome.Trim = '' then
      FRejeicoes.Add(CabecalhoMsg + '| Nome Produtor não informado|');
  end;
  if FProdutoresLista.Count = 0 then
    FRejeicoes.Add(CabecalhoMsg + '| Não há registros de produtores informados|');
end;

procedure TACBrMapaRecLeite.WriteRecord(Registro, Linha: String);
begin
  write(Arquivo, Linha + #13 + #10);
end;

{ TRecebimentosLeiteLista }

function TRecebimentosLeiteLista.GetObject(Index: Integer): TRecebimentosLeite;
begin
  Result := inherited GetItem(Index) as TRecebimentosLeite;
end;

procedure TRecebimentosLeiteLista.Insert(Index: Integer; Obj: TRecebimentosLeite);
begin
  inherited SetItem(Index, Obj);
end;

function TRecebimentosLeiteLista.New: TRecebimentosLeite;
begin
  Result := TRecebimentosLeite.Create;
  Add(Result);
end;

procedure TRecebimentosLeiteLista.SetObject(Index: Integer; Item: TRecebimentosLeite);
begin
  inherited SetItem(Index, Item);
end;

{ TNotasFiscaisLista }

function TNotasFiscaisLista.GetObject(Index: Integer): TNotasFiscais;
begin
  Result := inherited GetItem(Index) as TNotasFiscais
end;

procedure TNotasFiscaisLista.Insert(Index: Integer; Obj: TNotasFiscais);
begin
  inherited SetItem(Index, Obj);
end;

function TNotasFiscaisLista.New: TNotasFiscais;
begin
  Result := TNotasFiscais.Create;
  Add(Result);
end;

procedure TNotasFiscaisLista.SetObject(Index: Integer; Item: TNotasFiscais);
begin
  inherited SetItem(Index, Item);
end;

{ TProdutoresLista }

function TProdutoresLista.GetObject(Index: Integer): TProdutores;
begin
  Result := inherited GetItem(Index) as TProdutores
end;

procedure TProdutoresLista.Insert(Index: Integer; Obj: TProdutores);
begin
  inherited SetItem(Index, Obj);
end;

function TProdutoresLista.New: TProdutores;
begin
  Result := TProdutores.Create;
  Add(Result);
end;

procedure TProdutoresLista.SetObject(Index: Integer; Item: TProdutores);
begin
  inherited SetItem(Index, Item);
end;

end.
