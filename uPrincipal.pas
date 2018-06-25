unit uPrincipal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.VCLUI.Wait,
  Vcl.Grids, RLFilters, RLXLSFilter, Vcl.StdCtrls, Data.DB, FireDAC.Comp.Client,
  FireDAC.Phys.FB, FireDAC.Phys.FBDef, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.Comp.DataSet, Vcl.DBCtrls, System.Win.ComObj,
  Vcl.DBGrids, Datasnap.DBClient, Vcl.Menus, Vcl.ClipBrd, Vcl.ExtCtrls;

type
  TfPrincipal = class(TForm)
    conexao: TFDConnection;
    edIP: TEdit;
    btnAbrirExcel: TButton;
    RLXLSFilter1: TRLXLSFilter;
    gridXLS: TStringGrid;
    Label1: TLabel;
    AbrirXLS: TOpenDialog;
    lkTabelas: TDBLookupComboBox;
    dsTabelas: TDataSource;
    fdTabelas: TFDQuery;
    fdCampos: TFDQuery;
    dsCampos: TDataSource;
    edDestino: TEdit;
    lstColunas: TListBox;
    cdsVinculos: TClientDataSet;
    dsVinculos: TDataSource;
    lstCampos: TListBox;
    DBGrid1: TDBGrid;
    cdsVinculosCOLUNA: TStringField;
    cdsVinculosCAMPO: TStringField;
    btnGerar: TButton;
    rbInsert: TRadioButton;
    rbUpdate: TRadioButton;
    memScript: TMemo;
    ckShowScpt: TCheckBox;
    cdsVinculosCOL: TIntegerField;
    fdCamposCAMPO: TWideStringField;
    fdCamposTIPO: TSmallintField;
    fdCamposTTIPO: TStringField;
    cdsVinculosTIPO: TIntegerField;
    popCopiar: TPopupMenu;
    Copiar1: TMenuItem;
    Label2: TLabel;
    fdGenerator: TFDQuery;
    dsGenerator: TDataSource;
    popGen: TPopupMenu;
    MenuItem1: TMenuItem;
    lkGen: TDBLookupComboBox;
    fdGeneratorRDBGENERATOR_NAME: TWideStringField;
    fdGeneratorRDBGENERATOR_ID: TSmallintField;
    fdGeneratorRDBSYSTEM_FLAG: TSmallintField;
    fdGeneratorRDBDESCRIPTION: TMemoField;
    ListBox1: TListBox;
    btnOk: TButton;
    RadioGroup1: TRadioGroup;
    Chave1: TMenuItem;
    Label3: TLabel;
    lbChave: TLabel;
    lbCampoWhere: TLabel;
    Button1: TButton;
    procedure btnAbrirExcelClick(Sender: TObject);
    procedure lkTabelasClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure lkTabelasEnter(Sender: TObject);
    procedure enterDB(Sender: TObject);
    procedure exitDB(Sender: TObject);
    procedure ckShowScptClick(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure lstCamposDblClick(Sender: TObject);
    procedure fdCamposCalcFields(DataSet: TDataSet);
    procedure Copiar1Click(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure RadioGroup1Click(Sender: TObject);
    procedure Chave1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    str : string;
    Chave_i : integer;
    function XlsToStringGrid(XStringGrid: TStringGrid;  xFileXLS: string): Boolean;
    function TrocaCaracteres(Str, antes, depois: string): string;
    procedure ConfiguraConex;
    procedure PegarCampos;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fPrincipal: TfPrincipal;

implementation

{$R *.dfm}

function TfPrincipal.TrocaCaracteres(Str, antes, depois: string): string;
var
  s: string;
  x: Integer;
begin
  for x := 1 to Length(Str) do
    if Str[X] = antes then
      S := S + depois
    else
      S := S + Str[X];

  Result := S;
end;

procedure TfPrincipal.exitDB(Sender: TObject);
begin
  if not ((edIP.Text = '') and (edDestino.Text = '')) then
    ConfiguraConex;
end;

procedure TfPrincipal.fdCamposCalcFields(DataSet: TDataSet);
begin
  case fdCamposTIPO.Value of
    8 : fdCamposTTIPO.Value := 'Integer';
    35 : fdCamposTTIPO.Value := 'String';
  else
    fdCamposTTIPO.Value := 'Indefinido';
  end
end;

procedure TfPrincipal.btnGerarClick(Sender: TObject);
var
  SQL : String;
  i : integer;
begin
  memScript.Clear;

  if rbInsert.Checked then
    begin
      for i := 0 to gridXLS.RowCount-1 do
        begin
          SQL := 'insert into ' + lkTabelas.Text + '(';

          cdsVinculos.First;

          while not cdsVinculos.Eof do
            begin
              if (cdsVinculos.RecNo = 1) then
                SQL := SQL + cdsVinculosCAMPO.AsString
              else
                SQL := SQL + ', ' + cdsVinculosCAMPO.AsString;

              cdsVinculos.Next;
            end;

          SQL := SQL + ') values (';

          cdsVinculos.First;

          while not cdsVinculos.Eof do
            begin
              if not (cdsVinculos.RecNo = 1) then
                SQL := SQL + ', ';

                case cdsVinculosTIPO.Value of
                  -1 : SQL := SQL + 'gen_id('+lkGen.Text+',1)';
                  7  : SQL := SQL + gridXLS.Cells[cdsVinculosCOL.Value,i];
                  8  : SQL := SQL + gridXLS.Cells[cdsVinculosCOL.Value,i];
                  12 : SQL := SQL + 'cast('+''''+FormatDateTime('mm-dd-yyyy', StrToDate(gridXLS.Cells[cdsVinculosCOL.Value,i]))+''''+' as date)';
                  13 : SQL := SQL + 'cast('+''''+gridXLS.Cells[cdsVinculosCOL.Value,i]+''''+' as time)';
                  27 : SQL := SQL + TrocaCaracteres(gridXLS.Cells[cdsVinculosCOL.Value,i], ',', '.');
                  37 : SQL := SQL + ''''+gridXLS.Cells[cdsVinculosCOL.Value,i]+'''';
                end;

              cdsVinculos.Next;
            end;
          SQL := SQL + ');';

          memScript.Lines.Add(SQL);
        end;
    end
  else
    begin
      for i := 0 to gridXLS.RowCount-1 do
        begin
          SQL := 'update ' + lkTabelas.Text + ' set ';

          cdsVinculos.First;

          while not cdsVinculos.Eof do
            begin
              if not (cdsVinculos.RecNo = 1) then
                SQL := SQL + ', ';

                SQL := SQL + cdsVinculosCAMPO.AsString + ' = ';

                case cdsVinculosTIPO.Value of
                  -1 : SQL := SQL + 'gen_id('+lkGen.Text+',1)';
                  7  : SQL := SQL + gridXLS.Cells[cdsVinculosCOL.Value,i];
                  8  : SQL := SQL + gridXLS.Cells[cdsVinculosCOL.Value,i];
                  12 : SQL := SQL + 'cast('+''''+FormatDateTime('mm-dd-yyyy', StrToDate(gridXLS.Cells[cdsVinculosCOL.Value,i]))+''''+' as date)';
                  13 : SQL := SQL + 'cast('+''''+gridXLS.Cells[cdsVinculosCOL.Value,i]+''''+' as time)';
                  27 : SQL := SQL + TrocaCaracteres(gridXLS.Cells[cdsVinculosCOL.Value,i], ',', '.');
                  37 : SQL := SQL + ''''+gridXLS.Cells[cdsVinculosCOL.Value,i]+'''';
                end;

              cdsVinculos.Next;
            end;

          SQL := SQL + ' where ' + lbCampoWhere.Caption + ' = ';

          cdsVinculos.Locate('COL', Chave_i,[]);

          case cdsVinculosTIPO.Value of
            -1 : SQL := SQL + 'gen_id('+lkGen.Text+',1)';
            7  : SQL := SQL + gridXLS.Cells[cdsVinculosCOL.Value,i];
            8  : SQL := SQL + gridXLS.Cells[cdsVinculosCOL.Value,i];
            12 : SQL := SQL + 'cast('+''''+FormatDateTime('mm-dd-yyyy', StrToDate(gridXLS.Cells[cdsVinculosCOL.Value,i]))+''''+' as date)';
            13 : SQL := SQL + 'cast('+''''+gridXLS.Cells[cdsVinculosCOL.Value,i]+''''+' as time)';
            27 : SQL := SQL + TrocaCaracteres(gridXLS.Cells[cdsVinculosCOL.Value,i], ',', '.');
            37 : SQL := SQL + ''''+gridXLS.Cells[cdsVinculosCOL.Value,i]+'''';
          end;

          SQL := SQL + ';';

          memScript.Lines.Add(SQL);
        end;
    end;

  memScript.Lines.Add('');
  memScript.Lines.Add('COMMIT WORK;');

  ckShowScpt.Checked := True;
end;

procedure TfPrincipal.btnOkClick(Sender: TObject);
begin
  cdsVinculos.Append;
  cdsVinculosCOLUNA.Value := lstColunas.Items[lstColunas.ItemIndex];
  cdsVinculosCOL.Value := lstColunas.ItemIndex;
  cdsVinculosCAMPO.Value := lstCampos.Items[lstCampos.ItemIndex];;
  cdsVinculosTIPO.Value := -1;
  cdsVinculos.Post;

  lkGen.Visible := false;
  btnOk.Visible := false;
end;

procedure TfPrincipal.Button1Click(Sender: TObject);
var
  SQL : String;
  i : integer;
  valor : string;
begin
  memScript.Clear;

  for i := 0 to gridXLS.RowCount-1 do
    begin
      SQL := 'INSERT INTO TABELA (ID, DESCRICAO) VALUES ( ';
      SQL := SQL + 'gen_id(GEN_id,1),'+ gridXLS.Cells[0,i]+')';

      memScript.Lines.Add(SQL);
      memScript.Lines.Add('');
    end;

  memScript.Lines.Add('');
  memScript.Lines.Add('COMMIT WORK;');

  ckShowScpt.Checked := True;
end;

procedure TfPrincipal.Chave1Click(Sender: TObject);
begin
  lbChave.Caption := lstColunas.Items[lstColunas.ItemIndex];
  lbCampoWhere.Caption := lstCampos.Items[lstCampos.ItemIndex];
  Chave_i := lstCampos.ItemIndex;
end;

procedure TfPrincipal.ckShowScptClick(Sender: TObject);
begin
  memScript.Visible := ckShowScpt.Checked;

  if ckShowScpt.Checked then memScript.BringToFront;
end;

procedure TfPrincipal.ConfiguraConex;
begin
  try
    conexao.Connected := False;
    conexao.Params.Clear;
    conexao.Params.Add('User_Name=user');
    conexao.Params.Add('Password=pass');
    conexao.Params.Add('Database='+edDestino.Text);
    conexao.Params.Add('Server='+edIP.Text);
    conexao.Params.Add('DriverID=FB');
    conexao.Connected := True;

    lkTabelas.Enabled := True;
  except
    on E : Exception do
      begin
        ShowMessage(E.Message);
      end;
  end;
end;

procedure TfPrincipal.Copiar1Click(Sender: TObject);
begin
  ClipBoard.AsText := memScript.Text;
end;

procedure TfPrincipal.enterDB(Sender: TObject);
begin
  str := TEdit(Sender).Text;
end;

procedure TfPrincipal.FormCreate(Sender: TObject);
begin
  cdsVinculos.CreateDataSet;
end;

procedure TfPrincipal.lkTabelasClick(Sender: TObject);
begin
  fdCampos.Close;

//  fdCampos.SQL.Clear;

{  if RadioGroup1.ItemIndex = 0 then
    fdCampos.SQL.Add('select rdb$relation_name as tabelas from rdb$relations where rdb$system_flag = 0 and rdb$relation_name = :tbl')
  else
    fdCampos.SQL.Add('SELECT A.RDB$PROCEDURE_NAME FROM RDB$PROCEDURES A ORDER BY A.RDB$PROCEDURE_NAME');}

  fdCampos.Params[0].Value := lkTabelas.Text;
  fdCampos.Open();

  PegarCampos;
end;

procedure TfPrincipal.lkTabelasEnter(Sender: TObject);
begin
  if not fdTabelas.Active then fdTabelas.Active := True;
end;

procedure TfPrincipal.lstCamposDblClick(Sender: TObject);
begin
  cdsVinculos.Append;
  cdsVinculosCOLUNA.Value := lstColunas.Items[lstColunas.ItemIndex];
  cdsVinculosCOL.Value := lstColunas.ItemIndex;
  cdsVinculosCAMPO.Value := lstCampos.Items[lstCampos.ItemIndex];
  cdsVinculosTIPO.Value := ListBox1.Items[lstCampos.ItemIndex].ToInteger;
  cdsVinculos.Post;
end;

procedure TfPrincipal.MenuItem1Click(Sender: TObject);
begin
  fdGenerator.Close;
  fdGenerator.Open();

  lkGen.Visible := True;
  btnOk.Visible := True;
end;

procedure TfPrincipal.PegarCampos;
begin
  fdCampos.First;

  while not fdCampos.Eof do
    begin
      lstCampos.Items.Add(fdCamposCAMPO.Value);
      ListBox1.Items.Add(fdCamposTIPO.Value.ToString);

      fdCampos.Next;
    end;
end;

procedure TfPrincipal.RadioGroup1Click(Sender: TObject);
begin
  lkTabelasClick(Sender);
end;

function TfPrincipal.XlsToStringGrid(XStringGrid: TStringGrid;
  xFileXLS: string): Boolean;
const
   xlCellTypeLastCell = $0000000B;
var
   XLSAplicacao, AbaXLS: OLEVariant;
   RangeMatrix: Variant;
   x, y, k, r: Integer;
begin
   Result := False;
   // Cria Excel- OLE Object
   XLSAplicacao := CreateOleObject('Excel.Application');
   try
   // Esconde Excel
      XLSAplicacao.Visible := False;
      // Abre o Workbook
      XLSAplicacao.Workbooks.Open(xFileXLS);

      {Selecione aqui a aba que você deseja abrir primeiro - 1,2,3,4....}
      XLSAplicacao.WorkSheets[1].Activate;
      {Selecione aqui a aba que você deseja ativar - começando sempre no 1 (1,2,3,4) }
      AbaXLS := XLSAplicacao.Workbooks[ExtractFileName(xFileXLS)].WorkSheets[1];

      AbaXLS.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
      // Pegar o número da última linha
      x := XLSAplicacao.ActiveCell.Row;
      // Pegar o número da última coluna
      y := XLSAplicacao.ActiveCell.Column;
      // Seta xStringGrid linha e coluna
      XStringGrid.RowCount := x;
      XStringGrid.ColCount := y;
      // Associaca a variant WorkSheet com a variant do Delphi
      RangeMatrix := XLSAplicacao.Range['A1', XLSAplicacao.Cells.Item[x, y]].Value;
      // Cria o loop para listar os registros no TStringGrid
      k := 1;
      repeat
         for r := 1 to y do
            XStringGrid.Cells[(r - 1), (k - 1)] := RangeMatrix[k, r];
         Inc(k, 1);
      until k > x;
      RangeMatrix := Unassigned;
      finally
            // Fecha o Microsoft Excel
            if not VarIsEmpty(XLSAplicacao) then
            begin
                  XLSAplicacao.Quit;
                  XLSAplicacao := Unassigned;
                  AbaXLS := Unassigned;
                  Result := True;
            end;
      end;
end;

procedure TfPrincipal.btnAbrirExcelClick(Sender: TObject);
var
  I : integer;
begin
  if AbrirXLS.Execute then XlsToStringGrid(gridXLS, AbrirXLS.FileName);

  for I := 0 to gridXLS.ColCount-1 do
    lstColunas.Items.Add( 'Coluna ' + IntToStr(I+1));
end;

end.
