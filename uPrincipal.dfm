object fPrincipal: TfPrincipal
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  ClientHeight = 521
  ClientWidth = 769
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Verdana'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  DesignSize = (
    769
    521)
  PixelsPerInch = 96
  TextHeight = 14
  object Label1: TLabel
    Left = 8
    Top = 8
    Width = 154
    Height = 14
    Caption = 'Destino Banco de Dados'
  end
  object Label2: TLabel
    Left = 272
    Top = 8
    Width = 369
    Height = 14
    Alignment = taRightJustify
    Caption = 'Os campos do arquivo XLS devem estar em formato Texto!'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -12
    Font.Name = 'Verdana'
    Font.Style = []
    ParentFont = False
  end
  object Label3: TLabel
    Left = 406
    Top = 468
    Width = 93
    Height = 14
    Caption = 'Campo Chave:'
  end
  object lbChave: TLabel
    Left = 505
    Top = 468
    Width = 4
    Height = 14
  end
  object lbCampoWhere: TLabel
    Left = 406
    Top = 492
    Width = 4
    Height = 14
  end
  object ListBox1: TListBox
    Left = 584
    Top = 152
    Width = 57
    Height = 277
    ItemHeight = 14
    TabOrder = 14
    Visible = False
  end
  object memScript: TMemo
    Left = 8
    Top = 152
    Width = 1073
    Height = 310
    Lines.Strings = (
      'memScript')
    PopupMenu = popCopiar
    TabOrder = 11
    Visible = False
  end
  object edIP: TEdit
    Left = 8
    Top = 28
    Width = 177
    Height = 22
    TabOrder = 0
    Text = 'localhost'
    OnEnter = enterDB
    OnExit = exitDB
  end
  object btnAbrirExcel: TButton
    Left = 8
    Top = 56
    Width = 73
    Height = 90
    Caption = 'Abrir arquivo .XLS'
    TabOrder = 1
    WordWrap = True
    OnClick = btnAbrirExcelClick
  end
  object gridXLS: TStringGrid
    Left = 87
    Top = 56
    Width = 674
    Height = 90
    Anchors = [akLeft, akTop, akRight]
    ColCount = 1
    DefaultRowHeight = 20
    FixedCols = 0
    RowCount = 1
    FixedRows = 0
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Verdana'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    ColWidths = (
      64)
    RowHeights = (
      20)
  end
  object lkTabelas: TDBLookupComboBox
    Left = 471
    Top = 152
    Width = 290
    Height = 22
    Anchors = [akLeft, akTop, akRight]
    Enabled = False
    KeyField = 'TABELAS'
    ListField = 'TABELAS'
    ListSource = dsTabelas
    TabOrder = 3
    OnClick = lkTabelasClick
    OnEnter = lkTabelasEnter
  end
  object edDestino: TEdit
    Left = 191
    Top = 28
    Width = 570
    Height = 22
    Anchors = [akLeft, akTop, akRight]
    TabOrder = 4
    Text = 'D:\User\Will\BRFIN.FDB'
    OnEnter = enterDB
    OnExit = exitDB
  end
  object lstColunas: TListBox
    Left = 8
    Top = 152
    Width = 150
    Height = 310
    ItemHeight = 14
    PopupMenu = popGen
    TabOrder = 5
  end
  object lstCampos: TListBox
    Left = 164
    Top = 152
    Width = 150
    Height = 310
    ItemHeight = 14
    TabOrder = 6
    OnDblClick = lstCamposDblClick
  end
  object DBGrid1: TDBGrid
    Left = 320
    Top = 227
    Width = 441
    Height = 235
    DataSource = dsVinculos
    Options = [dgTitles, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit, dgTitleHotTrack]
    TabOrder = 7
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -12
    TitleFont.Name = 'Verdana'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'COLUNA'
        Width = 101
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CAMPO'
        Width = 100
        Visible = True
      end>
  end
  object btnGerar: TButton
    Left = 535
    Top = 488
    Width = 226
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = 'Gerar Script'
    TabOrder = 8
    OnClick = btnGerarClick
  end
  object rbInsert: TRadioButton
    Left = 8
    Top = 496
    Width = 73
    Height = 17
    Anchors = [akLeft, akBottom]
    Caption = 'Insert'
    Checked = True
    TabOrder = 9
    TabStop = True
  end
  object rbUpdate: TRadioButton
    Left = 87
    Top = 496
    Width = 75
    Height = 17
    Anchors = [akLeft, akBottom]
    Caption = 'Update'
    TabOrder = 10
  end
  object ckShowScpt: TCheckBox
    Left = 191
    Top = 496
    Width = 138
    Height = 17
    Anchors = [akLeft, akBottom]
    Caption = 'Mostrar Script'
    TabOrder = 12
    OnClick = ckShowScptClick
  end
  object lkGen: TDBLookupComboBox
    Left = 8
    Top = 468
    Width = 361
    Height = 22
    Anchors = [akLeft, akBottom]
    KeyField = 'RDB$GENERATOR_NAME'
    ListField = 'RDB$GENERATOR_NAME'
    ListSource = dsGenerator
    TabOrder = 13
    Visible = False
  end
  object btnOk: TButton
    Left = 375
    Top = 468
    Width = 25
    Height = 22
    Anchors = [akLeft, akBottom]
    Caption = 'ok'
    TabOrder = 15
    Visible = False
    OnClick = btnOkClick
  end
  object RadioGroup1: TRadioGroup
    Left = 320
    Top = 152
    Width = 145
    Height = 69
    Caption = ' Tipo de Importa'#231#227'o'
    Color = clWhite
    ItemIndex = 0
    Items.Strings = (
      'Tabelas'
      'Procedures')
    ParentBackground = False
    ParentColor = False
    TabOrder = 16
    OnClick = RadioGroup1Click
  end
  object Button1: TButton
    Left = 406
    Top = 488
    Width = 123
    Height = 25
    Caption = 'Button1'
    TabOrder = 17
    OnClick = Button1Click
  end
  object conexao: TFDConnection
    Params.Strings = (
      'User_Name=sysdba'
      'Password=masterkey'
      'DriverID=FB')
    LoginPrompt = False
    Left = 32
    Top = 312
  end
  object RLXLSFilter1: TRLXLSFilter
    FileName = 'D:\User\Will\feriados_nacionais.xls'
    DisplayName = 'Planilha Excel 97-2013'
    Left = 352
    Top = 80
  end
  object AbrirXLS: TOpenDialog
    DefaultExt = '*.xls'
    FileName = 'D:\User\Will\feriados_nacionais.xls'
    Filter = 'Arquivos Excel|*.xls'
    Left = 416
    Top = 312
  end
  object dsTabelas: TDataSource
    DataSet = fdTabelas
    Left = 56
    Top = 360
  end
  object fdTabelas: TFDQuery
    Connection = conexao
    SQL.Strings = (
      'select rdb$relation_name as tabelas from rdb$relations'
      'where rdb$system_flag = 0')
    Left = 24
    Top = 360
  end
  object fdCampos: TFDQuery
    OnCalcFields = fdCamposCalcFields
    Connection = conexao
    SQL.Strings = (
      'select '
      'A.RDB$FIELD_NAME AS campo,'
      'B.RDB$FIELD_TYPE as tipo'
      'from rdb$relation_fields a'
      
        'inner join rdb$fields b on (a.rdb$field_source = b.rdb$field_nam' +
        'e)'
      'where rdb$relation_name = :tabela')
    Left = 16
    Top = 456
    ParamData = <
      item
        Name = 'TABELA'
        DataType = ftFixedWideChar
        ParamType = ptInput
        Size = 31
      end>
    object fdCamposCAMPO: TWideStringField
      FieldName = 'CAMPO'
      Origin = 'RDB$FIELD_NAME'
      FixedChar = True
      Size = 31
    end
    object fdCamposTIPO: TSmallintField
      AutoGenerateValue = arDefault
      FieldName = 'TIPO'
      Origin = 'RDB$FIELD_TYPE'
      ProviderFlags = []
      ReadOnly = True
    end
    object fdCamposTTIPO: TStringField
      FieldKind = fkCalculated
      FieldName = 'TTIPO'
      Size = 30
      Calculated = True
    end
  end
  object dsCampos: TDataSource
    DataSet = fdCampos
    Left = 48
    Top = 456
  end
  object cdsVinculos: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'COLUNA'
        DataType = ftString
        Size = 40
      end
      item
        Name = 'CAMPO'
        DataType = ftString
        Size = 40
      end
      item
        Name = 'COL'
        DataType = ftInteger
      end
      item
        Name = 'TIPO'
        DataType = ftInteger
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 24
    Top = 272
    object cdsVinculosCOLUNA: TStringField
      FieldName = 'COLUNA'
      Size = 40
    end
    object cdsVinculosCAMPO: TStringField
      FieldName = 'CAMPO'
      Size = 40
    end
    object cdsVinculosCOL: TIntegerField
      FieldName = 'COL'
    end
    object cdsVinculosTIPO: TIntegerField
      FieldName = 'TIPO'
    end
  end
  object dsVinculos: TDataSource
    DataSet = cdsVinculos
    Left = 56
    Top = 272
  end
  object popCopiar: TPopupMenu
    Left = 80
    Top = 200
    object Copiar1: TMenuItem
      Caption = 'Copiar...'
      OnClick = Copiar1Click
    end
  end
  object fdGenerator: TFDQuery
    Connection = conexao
    SQL.Strings = (
      'select * from RDB$GENERATORS'
      'WHERE RDB$SYSTEM_FLAG = 0')
    Left = 24
    Top = 408
    object fdGeneratorRDBGENERATOR_NAME: TWideStringField
      FieldName = 'RDB$GENERATOR_NAME'
      Origin = 'RDB$GENERATOR_NAME'
      FixedChar = True
      Size = 31
    end
    object fdGeneratorRDBGENERATOR_ID: TSmallintField
      FieldName = 'RDB$GENERATOR_ID'
      Origin = 'RDB$GENERATOR_ID'
    end
    object fdGeneratorRDBSYSTEM_FLAG: TSmallintField
      FieldName = 'RDB$SYSTEM_FLAG'
      Origin = 'RDB$SYSTEM_FLAG'
    end
    object fdGeneratorRDBDESCRIPTION: TMemoField
      FieldName = 'RDB$DESCRIPTION'
      Origin = 'RDB$DESCRIPTION'
      BlobType = ftMemo
    end
  end
  object dsGenerator: TDataSource
    DataSet = fdGenerator
    Left = 56
    Top = 408
  end
  object popGen: TPopupMenu
    Left = 48
    Top = 200
    object MenuItem1: TMenuItem
      Caption = 'Campo AutoInc...'
      OnClick = MenuItem1Click
    end
    object Chave1: TMenuItem
      Caption = 'Chave...'
      OnClick = Chave1Click
    end
  end
end
