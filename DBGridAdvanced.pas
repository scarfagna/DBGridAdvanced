unit DBGridAdvanced;


//Stefano Carfagna 
//---------------------------------------------- 
//www.data-ware.it 
//Via Germania 8 - 04016 Sabaudia (LT)
//tel: +39 0773 51 20 01 
//e-mail: scarfagna@data-ware.it 
//----------------------------------------------
//data flows through software
//----------------------------------------------


interface

uses System.Variants, Winapi.Windows, System.SysUtils, Winapi.Messages,
  Vcl.Grids, Vcl.DBGrids, Vcl.Graphics,
  System.Classes, Vcl.Controls, Vcl.Forms, Vcl.StdCtrls,
  Vcl.DBCtrls, Data.DB, Vcl.Menus, Vcl.ImgList;

type

  TItemSearchEvent        = procedure (const Sender: TObject; const column: TColumn; const key: string; var nearestItem : string) of object;

  TDBGridAdvanced = class(TDBGrid)
  private
    { Private declarations }

    // item search event
    fItemSearchFields    : TStringList;              // item search handled
    fOnItemSearch        : TItemSearchEvent;         // manage event

    // autocomplete
    fAutocomplete        : boolean;
    fNextColumnOnReturn  : boolean;
    fAutoappend          : boolean;

    // ...
    inKeepDigit          : boolean;


    FOnColResize         : TNotifyEvent;

  protected
    // autocomplete
    procedure KeyDown (var Key: Word; Shift: TShiftState); override;
    procedure KeyUp   (var Key: Word; Shift: TShiftState); override;
    procedure KeyPress(var Key: Char)                    ; override;
    procedure ColWidthsChanged                           ; override;

    procedure autocompleteKeepDigit(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure autocompleteEndDigit (Sender: TObject; var Key: Word; Shift: TShiftState);

  type
    TInplaceEditListHack = class(TInplaceEditList);
    TDBGridHack          = class(TDBGrid);

  protected
    { Protected declarations }
    procedure ColExit; override;

  public
    { Public declarations }
    // constructor & distructor
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;

    // ...
    property ItemSearchFields    : TStringList      read fItemSearchFields    write fItemSearchFields;

  published
    { Published declarations }

    property Autocomplete        : boolean          read fAutocomplete        write fAutocomplete         default true;
    property NextColumnOnReturn  : boolean          read fNextColumnOnReturn  write fNextColumnOnReturn   default true;
    property Autoappend          : boolean          read fAutoappend          write fAutoappend           default true;

    // ...
    property OnItemSearch        : TItemSearchEvent read fOnItemSearch        write fOnItemSearch;
    property OnColResize         : TNotifyEvent     read FOnColResize         write FOnColResize;
  end;

procedure Register;

implementation  uses System.Math, Vcl.Dialogs;

// Articoli da leggere
// http://www.codeproject.com/Articles/199506/Improving-Delphi-TDBGrid

procedure Register;
begin
  RegisterComponents('Data Controls', [TDBGridAdvanced]);
end;

(* -----------------------------------------------
   | prossima visibile                           |
   ----------------------------------------------- *)

function countInvisible(grid : TDBGrid) : integer;

var
  cnt       : integer;
  invisible : integer;

begin
  invisible := 0;
  for cnt := 0 to grid.columns.Count - 1 do begin
    if not grid.columns[cnt].visible then inc(invisible);
  end;

  result := invisible;
end;

function searchNextVisible(grid : TDBGrid; selected : integer) : integer;

var
  cntInvisible     : integer;
  cntSelezionabili : integer;
  cnt : integer;

begin
  // zero based ...
  result := 0;

  // conta colonne selezionabili = tutte le colonne - quelle invisibili
  cntInvisible     := countInvisible(grid);
  cntSelezionabili := grid.columns.count - cntInvisible;

  // cerca tra le colonne successive alla attuale la prima
  // che risulta visibile ( se c'è )
  for cnt := selected + 1 to cntSelezionabili do begin
    // se visibile : salva id ed esci
    if cnt < grid.columns.count then begin
      if grid.columns[cnt].visible then begin
        result := cnt;
        break;
      end;
    end;
  end;
end;

(* -----------------------------------------------
   | passa alla colonna successiva nella griglia |
   ----------------------------------------------- *)

procedure colonnaSuccessiva(grid : TDBGrid);

var
  datasource   : TDataSource;
  dataset      : TDataSet;

  // ...
  nextVisible  : integer;

begin
  // data source & data set
  datasource := grid.datasource;
  dataset    := datasource.dataset;

  // cerca la prossima colonna a cui passare ...
  nextVisible :=  searchNextVisible(grid, grid.selectedIndex);

  // se è minore della attuale ho finito la riga
  // quindi salvo la riga
  if nextVisible <= grid.selectedIndex  then begin
    // se edit : conferma i dati
    if dataset.state in [dsEdit] then begin
      dataset.post;
    end else
    // se insert : conferma i dati ed inserisce nuova riga
    if dataset.state in [dsInsert] then begin
      dataset.post;
      dataset.append;
    end;
  end;

  // passa alla colonna successiva
  grid.selectedIndex := nextVisible;
end;

(* -----------------------------------------------
   | cerca la prima occorrenza nella pick list uguale o maggiore di  |
   | quella digitata
   ----------------------------------------------- *)

procedure itemsSearch(const items : TStrings; const key: string; var nearestItem : string);

var
  ukey  : string;
  item  : string;
  uitem : string;
  i     : integer;

begin
  ukey               := uppercase(key);
  nearestItem        := '';

  if items.count = 0 then exit;

  for i := 0 to items.count-1 do begin
    item  := items[i];
    uitem := uppercase(item);

    if (uitem >= ukey) then begin
      nearestItem := item;
      break;
    end;
  end;
end;

(* -----------------------------------------------
   | autocompleteOnKeyUp                         |
   ----------------------------------------------- *)

// http://delphidabbler.com/tips/111
// http://delphi.about.com/cs/adptips2003/a/bltip0603_3.htm
// http://marcosalles.wordpress.com/2010/02/27/picklist-do-dbgrid-dropdown/

procedure TDBGridAdvanced.autocompleteKeepDigit(Sender: TObject; var Key: Word;  Shift: TShiftState);

const
  MSG_AVVISO  : string = 'Il controllo del campo è assegnato a FOnItemSearch' + #13#10 + ', ma questa risulta non impostata. ';

var
  s1             : string;
  s2             : string;
  digitato       : string;
  udigitato      : string;

  cutItem        : string;

  editControl    : TInplaceEditListHack;
  pickList       : TStrings;
  currentColumn  : TColumn;

  // result of search
  nearestItem          : string;

  // item search
  isHadledByItemSearch : boolean;
  dummyIndex           : integer;

begin
  // check enabled
  if not enabled then exit;
  // check in place editor
  if TDBGridHack(self).InplaceEditor = nil then exit;

  // check key
  { vk0 thru vk9 are the same as ASCII '0' thru '9' ($30 - $39) }
  { vkA thru vkZ are the same as ASCII 'A' thru 'Z' ($41 - $5A) }
  { 32 spazio  }
  { vkNumpad0 thru vkNumpad9 are the same as ASCII '0' thru '9' (96 - 105) }

  if not ((Key in [$30 .. $39]) or
          (Key in [$41 .. $5A]) or
          (Key in [32, 96 ..105])) then exit;

  // ...
  currentColumn := self.columns[self.selectedIndex];

  // check column editable
  if currentColumn.readOnly                          then exit;

  // check field type
  if (not (currentColumn.field is TWideStringField)) and
     (not (currentColumn.field is TStringField))     then exit;

  // ...
  pickList      := currentColumn.pickList;

  // if no pick list and no event search : exit now
  if (pickList = nil)  and (not Assigned(FOnItemSearch)) then exit;

  // get editor
  editControl := TInplaceEditListHack(TDBGridHack(self).InplaceEditor);

  // check text
  if editControl.text = '' then exit;

  // ...
  s2          := editControl.Text;
  digitato    := s2;

  // check if the search by digit is handled by item search ( external event )
  isHadledByItemSearch := fItemSearchFields.find(currentColumn.FieldName, dummyIndex);

  // default value ...
  nearestItem := '';

  // ...
  if isHadledByItemSearch then begin
    // search for item from item search event or
    if Assigned(FOnItemSearch)    then begin
        // call item search event
        // FOnItemSearch(self, currentColumn, digitato, oisIndex, oisCount, oisFound, nearestItem);
        FOnItemSearch(self, currentColumn, digitato, nearestItem);
    end else begin
        // avvisa che c'è impostazione sbagliata
        messageDlg(MSG_AVVISO + currentColumn.FieldName, mtInformation, [mbOk], 0);
    end;
  end else begin
    // search for item from pick list
    if (pickList <> nil)          then itemsSearch(pickList, digitato, nearestItem);
  end;

  // se sono in una possible sottostringa nella lista
  // dei pick list (non sono andato oltre la pick list)
  if (length(digitato) <= length(nearestItem)) then begin
    s1          := nearestItem;
    cutItem     := uppercase(copy(s1, 1, length(digitato)));
    udigitato   := uppercase(digitato);

    // se la parte iniziale del digitato e dell'item
    // coincidono significa che ho speranza di trovare un item
    if (udigitato = cutItem) then begin
      // copio la parte digitata + altra parte frutto della ricerca
      editControl.text      := copy(digitato, 1, length(s2)) + copy(s1, length(s2) + 1, length(s1) );
      // seleziona la stringa dopo la parte digitata ( quella parte frutto della ricerca )
      editControl.selStart  := length(s2);
      editControl.selLength := length(s1) - length(s2);
    end;
  end else begin
    // non devo fare nulla
    // il controllo gestisce da solo la cosa.
  end;
end;

procedure TDBGridAdvanced.autocompleteEndDigit(Sender: TObject; var Key: Word; Shift: TShiftState);

const
  MSG_AVVISO  : string = 'Il controllo del campo è assegnato a FOnItemSearch' + #13#10 + ', ma questa risulta non impostata. ';

var
  digitato               : string;

  editControl            : TInplaceEditListHack;
  pickList               : TStrings;
  currentColumn          : TColumn;

  // result of search
  nearestItem            : string;

  isHandledBySearchEvent : boolean;
  dummyIndex             : integer;

begin
  // check enabled
  if not enabled then exit;
  // check in place editor
  if TDBGridHack(self).InplaceEditor = nil then exit;
  // check key
  if key <> 13 then exit;

  // prende la colonna corrente ...
  currentColumn := self.columns[self.selectedIndex];

  // check column
  if currentColumn.readOnly                        then exit;
  // check field type
  if (not (currentColumn.field is TWideStringField)) and
     (not (currentColumn.field is TStringField))     then exit;

  // prende pick list della colonna corrente ...
  pickList      := self.columns[self.selectedIndex].pickList;

  // gestito dall evento di ricerca ?
  isHandledBySearchEvent := fItemSearchFields.find(currentColumn.FieldName, dummyIndex);

  // se gestito da search handle ...
  if isHandledBySearchEvent then begin
    // ...
    if not Assigned(FOnItemSearch) then exit;
  end else begin
    // ...
    if (pickList = nil) then exit;
  end;

  try
    // prende il controllo del testo digitato
    editControl := TInplaceEditListHack(TDBGridHack(self).InplaceEditor);
    // prende il digitato ...
    digitato    := editControl.Text;
  except
    exit;
  end;

  // salva il digitato nel controllo
  if self.DataSource.DataSet.State in [dsInsert, dsEdit] then begin
    // chiude drop down
    editControl  .closeUp(true);

    // se gestito da search handle ...
    if isHandledBySearchEvent then begin
      // ...
      if assigned(FOnItemSearch) then begin
        // FOnItemSearch(self, currentColumn, digitato, itemIndex, itemCount, found, nearestItem);
        FOnItemSearch(self, currentColumn, digitato, nearestItem);
      end else begin
        // avvisa che c'è impostazione sbagliata
        messageDlg(MSG_AVVISO + currentColumn.FieldName, mtInformation, [mbOk], 0);
      end;
    end else begin
      // ...
      if (pickList <> nil) then itemsSearch(pickList, digitato, nearestItem);
    end;

    // se risultato coincide con quanto digitato ...
    if uppercase(digitato) = uppercase(nearestItem) then begin
      // usa il valore trovato case sensitive ( non quello digitato )
      digitato := nearestItem;
    end;

    // assegna risultato
    currentColumn.Field.asString := digitato;


    // seleziona tutto
    editControl.selLength        := length(digitato);

    // (testare prima di cancellare impostazione prec )
    // deseleziona tutto siamo a fine digitazione ...
    // editControl.text      := '';
    // editControl.selLength := 0;
  end;
end;

constructor TDBGridAdvanced.Create(AOwner: TComponent);

begin
  inherited Create(AOwner);

  inKeepDigit          := false;

  fAutocomplete        := true;
  fNextColumnOnReturn  := true;
  fAutoappend          := true;
  DefaultRowHeight     := 24;

  fItemSearchFields   := TStringList.Create;
  fItemSearchFields.caseSensitive := false;
end;

destructor TDBGridAdvanced.Destroy;
begin
  fItemSearchFields.free;

  inherited destroy;
end;

procedure TDBGridAdvanced.KeyDown(var Key: Word; Shift: TShiftState);

var
  LookupResultField : TField;

begin
  inherited KeyDown(Key,Shift);

  // mette il valore NULL o clear nel campo
  // relativo alla colonna Lookup
  if Key = VK_DELETE then begin
    // is it a lookup field ?
    if self.selectedField.FieldKind = fkLookup then begin
      // is it in the edit mode ?
      if not (self.DataSource.DataSet.State in [dsInsert, dsEdit]) then begin
        // enter in edit mode
        self.DataSource.DataSet.Edit;
      end;

      //  get field
      LookupResultField := self.DataSource.DataSet.FieldByName (self.SelectedField.KeyFields);
      // set the field to NULL
      LookupResultField.Clear;
    end
  end;
end;

procedure TDBGridAdvanced.KeyUp(var Key: Word; Shift: TShiftState);

begin
  inherited KeyUp(Key,Shift);

  if autocomplete then begin
    if not (key in [VK_F1, VK_F2, VK_DELETE, VK_UP, VK_DOWN, VK_PRIOR, VK_NEXT, VK_ESCAPE]) then begin
      // avoid many call at same time
      if inKeepDigit then exit;
      inKeepDigit := true;
      // ...
      autocompleteKeepDigit(self, Key, Shift);
      // ...
      inKeepDigit := false;
    end;
  end;
end;

procedure TDBGridAdvanced.KeyPress(var Key: Char);

var
 keyCode : word;

begin
  if autocomplete then begin
    // invio ?
    if key = #13 then begin
      keyCode := VK_RETURN;
      // salva il dato prima del cambio della colonna
      autocompleteEndDigit(self, keyCode, []);
    end;
  end;

  // ...
  inherited KeyPress(Key);

  // cambio colonna va eseguito per ultimo
  if fnextcolumnonreturn then begin
    if key = #13 then begin
      // passa alla colonna successiva facendo il "post" se necessario
      colonnaSuccessiva(self);
    end;
  end;
end;

procedure TDBGridAdvanced.ColWidthsChanged;
begin
  inherited;

  // column resize evente
  if Assigned(FOnColResize) then FOnColResize(Self);
end;

procedure TDBGridAdvanced.ColExit;

var
  editControl : TInplaceEditListHack;

begin
  // azzera il valore di in place editor
  // in quanto passando da una colonna all'altra
  // mi proponeva il valore dell'ultima colonna modificata

  // check in place editor
  if TDBGridHack(self).InplaceEditor <> nil then begin
    // get editor
    editControl := TInplaceEditListHack(TDBGridHack(self).InplaceEditor);

    // ...
    editControl.text      := '';
    editControl.selLength := 0;
  end;

  // gestione evento precedente ...
  inherited;
end;


end.
