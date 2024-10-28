unit Unit3;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, StdCtrls, Grids,
  Dialogs, ExtCtrls, Eingabe;

type

  { TFrame2 }

  TFrame2 = class(TFrame)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    Edit1: TEdit;
    Edit10: TEdit;
    Edit11: TEdit;
    Edit12: TEdit;
    Edit13: TEdit;
    Edit14: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit7: TEdit;
    Edit8: TEdit;
    Edit9: TEdit;
    Label1: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label2: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    StringGrid1: TStringGrid;
    Timer1: TTimer;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);
    procedure Edit10KeyPress(Sender: TObject; var Key: char);
    procedure Edit10KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit11KeyPress(Sender: TObject; var Key: char);
    procedure Edit11KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit12KeyPress(Sender: TObject; var Key: char);
    procedure Edit12KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit13KeyPress(Sender: TObject; var Key: char);
    procedure Edit13KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit14KeyPress(Sender: TObject; var Key: char);
    procedure Edit14KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit1KeyPress(Sender: TObject; var Key: char);
    procedure Edit1KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit2KeyPress(Sender: TObject; var Key: char);
    procedure Edit2KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit3KeyPress(Sender: TObject; var Key: char);
    procedure Edit3KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit4KeyPress(Sender: TObject; var Key: char);
    procedure Edit4KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit5KeyPress(Sender: TObject; var Key: char);
    procedure Edit5KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit6KeyPress(Sender: TObject; var Key: char);
    procedure Edit6KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit7KeyPress(Sender: TObject; var Key: char);
    procedure Edit7KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit8KeyPress(Sender: TObject; var Key: char);
    procedure Edit8KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit9KeyPress(Sender: TObject; var Key: char);
    procedure Edit9KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure StringGrid1SelectCell(Sender: TObject; aCol, aRow: Integer;
      var CanSelect: Boolean);
    procedure Timer1Timer(Sender: TObject);
  private
    AOwner: TComponent;
  public
    constructor Create(AOwner1: TComponent); override;                          // Frame initialisieren (Variable und Komponenteneigenschaften)
    destructor Destroy; override;                                               // Frame schliessen
    procedure FBLoe;                                                            // Löschen Eingabefelder Bildschirm
    procedure Ladenpersonen;                                                    // Laden Daten Personen aus Tabelle in Stringgrid
  end;

type Tfperson = record                                                           // Satzaufbau Personen
  Nummer: integer;
  Art: string;                                                                  // Person/Firma
  Firmenname: string;
  Vorname: string;
  Nachname: string;
  Strasse: string;
  Hausnummer: string;
  Plz: string;
  Ort: string;
  Land: string;
  Telefon: string;
  Handy: string;
  Email: string;
  Partneranrede: string;
  Partnervorname: string;
  Partnernachname: string;
  Partnertelefon: string;
  Satzart: Boolean;                                                             // aktiv/passiv
end;

var
  ftasts: Boolean;                                                              // Darf ein Button gedrückt werden
  fart: integer;                                                                // Verarbeitungsart 1=Insert 2=Update 3=Delete
  flauf: integer;                                                               // Welche Eingabe ist zulässig 0=keine Eingabe
  fgrid: integer;                                                               // Nummer der angeklickten Zeile im Stringgrid
  fperson: Tfperson;

implementation

{$R *.lfm}

uses Unit1;

{ TFrame2 }

destructor TFrame2.Destroy;
begin
  AOwner.Free;
  inherited Destroy;
end;

procedure TFrame2.Button4Click(Sender: TObject);
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  if ftasts then
  begin
    Label2.Caption:='';
    ftasts:=False;
    fart:=0;
    flauf:=0;
    Form1.Frame2Close;
  end;
end;

procedure TFrame2.ComboBox1Change(Sender: TObject);
  var h1: string;
begin
  if flauf = 1 then
  begin
    Label2.Caption:='';
    h1:=ComboBox1.Text;
    ialpha:=Form1.Blankweg(3, h1);
    if ialpha <> '' then
    begin
      fperson.Art:=ialpha;
      ComboBox1.Enabled:=False;
      ComboBox1.TabStop:=False;
      ComboBox1.ReadOnly:=True;
      ComboBox1.Color:=$00C0C0C0;
      if ialpha = 'Person' then
      begin
        flauf:=3;
        Edit1.Enabled:=True;
        Edit1.ReadOnly:=False;
        Edit1.TabStop:=True;
        Edit1.Color:=$0080FFFF;
        Edit1.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=25;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit1.Text:=fperson.Vorname;
          ialpha:=fperson.Vorname;
          iart:=1;
        end;
      end else begin
        flauf:=2;
        Edit9.Enabled:=True;
        Edit9.ReadOnly:=False;
        Edit9.TabStop:=True;
        Edit9.Color:=$0080FFFF;
        Edit9.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=30;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit9.Text:=fperson.Firmenname;
          ialpha:=fperson.Firmenname;
          iart:=1;
        end;
      end;
    end else begin
      Label2.Caption:='Personenart muss ausgewählt werden';
    end;
  end;
end;

procedure TFrame2.ComboBox2Change(Sender: TObject);
  var h1: string;
  var h2: integer;
begin
  if flauf = 13 then
  begin
    Label2.Caption:='';
    h1:=ComboBox2.Text;
    ialpha:=Form1.Blankweg(3, h1);
    if ialpha <> '' then
    begin
      fperson.Partneranrede:=ialpha;
      ComboBox2.Enabled:=True;
      ComboBox2.ReadOnly:=False;
      ComboBox2.TabStop:=True;
      ComboBox2.Color:=$00C0C0C0;
      if ialpha = 'ohne' then
      begin
        if fart = 1 then
        begin
          JaNein:=messagedlg('Person/Firma anlegen?', mtConfirmation, [mbYes, mbNo], 0);
          if (JaNein = mrYes) then
          begin
            h2:=fperson.Nummer;
            Persontab[h2].Nummer:=h2;
            Persontab[h2].Art:=fperson.Art;
            Persontab[h2].Land:=fperson.Land;
            Persontab[h2].Email:=fperson.Email;
            Persontab[h2].Handy:=fperson.Handy;
            Persontab[h2].Telefon:=fperson.Telefon;
            Persontab[h2].Ort:=fperson.Ort;
            Persontab[h2].Plz:=fperson.Plz;
            Persontab[h2].Hausnummer:=fperson.Hausnummer;
            Persontab[h2].Strasse:=fperson.Strasse;
            Persontab[h2].Nachname:=fperson.Nachname;
            Persontab[h2].Firmenname:=fperson.Firmenname;
            Persontab[h2].Vorname:=fperson.Vorname;
            Persontab[h2].Partnertelefon:=fperson.Partnertelefon;
            Persontab[h2].Partnernachname:=fperson.Partnernachname;
            Persontab[h2].Partnervorname:=fperson.Partnervorname;
            Persontab[h2].Partneranrede:=fperson.Partneranrede;
            Persontab[h2].Satzart:=True;
            Personzaehl:=Personzaehl+1;
            Updperson:=True;
            Ladenpersonen;
          end;
        end else begin
          JaNein:=messagedlg('Person/Firma ändern?', mtConfirmation, [mbYes, mbNo], 0);
          if (JaNein = mrYes) then
          begin
            h2:=fperson.Nummer;
            Persontab[h2].Nummer:=h2;
            Persontab[h2].Art:=fperson.Art;
            Persontab[h2].Land:=fperson.Land;
            Persontab[h2].Email:=fperson.Email;
            Persontab[h2].Handy:=fperson.Handy;
            Persontab[h2].Telefon:=fperson.Telefon;
            Persontab[h2].Ort:=fperson.Ort;
            Persontab[h2].Plz:=fperson.Plz;
            Persontab[h2].Hausnummer:=fperson.Hausnummer;
            Persontab[h2].Strasse:=fperson.Strasse;
            Persontab[h2].Nachname:=fperson.Nachname;
            Persontab[h2].Firmenname:=fperson.Firmenname;
            Persontab[h2].Vorname:=fperson.Vorname;
            Persontab[h2].Partnertelefon:=fperson.Partnertelefon;
            Persontab[h2].Partnernachname:=fperson.Partnernachname;
            Persontab[h2].Partnervorname:=fperson.Partnervorname;
            Persontab[h2].Partneranrede:=fperson.Partneranrede;
            Persontab[h2].Satzart:=True;
            Updperson:=True;
            Ladenpersonen;
          end;
        end;
        flauf:=0;
        fart:=0;
        ftasts:=True;
        FBLoe;
      end else begin
        flauf:=14;
        Edit12.Enabled:=True;
        Edit12.ReadOnly:=False;
        Edit12.TabStop:=True;
        Edit12.Color:=$0080FFFF;
        Edit12.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=20;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit12.Text:=fperson.Partnertelefon;
          ialpha:=fperson.Partnertelefon;
          iart:=1;
        end;
      end;
    end else begin
      Label2.Caption:='Partneranrede muss ausgewählt werden';
    end;
  end;
end;

procedure TFrame2.Edit10KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 11 then
  begin
    itaste:=Key;
    FeldEingabe(Edit10);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit10KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
begin
  Shift:=Shift;
  if flauf = 11 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit10.Enabled:=False;
      Edit10.ReadOnly:=True;
      Edit10.TabStop:=False;
      Edit10.Color:=$00C0C0C0;
      h1:=Edit10.Text;
      ialpha:=Form1.Blankweg(3, h1);
      fperson.Email:=ialpha;
      flauf:=12;
      Edit11.Enabled:=True;
      Edit11.ReadOnly:=False;
      Edit11.TabStop:=True;
      Edit11.Color:=$0080FFFF;
      Edit11.Text:='';
      ialpha:='';
      inummer:=0;
      iart:=2;
      inumkom:=0;
      ikomma:=0;
      izeich:=3;
      istell:=0;
      ifunc:=5;
      ilanmax:=30;
      ilanmin:=1;
      iautocr:=1;
      if fart = 2 then
      begin
        Edit11.Text:=fperson.Land;
        ialpha:=fperson.Land;
        iart:=1;
      end;
    end;
  end;
end;

procedure TFrame2.Edit11KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 12 then
  begin
    itaste:=Key;
    FeldEingabe(Edit11);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit11KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
  var h2: integer;
begin
  Shift:=Shift;
  if flauf = 12 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit11.Enabled:=False;
      Edit11.ReadOnly:=True;
      Edit11.TabStop:=False;
      Edit11.Color:=$00C0C0C0;
      h1:=Edit11.Text;
      ialpha:=Form1.Blankweg(3, h1);
      if ialpha <> '' then
      begin
        fperson.Land:=ialpha;
        if fperson.Art = 'Person' then
        begin
          if fart = 1 then
          begin
            JaNein:=messagedlg('Person/Firma anlegen?', mtConfirmation, [mbYes, mbNo], 0);
            if (JaNein = mrYes) then
            begin
              h2:=fperson.Nummer;
              Persontab[h2].Nummer:=h2;
              Persontab[h2].Art:=fperson.Art;
              Persontab[h2].Land:=fperson.Land;
              Persontab[h2].Email:=fperson.Email;
              Persontab[h2].Handy:=fperson.Handy;
              Persontab[h2].Telefon:=fperson.Telefon;
              Persontab[h2].Ort:=fperson.Ort;
              Persontab[h2].Plz:=fperson.Plz;
              Persontab[h2].Hausnummer:=fperson.Hausnummer;
              Persontab[h2].Strasse:=fperson.Strasse;
              Persontab[h2].Nachname:=fperson.Nachname;
              Persontab[h2].Firmenname:=fperson.Firmenname;
              Persontab[h2].Vorname:=fperson.Vorname;
              Persontab[h2].Partnertelefon:=fperson.Partnertelefon;
              Persontab[h2].Partnernachname:=fperson.Partnernachname;
              Persontab[h2].Partnervorname:=fperson.Partnervorname;
              Persontab[h2].Partneranrede:=fperson.Partneranrede;
              Persontab[h2].Satzart:=True;
              Personzaehl:=Personzaehl+1;
              Updperson:=True;
              Ladenpersonen;
            end;
          end else begin
            JaNein:=messagedlg('Person/Firma ändern?', mtConfirmation, [mbYes, mbNo], 0);
            if (JaNein = mrYes) then
            begin
              h2:=fperson.Nummer;
              Persontab[h2].Nummer:=h2;
              Persontab[h2].Art:=fperson.Art;
              Persontab[h2].Land:=fperson.Land;
              Persontab[h2].Email:=fperson.Email;
              Persontab[h2].Handy:=fperson.Handy;
              Persontab[h2].Telefon:=fperson.Telefon;
              Persontab[h2].Ort:=fperson.Ort;
              Persontab[h2].Plz:=fperson.Plz;
              Persontab[h2].Hausnummer:=fperson.Hausnummer;
              Persontab[h2].Strasse:=fperson.Strasse;
              Persontab[h2].Nachname:=fperson.Nachname;
              Persontab[h2].Firmenname:=fperson.Firmenname;
              Persontab[h2].Vorname:=fperson.Vorname;
              Persontab[h2].Partnertelefon:=fperson.Partnertelefon;
              Persontab[h2].Partnernachname:=fperson.Partnernachname;
              Persontab[h2].Partnervorname:=fperson.Partnervorname;
              Persontab[h2].Partneranrede:=fperson.Partneranrede;
              Persontab[h2].Satzart:=True;
              Updperson:=True;
              Ladenpersonen;
            end;
          end;
          flauf:=0;
          fart:=0;
          ftasts:=True;
          FBLoe;
        end else begin
          flauf:=13;
          ComboBox2.Enabled:=True;
          ComboBox2.ReadOnly:=False;
          ComboBox2.TabStop:=True;
          ComboBox2.Color:=$0080FFFF;
          ComboBox2.Text:='ohne';
          h2:=ComboBox2.Items.IndexOf('ohne');
          ComboBox2.ItemIndex:=h2;
        end;
      end else begin
        Label2.Caption:='Land muss eingegeben werden';
        Edit11.Enabled:=True;
        Edit11.ReadOnly:=False;
        Edit11.TabStop:=True;
        Edit11.Color:=$0080FFFF;
        Edit11.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=30;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit11.Text:=fperson.Land;
          ialpha:=fperson.Land;
          iart:=1;
        end;
      end;
    end;
  end;
end;

procedure TFrame2.Edit12KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 14 then
  begin
    itaste:=Key;
    FeldEingabe(Edit12);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit12KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
begin
  Shift:=Shift;
  if flauf = 14 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit12.Enabled:=False;
      Edit12.ReadOnly:=True;
      Edit12.TabStop:=False;
      Edit12.Color:=$00C0C0C0;
      h1:=ialpha;
      ialpha:=Form1.Blankweg(3, h1);
      if ialpha <> '' then
      begin
        fperson.Partnertelefon:=ialpha;
        flauf:=15;
        Edit13.Enabled:=True;
        Edit13.ReadOnly:=False;
        Edit13.TabStop:=True;
        Edit13.Color:=$0080FFFF;
        Edit13.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=25;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit13.Text:=fperson.Partnervorname;
          ialpha:=fperson.Partnervorname;
          iart:=1;
        end;
      end else begin
        Label2.Caption:='Partnertelefon muss eingegeben werden';
        Edit12.Enabled:=True;
        Edit12.ReadOnly:=False;
        Edit12.TabStop:=True;
        Edit12.Color:=$0080FFFF;
        Edit12.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=20;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit12.Text:=fperson.Partnertelefon;
          ialpha:=fperson.Partnertelefon;
          iart:=1;
        end;
      end;
    end;
  end;
end;

procedure TFrame2.Edit13KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 15 then
  begin
    itaste:=Key;
    FeldEingabe(Edit13);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit13KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
begin
  Shift:=Shift;
  if flauf = 15 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit13.Enabled:=False;
      Edit13.ReadOnly:=True;
      Edit13.TabStop:=False;
      Edit13.Color:=$00C0C0C0;
      h1:=Edit13.Text;
      ialpha:=Form1.Blankweg(3, h1);
      if ialpha <> '' then
      begin
        fperson.Partnervorname:=ialpha;
        flauf:=16;
        Edit14.Enabled:=True;
        Edit14.ReadOnly:=False;
        Edit14.TabStop:=True;
        Edit14.Color:=$0080FFFF;
        Edit14.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=25;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit14.Text:=fperson.Partnernachname;
          ialpha:=fperson.Partnernachname;
          iart:=1;
        end;
      end else begin
        Label2.Caption:='Partnervorname muss eingegeben werden';
        Edit13.Enabled:=True;
        Edit13.ReadOnly:=False;
        Edit13.TabStop:=True;
        Edit13.Color:=$0080FFFF;
        Edit13.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=20;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit13.Text:=fperson.Partnertelefon;
          ialpha:=fperson.Partnertelefon;
          iart:=1;
        end;
      end;
    end;
  end;
end;

procedure TFrame2.Edit14KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 16 then
  begin
    itaste:=Key;
    FeldEingabe(Edit14);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit14KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
  var h2: integer;
begin
  Shift:=Shift;
  if flauf = 16 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit14.Enabled:=False;
      Edit14.ReadOnly:=True;
      Edit14.TabStop:=False;
      Edit14.Color:=$00C0C0C0;
      h1:=Edit14.Text;
      ialpha:=Form1.Blankweg(3, h1);
      if ialpha <> '' then
      begin
        fperson.Partnernachname:=ialpha;
        if fart = 1 then
        begin
          JaNein:=messagedlg('Person/Firma anlegen?', mtConfirmation, [mbYes, mbNo], 0);
          if (JaNein = mrYes) then
          begin
            h2:=fperson.Nummer;
            Persontab[h2].Nummer:=h2;
            Persontab[h2].Art:=fperson.Art;
            Persontab[h2].Land:=fperson.Land;
            Persontab[h2].Email:=fperson.Email;
            Persontab[h2].Handy:=fperson.Handy;
            Persontab[h2].Telefon:=fperson.Telefon;
            Persontab[h2].Ort:=fperson.Ort;
            Persontab[h2].Plz:=fperson.Plz;
            Persontab[h2].Hausnummer:=fperson.Hausnummer;
            Persontab[h2].Strasse:=fperson.Strasse;
            Persontab[h2].Nachname:=fperson.Nachname;
            Persontab[h2].Firmenname:=fperson.Firmenname;
            Persontab[h2].Vorname:=fperson.Vorname;
            Persontab[h2].Partnertelefon:=fperson.Partnertelefon;
            Persontab[h2].Partnernachname:=fperson.Partnernachname;
            Persontab[h2].Partnervorname:=fperson.Partnervorname;
            Persontab[h2].Partneranrede:=fperson.Partneranrede;
            Persontab[h2].Satzart:=True;
            Personzaehl:=Personzaehl+1;
            Updperson:=True;
            Ladenpersonen;
          end;
        end else begin
          JaNein:=messagedlg('Person/Firma ändern?', mtConfirmation, [mbYes, mbNo], 0);
          if (JaNein = mrYes) then
          begin
            h2:=fperson.Nummer;
            Persontab[h2].Nummer:=h2;
            Persontab[h2].Art:=fperson.Art;
            Persontab[h2].Land:=fperson.Land;
            Persontab[h2].Email:=fperson.Email;
            Persontab[h2].Handy:=fperson.Handy;
            Persontab[h2].Telefon:=fperson.Telefon;
            Persontab[h2].Ort:=fperson.Ort;
            Persontab[h2].Plz:=fperson.Plz;
            Persontab[h2].Hausnummer:=fperson.Hausnummer;
            Persontab[h2].Strasse:=fperson.Strasse;
            Persontab[h2].Nachname:=fperson.Nachname;
            Persontab[h2].Firmenname:=fperson.Firmenname;
            Persontab[h2].Vorname:=fperson.Vorname;
            Persontab[h2].Partnertelefon:=fperson.Partnertelefon;
            Persontab[h2].Partnernachname:=fperson.Partnernachname;
            Persontab[h2].Partnervorname:=fperson.Partnervorname;
            Persontab[h2].Partneranrede:=fperson.Partneranrede;
            Persontab[h2].Satzart:=True;
            Updperson:=True;
            Ladenpersonen;
          end;
        end;
        flauf:=0;
        fart:=0;
        ftasts:=True;
        FBLoe;
      end else begin
        Label2.Caption:='Partnernachname muss eingegeben werden';
        Edit14.Enabled:=True;
        Edit14.ReadOnly:=False;
        Edit14.TabStop:=True;
        Edit14.Color:=$0080FFFF;
        Edit14.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=25;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit14.Text:=fperson.Partnernachname;
          ialpha:=fperson.Partnernachname;
          iart:=1;
        end;
      end;
    end;
  end;
end;

procedure TFrame2.Edit1KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 3 then
  begin
    itaste:=Key;
    FeldEingabe(Form1.Frame2.Edit1);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit1KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
begin
  Shift:=Shift;
  if flauf = 3 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit1.Enabled:=False;
      Edit1.ReadOnly:=True;
      Edit1.TabStop:=False;
      Edit1.Color:=$00C0C0C0;
      h1:=Edit1.Text;
      ialpha:=Form1.Blankweg(3, h1);
      if ialpha <> '' then
      begin
        fperson.Vorname:=ialpha;
        flauf:=4;
        Edit2.Enabled:=True;
        Edit2.ReadOnly:=False;
        Edit2.TabStop:=True;
        Edit2.Color:=$0080FFFF;
        Edit2.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=25;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit2.Text:=fperson.Nachname;
          ialpha:=fperson.Nachname;
          iart:=1;
        end;
      end else begin
        Label2.Caption:='Vorname muss eingegeben werden';
        Edit1.Enabled:=True;
        Edit1.ReadOnly:=False;
        Edit1.TabStop:=True;
        Edit1.Color:=$0080FFFF;
        Edit1.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=25;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit1.Text:=fperson.Vorname;
          ialpha:=fperson.Vorname;
          iart:=1;
        end;
      end;
    end;
  end;
end;

procedure TFrame2.Edit2KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 4 then
  begin
    itaste:=Key;
    FeldEingabe(Form1.Frame2.Edit2);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit2KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
begin
  Shift:=Shift;
  if flauf = 4 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit2.Enabled:=False;
      Edit2.ReadOnly:=True;
      Edit2.TabStop:=False;
      Edit2.Color:=$00C0C0C0;
      h1:=Edit2.Text;
      ialpha:=Form1.Blankweg(3, h1);
      if ialpha <> '' then
      begin
        fperson.Nachname:=ialpha;
        flauf:=5;
        Edit3.Enabled:=True;
        Edit3.ReadOnly:=False;
        Edit3.TabStop:=True;
        Edit3.Color:=$0080FFFF;
        Edit3.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=30;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit3.Text:=fperson.Strasse;
          ialpha:=fperson.Strasse;
          iart:=1;
        end;
      end else begin
        Label2.Caption:='Nachname muss eingegeben werden';
        Edit2.Enabled:=True;
        Edit2.ReadOnly:=False;
        Edit2.TabStop:=True;
        Edit2.Color:=$0080FFFF;
        Edit2.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=25;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit2.Text:=fperson.Nachname;
          ialpha:=fperson.Nachname;
          iart:=1;
        end;
      end;
    end;
  end;
end;

procedure TFrame2.Edit3KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 5 then
  begin
    itaste:=Key;
    FeldEingabe(Edit3);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit3KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
begin
  Shift:=Shift;
  if flauf = 5 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit3.Enabled:=False;
      Edit3.ReadOnly:=True;
      Edit3.TabStop:=False;
      Edit3.Color:=$00C0C0C0;
      h1:=Edit3.Text;
      ialpha:=Form1.Blankweg(3, h1);
      if ialpha <> '' then
      begin
        fperson.Strasse:=ialpha;
        flauf:=6;
        Edit4.Enabled:=True;
        Edit4.ReadOnly:=False;
        Edit4.TabStop:=True;
        Edit4.Color:=$0080FFFF;
        Edit4.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=5;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit4.Text:=fperson.Hausnummer;
          ialpha:=fperson.Hausnummer;
          iart:=1;
        end;
      end else begin
        Label2.Caption:='Strasse muss eingegeben werden';
        Edit3.Enabled:=True;
        Edit3.ReadOnly:=False;
        Edit3.TabStop:=True;
        Edit3.Color:=$0080FFFF;
        Edit3.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=30;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit3.Text:=fperson.Strasse;
          ialpha:=fperson.Strasse;
          iart:=1;
        end;
      end;
    end;
  end;
end;

procedure TFrame2.Edit4KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 6 then
  begin
    itaste:=Key;
    FeldEingabe(Edit4);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit4KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
begin
  Shift:=Shift;
  if flauf = 6 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit4.Enabled:=False;
      Edit4.ReadOnly:=True;
      Edit4.TabStop:=False;
      Edit4.Color:=$00C0C0C0;
      h1:=ialpha;
      ialpha:=Form1.Blankweg(3, h1);
      if ialpha <> '' then
      begin
        fperson.Hausnummer:=ialpha;
        flauf:=7;
        Edit5.Enabled:=True;
        Edit5.ReadOnly:=False;
        Edit5.TabStop:=True;
        Edit5.Color:=$0080FFFF;
        Edit5.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=10;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit5.Text:=fperson.Plz;
          ialpha:=fperson.Plz;
          iart:=1;
        end;
      end else begin
        Label2.Caption:='Hausnummer muss eingegeben werden';
        Edit4.Enabled:=True;
        Edit4.ReadOnly:=False;
        Edit4.TabStop:=True;
        Edit4.Color:=$0080FFFF;
        Edit4.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=5;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit4.Text:=fperson.Hausnummer;
          ialpha:=fperson.Hausnummer;
          iart:=1;
        end;
      end;
    end;
  end;
end;

procedure TFrame2.Edit5KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 7 then
  begin
    itaste:=Key;
    FeldEingabe(Edit5);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit5KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
begin
  Shift:=Shift;
  if flauf = 7 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit5.Enabled:=False;
      Edit5.ReadOnly:=True;
      Edit5.TabStop:=False;
      Edit5.Color:=$00C0C0C0;
      h1:=ialpha;
      ialpha:=Form1.Blankweg(3, h1);
      if ialpha <> '' then
      begin
        fperson.Plz:=ialpha;
        flauf:=8;
        Edit6.Enabled:=True;
        Edit6.ReadOnly:=False;
        Edit6.TabStop:=True;
        Edit6.Color:=$0080FFFF;
        Edit6.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=30;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit6.Text:=fperson.Ort;
          ialpha:=fperson.Ort;
          iart:=1;
        end;
      end else begin
        Label2.Caption:='Postleitzahl muss eingegeben werden';
        Edit5.Enabled:=True;
        Edit5.ReadOnly:=False;
        Edit5.TabStop:=True;
        Edit5.Color:=$0080FFFF;
        Edit5.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=10;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit5.Text:=fperson.Plz;
          ialpha:=fperson.Plz;
          iart:=1;
        end;
      end;
    end;
  end;
end;

procedure TFrame2.Edit6KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 8 then
  begin
    itaste:=Key;
    FeldEingabe(Edit6);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit6KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
begin
  Shift:=Shift;
  if flauf = 8 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit6.Enabled:=False;
      Edit6.ReadOnly:=True;
      Edit6.TabStop:=False;
      Edit6.Color:=$00C0C0C0;
      h1:=Edit6.Text;
      ialpha:=Form1.Blankweg(3, h1);
      if ialpha <> '' then
      begin
        fperson.Ort:=ialpha;
        flauf:=9;
        Edit7.Enabled:=True;
        Edit7.ReadOnly:=False;
        Edit7.TabStop:=True;
        Edit7.Color:=$0080FFFF;
        Edit7.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=20;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit7.Text:=fperson.Telefon;
          ialpha:=fperson.Telefon;
          iart:=1;
        end;
      end else begin
        Label2.Caption:='Ort muss eingegeben werden';
        Edit6.Enabled:=True;
        Edit6.ReadOnly:=False;
        Edit6.TabStop:=True;
        Edit6.Color:=$0080FFFF;
        Edit6.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=30;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit6.Text:=fperson.Ort;
          ialpha:=fperson.Ort;
          iart:=1;
        end;
      end;
    end;
  end;
end;

procedure TFrame2.Edit7KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 9 then
  begin
    itaste:=Key;
    FeldEingabe(Edit7);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit7KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
begin
  Shift:=Shift;
  if flauf = 9 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit7.Enabled:=False;
      Edit7.ReadOnly:=True;
      Edit7.TabStop:=False;
      Edit7.Color:=$00C0C0C0;
      h1:=ialpha;
      ialpha:=Form1.Blankweg(3, h1);
      fperson.Telefon:=ialpha;
      flauf:=10;
      Edit8.Enabled:=True;
      Edit8.ReadOnly:=False;
      Edit8.TabStop:=True;
      Edit8.Color:=$0080FFFF;
      Edit8.Text:='';
      ialpha:='';
      inummer:=0;
      iart:=2;
      inumkom:=0;
      ikomma:=0;
      izeich:=3;
      istell:=0;
      ifunc:=5;
      ilanmax:=20;
      ilanmin:=1;
      iautocr:=1;
      if fart = 2 then
      begin
        Edit8.Text:=fperson.Handy;
        ialpha:=fperson.Handy;
        iart:=1;
      end;
    end;
  end;
end;

procedure TFrame2.Edit8KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 10 then
  begin
    itaste:=Key;
    FeldEingabe(Edit8);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit8KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
begin
  Shift:=Shift;
  if flauf = 10 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit8.Enabled:=False;
      Edit8.ReadOnly:=True;
      Edit8.TabStop:=False;
      Edit8.Color:=$00C0C0C0;
      h1:=ialpha;
      ialpha:=Form1.Blankweg(3, h1);
      fperson.Handy:=ialpha;
      if ((fperson.Telefon = '') and (fperson.Handy = '')) then
      begin
        Label2.Caption:='Telefon und/oder Handy muss eingegeben werden';
        flauf:=9;
        Edit7.Enabled:=True;
        Edit7.ReadOnly:=False;
        Edit7.TabStop:=True;
        Edit7.Color:=$0080FFFF;
        Edit7.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=20;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit7.Text:=fperson.Telefon;
          ialpha:=fperson.Telefon;
          iart:=1;
        end;
      end else begin
        flauf:=11;
        Edit10.Enabled:=True;
        Edit10.ReadOnly:=False;
        Edit10.TabStop:=True;
        Edit10.Color:=$0080FFFF;
        Edit10.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=30;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit10.Text:=fperson.Email;
          ialpha:=fperson.Email;
          iart:=1;
        end;
      end;
    end;
  end;
end;

procedure TFrame2.Edit9KeyPress(Sender: TObject; var Key: char);
begin
  if flauf = 2 then
  begin
    itaste:=Key;
    FeldEingabe(Edit9);
    Key:=itaste;
  end;
end;

procedure TFrame2.Edit9KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  var h1: string;
begin
  Shift:=Shift;
  if flauf = 2 then
  begin
    if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
    begin
      Key:=ord(chr(13));
    end;
    if ord(Key) = 13 then
    begin
      Label2.Caption:='';
      istell:=0;
      Edit9.Enabled:=False;
      Edit9.ReadOnly:=True;
      Edit9.TabStop:=False;
      Edit9.Color:=$00C0C0C0;
      h1:=Edit9.Text;
      ialpha:=Form1.Blankweg(3, h1);
      if ialpha <> '' then
      begin
        fperson.Firmenname:=ialpha;
        flauf:=3;
        Edit1.Enabled:=True;
        Edit1.ReadOnly:=False;
        Edit1.TabStop:=True;
        Edit1.Color:=$0080FFFF;
        Edit1.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=25;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit1.Text:=fperson.Vorname;
          ialpha:=fperson.Vorname;
          iart:=1;
        end;
      end else begin
        Label2.Caption:='Firmenname muss eingegeben werden';
        Edit9.Enabled:=True;
        Edit9.ReadOnly:=False;
        Edit9.TabStop:=True;
        Edit9.Color:=$0080FFFF;
        Edit9.Text:='';
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=5;
        ilanmax:=30;
        ilanmin:=1;
        iautocr:=1;
        if fart = 2 then
        begin
          Edit9.Text:=fperson.Firmenname;
          ialpha:=fperson.Firmenname;
          iart:=1;
        end;
      end;
    end;
  end;
end;

procedure TFrame2.Button1Click(Sender: TObject);
  var h1: integer;
  var h2: integer;
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  if ftasts  then
  begin
    Label2.Caption:='';
    h2:=0;
    for h1:=1 to Personmax do
    begin
      if h2 = 0 then
      begin
        if not Persontab[h1].Satzart then
        begin
          h2:=h1;
        end;
      end;
    end;
    if h2 > 0 then
    begin
      fperson.Nummer:=h2;
      fperson.Firmenname:='';
      fperson.Partneranrede:='';
      fperson.Partnervorname:='';
      fperson.Partnernachname:='';
      fperson.Partnertelefon:='';
      fperson.Satzart:=True;
      Label22.Caption:=IntToStr(h2);
      ftasts:=False;
      fart:=1;
      flauf:=1;
      ComboBox1.Enabled:=True;
      ComboBox1.ReadOnly:=False;
      ComboBox1.TabStop:=True;
      ComboBox1.Color:=$0080FFFF;
      ComboBox1.Text:='';
      Form1.ActiveControl:=Form1.Frame2.ComboBox1;
    end else begin
      Label2.Caption:='kein anlegen Person mehr möglich, maximale Anzahl erreicht';
    end;
  end;
end;

procedure TFrame2.Button2Click(Sender: TObject);
  var h1: string;
  var h2: integer;
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  if ftasts  then
  begin
    Label2.Caption:='';
    if fgrid > 0 then
    begin
      h1:=StringGrid1.Cells[0, fgrid];
      h2:=Length(h1);
      if h2 > 0 then
      begin
        h2:=StrToInt(h1);
        if ((h2 > 0) and (h2 < 10000)) then
        begin
          ftasts:=False;
          fart:=2;
          flauf:=1;
          fperson.Nummer:=h2;
          fperson.Art:=Persontab[h2].Art;
          fperson.Satzart:=Persontab[h2].Satzart;
          fperson.Vorname:=Persontab[h2].Vorname;
          fperson.Nachname:=Persontab[h2].Nachname;
          fperson.Strasse:=Persontab[h2].Strasse;
          fperson.Hausnummer:=Persontab[h2].Hausnummer;
          fperson.Plz:=Persontab[h2].Plz;
          fperson.Ort:=Persontab[h2].Ort;
          fperson.Telefon:=Persontab[h2].Telefon;
          fperson.Handy:=Persontab[h2].Handy;
          fperson.Email:=Persontab[h2].Email;
          fperson.Land:=Persontab[h2].Land;
          fperson.Firmenname:=Persontab[h2].Firmenname;
          fperson.Partneranrede:=Persontab[h2].Partneranrede;
          fperson.Partnervorname:=Persontab[h2].Partnervorname;
          fperson.Partnernachname:=Persontab[h2].Partnernachname;
          fperson.Partnertelefon:=Persontab[h2].Partnertelefon;
          fperson.Satzart:=True;
          ComboBox2.Text:=fperson.Partneranrede;
          Edit1.Text:=fperson.Vorname;
          Edit2.Text:=fperson.Nachname;
          Edit3.Text:=fperson.Strasse;
          Edit4.Text:=fperson.Hausnummer;
          Edit5.Text:=fperson.Plz;
          Edit6.Text:=fperson.Ort;
          Edit7.Text:=fperson.Telefon;
          Edit8.Text:=fperson.Handy;
          Edit9.Text:=fperson.Firmenname;
          Edit10.Text:=fperson.Email;
          Edit11.Text:=fperson.Land;
          Edit12.Text:=fperson.Partnertelefon;
          Edit13.Text:=fperson.Partnervorname;
          Edit14.Text:=fperson.Partnernachname;
          Label22.Caption:=IntToStr(h2);
          ComboBox1.Enabled:=True;
          ComboBox1.ReadOnly:=False;
          ComboBox1.TabStop:=True;
          ComboBox1.Color:=$0080FFFF;
          ComboBox1.Text:=fperson.Art;
          h2:=ComboBox1.Items.IndexOf(fperson.Art);
          ComboBox1.ItemIndex:=h2;
          Form1.ActiveControl:=Form1.Frame2.ComboBox1;
        end else begin
          Label2.Caption:='Perosnennummer nicht zulässig';
        end;
      end else begin
        Label2.Caption:='Perosnennummer ist leer';
      end;
    end else begin
      Label2.Caption:='erst Person/Firma im Stringgrid anklicken';
    end;
  end;
end;

procedure TFrame2.Button3Click(Sender: TObject);
  var h1: string;
  var h2: integer;
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  if ftasts  then
  begin
    Label2.Caption:='';
    if fgrid > 0 then
    begin
      h1:=StringGrid1.Cells[0, fgrid];
      h2:=Length(h1);
      if h2 > 0 then
      begin
        h2:=StrToInt(h1);
        if ((h2 > 0) and (h2 < 10000)) then
        begin
          ftasts:=False;
          fart:=3;
          flauf:=0;
          fperson.Nummer:=h2;
          fperson.Art:=Persontab[h2].Art;
          fperson.Satzart:=Persontab[h2].Satzart;
          fperson.Vorname:=Persontab[h2].Vorname;
          fperson.Nachname:=Persontab[h2].Nachname;
          fperson.Strasse:=Persontab[h2].Strasse;
          fperson.Hausnummer:=Persontab[h2].Hausnummer;
          fperson.Plz:=Persontab[h2].Plz;
          fperson.Ort:=Persontab[h2].Ort;
          fperson.Telefon:=Persontab[h2].Telefon;
          fperson.Handy:=Persontab[h2].Handy;
          fperson.Email:=Persontab[h2].Email;
          fperson.Land:=Persontab[h2].Land;
          fperson.Firmenname:=Persontab[h2].Firmenname;
          fperson.Partneranrede:=Persontab[h2].Partneranrede;
          fperson.Partnervorname:=Persontab[h2].Partnervorname;
          fperson.Partnernachname:=Persontab[h2].Partnernachname;
          fperson.Partnertelefon:=Persontab[h2].Partnertelefon;
          fperson.Satzart:=True;
          ComboBox1.Text:=fperson.Art;
          ComboBox2.Text:=fperson.Partneranrede;
          Edit1.Text:=fperson.Vorname;
          Edit2.Text:=fperson.Nachname;
          Edit3.Text:=fperson.Strasse;
          Edit4.Text:=fperson.Hausnummer;
          Edit5.Text:=fperson.Plz;
          Edit6.Text:=fperson.Ort;
          Edit7.Text:=fperson.Telefon;
          Edit8.Text:=fperson.Handy;
          Edit9.Text:=fperson.Firmenname;
          Edit10.Text:=fperson.Email;
          Edit11.Text:=fperson.Land;
          Edit12.Text:=fperson.Partnertelefon;
          Edit13.Text:=fperson.Partnervorname;
          Edit14.Text:=fperson.Partnernachname;
          Label22.Caption:=IntToStr(h2);
          JaNein:=messagedlg('Person/Firma löschen?', mtConfirmation, [mbYes, mbNo], 0);
          if (JaNein = mrYes) then
          begin
            h2:=fperson.Nummer;
            Persontab[h2].Satzart:=False;
            Personzaehl:=Personzaehl-1;
            Updperson:=True;
            Ladenpersonen;
          end;
          flauf:=0;
          fart:=0;
          ftasts:=True;
          FBLoe;
        end else begin
          Label2.Caption:='Perosnennummer nicht zulässig';
        end;
      end else begin
        Label2.Caption:='Perosnennummer ist leer';
      end;
    end else begin
      Label2.Caption:='erst Person/Firma im Stringgrid anklicken';
    end;
  end;
end;

procedure TFrame2.StringGrid1SelectCell(Sender: TObject; aCol, aRow: Integer;
  var CanSelect: Boolean);
begin
  aCol:=aCol;
  fgrid:=aRow;
  CanSelect:=True;
end;

procedure TFrame2.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled:=False;
  Ladenpersonen;
end;

constructor TFrame2.Create(AOwner1: TComponent);
begin
  inherited Create(AOwner);
  AOwner1:=AOwner1;
  ftasts:=true;
  fart:=0;
  flauf:=0;
  fgrid:=0;
  self.ComboBox1.Enabled:=False;
  self.ComboBox1.Color:=$00C0C0C0;
  self.ComboBox1.Text:='';
  self.ComboBox1.ItemIndex:=-1;
  self.ComboBox1.Items.Clear;
  self.ComboBox1.Items.Add('Person');
  self.ComboBox1.Items.Add('Firma');
  self.ComboBox2.Enabled:=False;
  self.ComboBox2.Color:=$00C0C0C0;
  self.ComboBox2.Text:='';
  self.ComboBox2.ItemIndex:=-1;
  self.ComboBox2.Items.Clear;
  self.ComboBox2.Items.Add('ohne');
  self.ComboBox2.Items.Add('Herr');
  self.ComboBox2.Items.Add('Frau');
  self.Edit1.Enabled:=False;
  self.Edit1.Color:=$00C0C0C0;
  self.Edit1.Text:='';
  self.Edit2.Enabled:=False;
  self.Edit2.Color:=$00C0C0C0;
  self.Edit2.Text:='';
  self.Edit3.Enabled:=False;
  self.Edit3.Color:=$00C0C0C0;
  self.Edit3.Text:='';
  self.Edit4.Enabled:=False;
  self.Edit4.Color:=$00C0C0C0;
  self.Edit4.Text:='';
  self.Edit5.Enabled:=False;
  self.Edit5.Color:=$00C0C0C0;
  self.Edit5.Text:='';
  self.Edit6.Enabled:=False;
  self.Edit6.Color:=$00C0C0C0;
  self.Edit6.Text:='';
  self.Edit7.Enabled:=False;
  self.Edit7.Color:=$00C0C0C0;
  self.Edit7.Text:='';
  self.Edit8.Enabled:=False;
  self.Edit8.Color:=$00C0C0C0;
  self.Edit8.Text:='';
  self.Edit9.Enabled:=False;
  self.Edit9.Color:=$00C0C0C0;
  self.Edit9.Text:='';
  self.Edit10.Enabled:=False;
  self.Edit10.Color:=$00C0C0C0;
  self.Edit10.Text:='';
  self.Edit11.Enabled:=False;
  self.Edit11.Color:=$00C0C0C0;
  self.Edit11.Text:='';
  self.Edit12.Enabled:=False;
  self.Edit12.Color:=$00C0C0C0;
  self.Edit12.Text:='';
  self.Edit13.Enabled:=False;
  self.Edit13.Color:=$00C0C0C0;
  self.Edit13.Text:='';
  self.Edit14.Enabled:=False;
  self.Edit14.Color:=$00C0C0C0;
  self.Edit14.Text:='';
  self.Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  self.Label2.Caption:='';
  self.StringGrid1.RowCount:=2;
  self.StringGrid1.ColCount:=17;
  self.StringGrid1.ColWidths[0]:=120;
  self.StringGrid1.ColWidths[1]:=110;
  self.StringGrid1.ColWidths[2]:=500;
  self.StringGrid1.ColWidths[3]:=430;
  self.StringGrid1.ColWidths[4]:=430;
  self.StringGrid1.ColWidths[5]:=500;
  self.StringGrid1.ColWidths[6]:=170;
  self.StringGrid1.ColWidths[7]:=180;
  self.StringGrid1.ColWidths[8]:=500;
  self.StringGrid1.ColWidths[9]:=350;
  self.StringGrid1.ColWidths[10]:=350;
  self.StringGrid1.ColWidths[11]:=500;
  self.StringGrid1.ColWidths[12]:=500;
  self.StringGrid1.ColWidths[13]:=130;
  self.StringGrid1.ColWidths[14]:=350;
  self.StringGrid1.ColWidths[15]:=430;
  self.StringGrid1.ColWidths[16]:=430;
  self.StringGrid1.Cells[0, 0]:='Nummer';
  self.StringGrid1.Cells[1, 0]:='Art';
  self.StringGrid1.Cells[2, 0]:='Firmenname';
  self.StringGrid1.Cells[3, 0]:='Vorname';
  self.StringGrid1.Cells[4, 0]:='Nachname';
  self.StringGrid1.Cells[5, 0]:='Strasse';
  self.StringGrid1.Cells[6, 0]:='Hausnummer';
  self.StringGrid1.Cells[7, 0]:='PLZ';
  self.StringGrid1.Cells[8, 0]:='Ort';
  self.StringGrid1.Cells[9, 0]:='Telefon';
  self.StringGrid1.Cells[10, 0]:='Handy';
  self.StringGrid1.Cells[11, 0]:='EMail';
  self.StringGrid1.Cells[12, 0]:='Land';
  self.StringGrid1.Cells[13, 0]:='PAnrede';
  self.StringGrid1.Cells[14, 0]:='PTelefon';
  self.StringGrid1.Cells[15, 0]:='Partnervorname';
  self.StringGrid1.Cells[16, 0]:='Partnernachname';
  self.StringGrid1.Cells[0, 1]:='';
  self.StringGrid1.Cells[1, 1]:='';
  self.StringGrid1.Cells[2, 1]:='';
  self.StringGrid1.Cells[3, 1]:='';
  self.StringGrid1.Cells[4, 1]:='';
  self.StringGrid1.Cells[5, 1]:='';
  self.StringGrid1.Cells[6, 1]:='';
  self.StringGrid1.Cells[7, 1]:='';
  self.StringGrid1.Cells[8, 1]:='';
  self.StringGrid1.Cells[9, 1]:='';
  self.StringGrid1.Cells[10, 1]:='';
  self.StringGrid1.Cells[11, 1]:='';
  self.StringGrid1.Cells[12, 1]:='';
  self.StringGrid1.Cells[13, 1]:='';
  self.StringGrid1.Cells[14, 1]:='';
  self.StringGrid1.Cells[15, 1]:='';
  self.StringGrid1.Cells[16, 1]:='';
  self.Button2.Visible:=False;
  self.Button3.Visible:=False;
  if Personzaehl > 0 then
  begin
    self.Button2.Visible:=True;
    self.Button3.Visible:=True;
  end;
end;

procedure TFrame2.FBLoe;
begin
  ComboBox1.Enabled:=False;
  ComboBox1.ReadOnly:=True;
  ComboBox1.TabStop:=False;
  ComboBox1.Color:=$00C0C0C0;
  ComboBox1.Text:='';
  ComboBox2.Enabled:=False;
  ComboBox2.ReadOnly:=True;
  ComboBox2.TabStop:=False;
  ComboBox2.Color:=$00C0C0C0;
  ComboBox2.Text:='';
  Edit1.Enabled:=False;
  Edit1.ReadOnly:=True;
  Edit1.TabStop:=False;
  Edit1.Color:=$00C0C0C0;
  Edit1.Text:='';
  Edit2.Enabled:=False;
  Edit2.ReadOnly:=True;
  Edit2.TabStop:=False;
  Edit2.Color:=$00C0C0C0;
  Edit2.Text:='';
  Edit3.Enabled:=False;
  Edit3.ReadOnly:=True;
  Edit3.TabStop:=False;
  Edit3.Color:=$00C0C0C0;
  Edit3.Text:='';
  Edit4.Enabled:=False;
  Edit4.ReadOnly:=True;
  Edit4.TabStop:=False;
  Edit4.Color:=$00C0C0C0;
  Edit4.Text:='';
  Edit5.Enabled:=False;
  Edit5.ReadOnly:=True;
  Edit5.TabStop:=False;
  Edit5.Color:=$00C0C0C0;
  Edit5.Text:='';
  Edit6.Enabled:=False;
  Edit6.ReadOnly:=True;
  Edit6.TabStop:=False;
  Edit6.Color:=$00C0C0C0;
  Edit6.Text:='';
  Edit7.Enabled:=False;
  Edit7.ReadOnly:=True;
  Edit7.TabStop:=False;
  Edit7.Color:=$00C0C0C0;
  Edit7.Text:='';
  Edit8.Enabled:=False;
  Edit8.ReadOnly:=True;
  Edit8.TabStop:=False;
  Edit8.Color:=$00C0C0C0;
  Edit8.Text:='';
  Edit9.Enabled:=False;
  Edit9.ReadOnly:=True;
  Edit9.TabStop:=False;
  Edit9.Color:=$00C0C0C0;
  Edit9.Text:='';
  Edit10.Enabled:=False;
  Edit10.ReadOnly:=True;
  Edit10.TabStop:=False;
  Edit10.Color:=$00C0C0C0;
  Edit10.Text:='';
  Edit11.Enabled:=False;
  Edit11.ReadOnly:=True;
  Edit11.TabStop:=False;
  Edit11.Color:=$00C0C0C0;
  Edit11.Text:='';
  Edit12.Enabled:=False;
  Edit12.ReadOnly:=True;
  Edit12.TabStop:=False;
  Edit12.Color:=$00C0C0C0;
  Edit12.Text:='';
  Edit13.Enabled:=False;
  Edit13.ReadOnly:=True;
  Edit13.TabStop:=False;
  Edit13.Color:=$00C0C0C0;
  Edit13.Text:='';
  Edit14.Enabled:=False;
  Edit14.ReadOnly:=True;
  Edit14.TabStop:=False;
  Edit14.Color:=$00C0C0C0;
  Edit14.Text:='';
  Label22.Caption:='';
end;

procedure TFrame2.Ladenpersonen;
  var z: integer;
  var h1: integer;
  var h2: string;
begin
  z:=1;
  Screen.Cursor:=crHourGlass;
  Form1.Frame2.Cursor:=crHourGlass;
  Form1.Frame2.Refresh;
  for h1:=1 to Personmax do
  begin
    if Persontab[h1].Satzart then
    begin
      z:=z + 1;
      StringGrid1.RowCount:=z;
      z:=z - 1;
      StringGrid1.Cells[0, z]:=IntToStr(Persontab[h1].Nummer);
      StringGrid1.Cells[1, z]:=Persontab[h1].Art;
      StringGrid1.Cells[2, z]:=Persontab[h1].Firmenname;
      StringGrid1.Cells[3, z]:=Persontab[h1].Vorname;
      StringGrid1.Cells[4, z]:=Persontab[h1].Nachname;
      StringGrid1.Cells[5, z]:=Persontab[h1].Strasse;
      StringGrid1.Cells[6, z]:=Persontab[h1].Hausnummer;
      StringGrid1.Cells[7, z]:=Persontab[h1].Plz;
      StringGrid1.Cells[8, z]:=Persontab[h1].Ort;
      StringGrid1.Cells[9, z]:=Persontab[h1].Telefon;
      StringGrid1.Cells[10, z]:=Persontab[h1].Handy;
      StringGrid1.Cells[11, z]:=Persontab[h1].Email;
      StringGrid1.Cells[12, z]:=Persontab[h1].Land;
      h2:=Persontab[h1].Partneranrede;
      if h2 = 'ohne' then h2:='';
      StringGrid1.Cells[13, z]:=h2;
      StringGrid1.Cells[14, z]:=Persontab[h1].Partnertelefon;
      StringGrid1.Cells[15, z]:=Persontab[h1].Partnervorname;
      StringGrid1.Cells[16, z]:=Persontab[h1].Partnernachname;
      z:=z + 1;
    end;
    Label20.Caption:=IntToStr(Personzaehl);
  end;
  Screen.Cursor:=crDefault;
  Form1.Frame2.Cursor:=crDefault;
  Form1.Frame2.Refresh;
  if z = 1 then
  begin
    StringGrid1.RowCount:=2;
    StringGrid1.Cells[0, 1]:='';
    StringGrid1.Cells[1, 1]:='';
    StringGrid1.Cells[2, 1]:='';
    StringGrid1.Cells[3, 1]:='';
    StringGrid1.Cells[4, 1]:='';
    StringGrid1.Cells[5, 1]:='';
    StringGrid1.Cells[6, 1]:='';
    StringGrid1.Cells[7, 1]:='';
    StringGrid1.Cells[8, 1]:='';
    StringGrid1.Cells[9, 1]:='';
    StringGrid1.Cells[10, 1]:='';
    StringGrid1.Cells[11, 1]:='';
    StringGrid1.Cells[12, 1]:='';
    StringGrid1.Cells[13, 1]:='';
    StringGrid1.Cells[14, 1]:='';
    StringGrid1.Cells[15, 1]:='';
    StringGrid1.Cells[16, 1]:='';
  end;
  fgrid:=0;
end;

end.

