unit Unit2;

{*

Hersteller:
           LinSoft Jürgen linder (juelin) dg5uap@darc.de

Lizens:
           Opensource entsprechend

Hardware:
           keine zusätzliche Hardware

Software:
           wird von Form1 in Unit1.pas aufgerufen (in FormActivate)
           und kehrt dahin mit dem Aufruf Form1.Frame1Close zurück

wie Compelieren:
           Lazarus IDE mit FPC
           Wird mit Unit1 zusammen compeliert

wie Ausführen:
           Kann nicht selbstständig ausgeführt werden

*}

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, StdCtrls, Eingabe;

type

  { TFrame1 }

  TFrame1 = class(TFrame)
    Button1: TButton;
    Edit1: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    procedure Button1Click(Sender: TObject);                                    // Anlegen Projektindexfile mit Edit1 Einträgen
    procedure Edit1KeyPress(Sender: TObject; var Key: char);                    // Prüfen der Eingabe jede Taste
    procedure Edit1KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);   // Prüfen der Eingabe und abspeichern
  private
    AOwner: TComponent;
  public
     constructor Create(AOwner1: TComponent); override;                         // Frame initialisieren (Variable und Komponenteneigenschaften)
     destructor Destroy; override;                                              // Frame schliessen
  end;

var
  Ausgang: Boolean;                                                             // Darf Button1 betätigt werden

implementation

{$R *.lfm}

uses Unit1;

{ TFrame1 }

destructor TFrame1.Destroy;
begin
  AOwner.Free;
  inherited Destroy;
end;

constructor TFrame1.Create(AOwner1: TComponent);
begin
  inherited Create(AOwner);
  AOwner1:=AOwner1;
  self.Label4.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  self.Label2.Caption:='';
  self.Edit1.Enabled:=True;
  self.Edit1.Color:=$0080FFFF;
  self.Edit1.Font.Color:=$00FF0000;
  self.Edit1.ReadOnly:=False;
  self.Edit1.Text:='';
  ialpha:='';
  inummer:=0;
  iart:=2;
  inumkom:=0;
  ikomma:=0;
  izeich:=3;
  istell:=0;
  ifunc:=1;
  ilanmax:=6;
  ilanmin:=1;
  iautocr:=1;
  Ausgang:=False;
end;

procedure TFrame1.Button1Click(Sender: TObject);
  var h1: string;
  var h2: string;
  var h3: integer;
  var h4: string;
begin
  if ausgang then
  begin
    h4:='';
    h1:=Indexpfad+'Projekt.index';
    assignFile(Fileadresse, h1);
    Rewrite(Fileadresse);
    for h3:=1 to inummer do
    begin
      h2:=IntToStr(h3)+' '+h4;
      Writeln(Fileadresse, h2);
    end;
    CloseFile(Fileadresse);
    Form1.Frame1Close;
  end;
end;

procedure TFrame1.Edit1KeyPress(Sender: TObject; var Key: char);
begin
  itaste:=Key;
  FeldEingabe(Edit1);
  Key:=itaste;
end;

procedure TFrame1.Edit1KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  Shift:=Shift;
  if ((istell = ilanmax) and (ord(itaste) > 0) and (iautocr = 1)) then
  begin
    Key:=ord(chr(13));
  end;
  if ord(Key) = 13 then
  begin
    Label2.Caption:='';
    istell:=0;
    Edit1.ReadOnly:=True;
    Edit1.Color:=$00C0C0C0;
    Edit1.Font.Color:=$00FF0000;
    Edit1.AutoSelect:=False;
    if inummer = 0 then
    begin
      Label2.Caption:='Anzahl Projekte muss eingegeben werden';
      Edit1.Color:=$0080FFFF;
      Edit1.Font.Color:=$00FF0000;
      Edit1.ReadOnly:=False;
      Edit1.Text:='';
      Edit1.AutoSelect:=True;
      ialpha:='';
      inummer:=0;
      iart:=2;
      inumkom:=0;
      ikomma:=0;
      izeich:=3;
      istell:=0;
      ifunc:=1;
      ilanmax:=6;
      ilanmin:=1;
      iautocr:=1;
    end else begin
      if ((inummer > 9) and (inummer < 1000000)) then
      begin
        Ausgang:=True;
      end else begin
        Label2.Caption:='Anzahl Projekte muss zwischen 10 und 999999 liegen';
        Edit1.Color:=$0080FFFF;
        Edit1.Font.Color:=$00FF0000;
        Edit1.ReadOnly:=False;
        Edit1.Text:='';
        Edit1.AutoSelect:=True;
        ialpha:='';
        inummer:=0;
        iart:=2;
        inumkom:=0;
        ikomma:=0;
        izeich:=3;
        istell:=0;
        ifunc:=1;
        ilanmax:=6;
        ilanmin:=1;
        iautocr:=1;
      end;
    end;
  end;
end;

end.

