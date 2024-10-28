unit Unit1;

{*

Hersteller:
           LinSoft Jürgen linder (juelin) dg5uap@darc.de

Lizens:
           Opensource entsprechend

Hardware:
           keine zusätzliche Hardware

Software:
           Folgende Files
           Projektverwaltung.exe muß im Soucepfad vorhanden sein

wie Compelieren:
           Lazarus IDE mit FPC
           Printer4Lazarus muss im Objektinspektor als neue
           Anforderung hinzugefügt sein

wie Ausführen:
           Ausführen Projektverwaltung.exe (Icon)

*}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, Eingabe, LCLType, ExtCtrls, FileCtrl,
  PrintersDlgs, Printers, Math, unit2, Unit3;

type

  { TForm1 }

  TForm1 = class(TForm)
    FileListBox1: TFileListBox;
    Frame1: TFrame1;
    Frame2: TFrame2;
    Image1: TImage;
    Image2: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem16: TMenuItem;
    MenuItem17: TMenuItem;
    MenuItem18: TMenuItem;
    MenuItem19: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem20: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem9: TMenuItem;
    OpenDialog1: TOpenDialog;
    PrinterSetupDialog1: TPrinterSetupDialog;
    SaveDialog1: TSaveDialog;
    ScrollBar1: TScrollBar;
    ScrollBar2: TScrollBar;
    SelectDirectoryDialog1: TSelectDirectoryDialog;
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure MainMenu1DrawItem(Sender: TObject; ACanvas: TCanvas;
      ARect: TRect; AState: TOwnerDrawState);
    procedure MainMenu1MeasureItem(Sender: TObject; ACanvas: TCanvas;
      var AWidth, AHeight: Integer);
    procedure MenuItem4Click(Sender: TObject);
    procedure MenuItem5Click(Sender: TObject);
    procedure MenuItem6Click(Sender: TObject);
    procedure MenuItem7Click(Sender: TObject);
    procedure MenuItem8Click(Sender: TObject);
    procedure MenuItem9Click(Sender: TObject);
    procedure ScrollBar1Change(Sender: TObject);
    procedure ScrollBar2Change(Sender: TObject);
  private
  public
    function Nummer(Komma: Boolean; var Ntextstring: string): string;           // prüfen ob Ntextstring numerisch ist
                                                                                // wenn nicht wird Result leer zurück gegeben
                                                                                // Kamma=True wenn Komma erlaubt ist
                                                                                // Bei Komma wird aus '.' ein ',' gemacht
    function Blankweg(Art: integer; var Btextstring: string): string;           // Blanks in Btextstring entfernen und in Result zurück geben
                                                                                // Art 1 = vor dem Text
                                                                                // Art 2 = hinter dem Text
                                                                                // Art 3 = vor und hinter dem Text
    function Blankdazu(Pos, Lan: integer; var Htextstring: string): string;     // Blanks im Htextstring einfügen und in Result zurück geben
                                                                                // Lan = Ausgabelänge von Result
                                                                                // Pos 1 = vor dem Text
                                                                                // Pos 2 = hinter dem Text
    procedure AnzeigeLoe1;                                                      // Löschen Image1 der Bildschirmanzeige
    procedure AnzeigeLoe2;                                                      // Löschen Image2 original
    procedure AnzeigeLoe3;                                                      // Anzeigen Projekt in Label5
    procedure PersonenLoe;                                                      // Löschen Personentabelle und Zähler
    procedure ProjektLoe;                                                       // Löschen Projekttabelle und Zäehler
    procedure VorgangLoe;                                                       // Löschen Vorgangstabelle und Zähler
    procedure Scollbarinit;                                                     // setzen von Position und Maximum der horizontal und vertikal ScrollBar
    procedure Frame1Close;                                                      // Frame1 freigeben und ausblenden, Menü setzen
    procedure Frame2Close;                                                      // Frame2 freigeben und ausblenden, Menü setzen
    procedure Menuefunktionen;                                                  // je nach Stand der Dinge (Zähler Personen, Projekte und Vorgänge) werden Menüpunke an- bzw ausgeschaltet
    procedure Bildladen;                                                        // lädt Image1 aus Image2 je nach ScrollBar
    procedure Kopfperson;                                                       // Drucken Überschrift für Personen
  end;

type Tperson = record                                                           // Satzaufbau Personen
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

type Tprojekt = record                                                          // Satzaufbau Projekte
  Projektnummer: integer;
  Projektname: string;
  Beschreibung: array of string;
  Anzahl: integer;
  Status: string;                                                               // erfasst/gestartet/beendet
  Sollstarttermin: TDateTime;
  Sollendtermin: TDateTime;
  Iststarttermin: TDateTime;
  Istendtermin: TDateTime;
  Terminart: string;                                                            // Festtermin/Variabletermin
  Verantwortlich: integer;
  Satzart: Boolean;                                                             // aktiv/passiv
end;

type Tvorgang = record                                                          // Satzaufbau Vorgänge
  Projektnummer: integer;
  Vorgangnummer: integer;
  Kurzbeschreibung: string;
  Langbeschreibung: array of string;
  Anzahl: integer;
  Status: string;                                                               // erfasst/gestartet/beendet
  Sollstarttermin: TDateTime;
  Sollendtermin: TDateTime;
  Iststarttermin: TDateTime;
  Istendtermin: TDateTime;
  Terminart: string;                                                            // Festtermin/Variabletermin
  Verantwortlich: integer;
  Ausfuehrend: integer;
  Vorgaenger: integer;
  Nachfolger: integer;
  Satzart: Boolean;                                                             // aktiv/passiv
  vonX: integer;
  vonY: integer;
  bisX: integer;
  bisY: integer;
end;

const
  Personmax: integer = 9999;
  Vorgangmax: integer = 9999;
  BildrandX: integer = 10;
  BildrandY: integer = 10;
  VorgangrandX: integer = 10;
  VorgangrandY: integer = 10;
  StepX: integer = 780;
  StepY: integer = 330;

var
  Form1: TForm1;
  JaNein: word;                                                                 // Ergebnis Ja/Nein Abfrage
  Closestat: integer;                                                           // Merker ob Programm beendet werden darf. 0=Ja 1=Nein
  abbruch: Boolean;                                                             // Programm-Abbruch. Wenn True wird ohne Abfrage das Programm abgebrochen
  mtasts: integer;                                                              // Dürfen die Buttons betätigt werden. 1=Ja 0=Nein
  mlauf: integer;                                                               // Welche Eingabe ist zulässig (1..n) 0=keine Eingabe
  mart: integer;                                                                // Art der bearbeitung. 1=Anlegen, 2=Updaten, 3=Löschen
  Tag: integer;                                                                 // Datum Tag
  Monat: integer;                                                               // Datum Monat
  Jahr: integer;                                                                // Datum Jahr
  BUser: string;                                                                // Username der angemeldet ist
  FSatz: string;                                                                // Hilfsvariable
  sta: integer;                                                                 // Status für Initialisierung der Variablen bei Programmstart. 1=vor Init, 2=nach Init
  vers: string;                                                                 // Versionsnummer des Programms
  drucker: string;                                                              // Druckername
  druckn: integer;                                                              // Druckernummer für Tabelle
  px: integer;                                                                  // X aktuell in Pixel
  py: integer;                                                                  // Y aktuell in Pixel
  vpx: integer;                                                                 // X-Differenz in Pixel
  vpy: integer;                                                                 // Y-Differenz in Pixel
  format: TPrinterOrientation;                                                  // Druckerformat
  Seite: integer;                                                               // Seitenanzahl Frucker
  Zeile: integer;                                                               // Zeilenanzahl Drucker
  Fileadresse: TextFile;                                                        // Opennummer eines Textfiles (Assign)
  Updperson: Boolean;                                                           // Personentabelle wurde geändert
  Updprojekt: Boolean;                                                          // Projekttabelle wurde geändert
  Updvorgang: Boolean;                                                          // Vorgangstabelle wurde geändert
  Persontab: array of Tperson;                                                  // Tabelle Personen
  Personzaehl: integer;                                                         // Zähler Tabelle Personen
  Projekt: Tprojekt;                                                            // Daten Projekt
  Projektzaehl: integer;                                                        // Zähler Tabelle Vorgänge
  Vorgangtab: array of Tvorgang;                                                // Tabelle Vorgänge
  Vorgangzaehl: integer;                                                        // Zähler Tabelle Vorgänge
  Indexpfad: string;                                                            // Pfad des Projektindexfile
  Personpfad: string;                                                           // Pfad der Personenfiles
  Projektpfad: string;                                                          // Pfad der Projektfiles
  Projektfile: string;                                                          // Filename des Projektfile
  Projektname: string;                                                          // Projektname
  Projektnummer: integer;                                                       // Projektnummer
  Projektstatus: string;                                                          // Projektstatus (E=erfasst G=gestartet B=beendet)
  Projektstart: TDateTime;                                                      // Projekt-Start Termin (Datum und Zeit)
  Projektende: TDateTime;                                                       // Projekt-Ende Termin (Datum und Zeit)
  GraphikstartX: integer;                                                       // Startposition X-Pos von Image1 in Image2
  GraphikstartY: integer;                                                       // Startposition Y-Pos von Image1 in Image2
  ScrollposaltX: integer;                                                       // Position der ScrollBar X
  ScrollposaltY: integer;                                                       // Position der ScrollBar Y

implementation

{$R *.lfm}

{ TForm1 }

function TForm1.Nummer(Komma: Boolean; var Ntextstring: string): string;
  var laenge: integer;
  var stelle: integer;
  var m: integer;
  var zeichen: string;
  var h1: string;
  var h2: string;
  var tt: string;
begin
  laenge:=Length(Ntextstring);
  tt:='';
  if laenge > 0 then
  begin
    Ntextstring:=Blankweg(3, Ntextstring);
    laenge:=Length(Ntextstring);
    if laenge > 0 then
    begin
      tt:=ialpha;
      for stelle:=1 to laenge do
      begin
        zeichen:=Copy(Ntextstring,stelle,1);
        m:=0;
        if zeichen = '0' then
        begin
          m:=1;
        end;
        if zeichen = '1' then
        begin
          m:=1;
        end;
        if zeichen = '2' then
        begin
          m:=1;
        end;
        if zeichen = '3' then
        begin
          m:=1;
        end;
        if zeichen = '4' then
        begin
          m:=1;
        end;
        if zeichen = '5' then
        begin
          m:=1;
        end;
        if zeichen = '6' then
        begin
          m:=1;
        end;
        if zeichen = '7' then
        begin
          m:=1;
        end;
        if zeichen = '8' then
        begin
          m:=1;
        end;
        if zeichen = '9' then
        begin
          m:=1;
        end;
        if ((komma) and (zeichen = '.')) then
        begin
          m:=1;
          h1:='';
          h2:='';
          if stelle > 1 then h1:=Copy(Ntextstring,1,stelle-1);
          if ((stelle > 1) and (stelle < laenge)) then h2:=Copy(Ntextstring,stelle+1,laenge-stelle);
          tt:=h1+','+h2;
        end;
        if ((komma) and (zeichen = ',')) then
        begin
          m:=1;
        end;
        if m = 0 then
        begin
          tt:='';
        end;
      end;
    end;
  end;
  Result:=tt;
end;

procedure TForm1.Kopfperson;
  var h1: string;
  var h4: string;
begin
  if seite > 0 then
  begin
    Printers.Printer.NewPage;
  end;
  seite:=seite+1;
  py:=vpy;
  zeile:=0;
  h1:=IntToStr(seite);
  h4:=Blankdazu(1, 4, h1);
  Printers.Printer.Canvas.TextOut(px, py, '    Personentabelle für Projejte                                                                                          Seite: '+h4+'          Datum: '+DateToStr(now));
  py:=py+vpy;
  zeile:=zeile+1;
  Printers.Printer.Canvas.TextOut(px, py, 'Nummer Personenart Vorname                   Nachname                  Strasse                         Hausnummer Firmenname');
  py:=py+vpy;
  zeile:=zeile+1;
  Printers.Printer.Canvas.TextOut(px, py, 'PLZ        Ort                            Telefon              Handy                EMail                          Land');
  py:=py+vpy;
  zeile:=zeile+1;
  Printers.Printer.Canvas.TextOut(px, py, 'Partner: Anrede Vorname                   Nachname                  Telefon');
  //                                                1234   123456789.123456789.12345 123456789.123456789.12345 12345
  //                                                    123
  py:=py+vpy;
  zeile:=zeile+1;
  Printers.Printer.Canvas.TextOut(px, py, '====================================================================================================================================================================');
  py:=py+vpy;
  zeile:=zeile+1;
end;
// 123456789.123456789.123456789.123456789.123456789.123456789.123456789.123456789.123456789.123456789.123456789.123456789.123456789.123456789.123456789.123456789.1234
//     Personentabelle für Projejte                                                                                          Seite: 1234          Datum: 12.12.1234
// Nummer Personenart Vorname                   Nachname                  Strasse                        Hausnummer Firmenname
//   1234    123456   123456789.123456789.12345 123456789.123456789.12345 123456789.123456789.123456789. 12345      123456789.123456789.123456789.
// PLZ        Ort                            Telefon              Handy                EMail                          Land
// 123456789. 123456789.123456789.123456789. 123456789.123456789. 123456789.123456789. 123456789.123456789.123456789. 123456789.123456789.123456789.
// wenn Firma Partner: Anrede Vorname                   Nachname                  Telefon
//                     1234   123456789.123456789.12345 123456789.123456789.12345 123456789.123456789.
// ====================================================================================================================================================================

procedure TForm1.Bildladen;
  var h: integer;
  var w: integer;
  var x: integer;
  var y: integer;
begin
  AnzeigeLoe1;
  GraphikstartX:=0;
  GraphikstartY:=0;
  if ScrollBar1.Position > 1 then
  begin
    GraphikstartX:=GraphikstartX+((ScrollBar1.Position-1)*StepX);
  end;
  if ScrollBar2.Position > 1 then
  begin
    GraphikstartY:=GraphikstartY+((ScrollBar2.Position-1)*StepY);
  end;
  x:=GraphikstartX+Image1.Width;
  y:=GraphikstartY+Image1.Height;
  if x > Image2.Width then x:=Image2.Width;
  if y > Image2.Height then y:=Image2.Height;
  w:=x-GraphikstartX;
  h:=y-GraphikstartY;
  Image1.Canvas.CopyRect(Rect(0,0,w,h),Image2.Canvas,Rect(GraphikstartX,GraphikstartY,x,y));
end;

procedure TForm1.Menuefunktionen;
begin
  MenuItem1.Visible:=True;
  MenuItem5.Visible:=True;
  MenuItem7.Visible:=True;
  MenuItem6.Visible:=True;
  MenuItem8.Visible:=True;
  MenuItem9.Visible:=False;
  if Personzaehl > 0 then
  begin
    MenuItem9.Visible:=True;
  end;
  MenuItem2.Visible:=False;
  MenuItem10.Visible:=False;
  MenuItem11.Visible:=False;
  MenuItem12.Visible:=False;
  MenuItem13.Visible:=False;
  MenuItem14.Visible:=False;
  MenuItem15.Visible:=False;
  if Personzaehl > 0 then
  begin
    MenuItem2.Visible:=True;
    MenuItem10.Visible:=True;
    MenuItem11.Visible:=True;
    MenuItem12.Visible:=False;
    MenuItem13.Visible:=False;
    MenuItem14.Visible:=False;
    MenuItem15.Visible:=False;
    if Projektzaehl > 0 then
    begin
      MenuItem12.Visible:=True;
      MenuItem13.Visible:=True;
      MenuItem14.Visible:=True;
      MenuItem15.Visible:=True;
    end;
  end;
  if Projektzaehl > 0 then
  begin
    MenuItem3.Visible:=True;
    MenuItem16.Visible:=False;
    MenuItem17.Visible:=False;
    MenuItem18.Visible:=True;
    MenuItem19.Visible:=False;
    MenuItem20.Visible:=False;
    if Vorgangzaehl > 0 then
    begin
      MenuItem16.Visible:=True;
      MenuItem17.Visible:=True;
      MenuItem19.Visible:=True;
      MenuItem20.Visible:=True;
    end;
  end;
  MenuItem4.Visible:=True;
end;

procedure TForm1.Frame1Close;
begin
  Frame1.Visible:=False;
  TFrame1(FindComponent('Frame1')).Free;
  Menuefunktionen;
end;

procedure TForm1.Frame2Close;
begin
  Frame2.Visible:=False;
  TFrame2(FindComponent('Frame2')).Free;
  Menuefunktionen;
end;

procedure TForm1.Scollbarinit;
  var h1: integer;
  var h2: integer;
begin
  ScrollBar1.Max:=1;
  ScrollBar1.Position:=1;
  h1:=Image2.Width;
  if h1 > 1565 then
  begin
    h2:=Ceil(h1 / StepX);
    ScrollBar1.Max:=h2;
  end;
  ScrollBar2.Max:=1;
  ScrollBar2.Position:=1;
  h1:=Image2.Height;
  if h1 > 662 then
  begin
    h2:=Ceil(h1 / StepY);
    ScrollBar2.Max:=h2;
  end;
end;

procedure TForm1.PersonenLoe;
  var h1: integer;
begin
  for h1:=0 to Personmax do
  begin
    Persontab[h1].Nummer:=h1;
    if h1 = 0 then
    begin
      Persontab[h1].Art:='I';
    end else begin
      Persontab[h1].Art:=' ';
    end;
    Persontab[h1].Satzart:=False;
  end;
  personzaehl:=0;
end;

procedure TForm1.ProjektLoe;
begin
  Projekt.Projektnummer:=0;
  Projekt.Projektname:='';
  Projekt.Anzahl:=0;
  Projekt.Status:=' ';
  Projekt.Sollstarttermin:=0;
  Projekt.Sollendtermin:=0;
  Projekt.Iststarttermin:=0;
  Projekt.Istendtermin:=0;
  Projekt.Terminart:='V';
  projekt.Verantwortlich:=0;
  projekt.Satzart:=False;
  Projektzaehl:=0;                                                        // Zähler Tabelle Vorgänge
end;

procedure TForm1.VorgangLoe;
  var h1: integer;
begin
  for h1:=0 to Vorgangmax do
  begin
    Vorgangtab[h1].Projektnummer:=0;
    Vorgangtab[h1].Vorgangnummer:=h1;
    Vorgangtab[h1].Kurzbeschreibung:='';
    SetLength(Vorgangtab[h1].Langbeschreibung, 0);
    Vorgangtab[h1].Anzahl:=0;
    Vorgangtab[h1].Status:=' ';
    Vorgangtab[h1].Sollstarttermin:=0;
    Vorgangtab[h1].Sollendtermin:=0;
    Vorgangtab[h1].Iststarttermin:=0;
    Vorgangtab[h1].Istendtermin:=0;
    Vorgangtab[h1].Terminart:='V';
    Vorgangtab[h1].Verantwortlich:=0;
    Vorgangtab[h1].Ausfuehrend:=0;
    Vorgangtab[h1].Vorgaenger:=0;
    Vorgangtab[h1].Nachfolger:=0;
    Vorgangtab[h1].Satzart:=False;
    Vorgangtab[h1].vonX:=0;
    Vorgangtab[h1].vonY:=0;
    Vorgangtab[h1].bisX:=0;
    Vorgangtab[h1].bisY:=0;
  end;
  vorgangzaehl:=0;
end;

procedure TForm1.AnzeigeLoe1;
begin
  image1.Canvas.Pen.Color:=clBlack;
  image1.Canvas.Pen.Style:=psSolid;
  image1.Canvas.Pen.Width:=1;
  Image1.Canvas.Brush.Color:=clBlack;
  Image1.Canvas.Brush.Style:=bsSolid;
  Image1.Canvas.FillRect(Rect(0,0,Image1.Width,Image1.Height));
end;

procedure TForm1.AnzeigeLoe2;
begin
  image2.Canvas.Pen.Color:=clBlack;
  image2.Canvas.Pen.Style:=psSolid;
  image2.Canvas.Pen.Width:=1;
  Image2.Canvas.Brush.Color:=clWhite;
  Image2.Canvas.Brush.Style:=bsSolid;
  Image2.Canvas.FillRect(Rect(0,0,Image2.Width,Image2.Height));
end;

procedure TForm1.AnzeigeLoe3;
  var h1: string;
  var h2: string;
  var h3: integer;
  var h4: string;
  var h5: string;
  var h6: TDateTime;
  var h7: string;
  var h8: TDate;
  var h9: string;
  var ha: integer;
  var hb: integer;
  var hc: string;
begin
  h1:='';
  if Projektnummer > 0 then
  begin
    h1:=IntToStr(Projektnummer);
    h3:=Length(h1);
    if h3 < 4 then
    begin
      while h3 < 4 do
      begin
        h1:='0'+h1;
        h3:=Length(h1);
      end;
    end;
    h2:='';
    if Projektstatus = 'E' then h2:='rrfasst';
    if Projektstatus = 'G' then h2:='gestartet';
    if Projektstatus = 'B' then h2:='beendet';
    h4:='';
    h5:='';
    h7:='';
    if Projektstart > 0 then h4:=DateTimeToStr(Projektstart);
    if Projektende > 0 then h5:=DateTimeToStr(Projektende);
    if ((Projektstart > 0) and (Projektende > 0)) then
    begin
      if Projektende > Projektstart then
      begin
        h6:=Projektende-Projektstart;
        h7:=DateToStr(h6);
        h8:=StrToDate(h7);
        h7:='Projektdauer: ';
        if h8 > 0 then
        begin
          ha:=Trunc(h8);
          h7:=IntToStr(ha)+' Tage ';
        end;
        h9:=TimeToStr(h6);
        hc:=Copy(h9,1,2);
        ha:=StrToInt(hc);
        hc:=Copy(h9,4,2);
        hb:=StrToInt(hc);
        h7:=h7+IntToStr(ha)+' Stunden '+IntToStr(hb)+' Minuten';
      end;
    end;
    Label5.Caption:='Projekt: '+h1+' '+Projektname+'  Status: '+h2+'  Starttermin: '+h4+'  Endtermin: '+h5+'  '+h7;
  end else begin
    Label5.Caption:='';
  end;
end;

function TForm1.Blankweg(Art: integer; var Btextstring: string): string;
  var laenge: integer;
  var stelle: integer;
  var vari12: integer;
  var zeichen: string;
  var ausgabe: string;
begin
  laenge:=Length(Btextstring);
  ausgabe:='';
  if laenge > 0 then
  begin
    ausgabe:=Btextstring;
    if ((Art = 1) or (Art = 3)) then
    begin
      vari12:=0;
      for stelle:=1 to laenge do
      begin
        zeichen:=Copy(Btextstring,stelle,1);
        if (vari12 = 0) then
        begin
          if (zeichen <> ' ') then
          begin
            vari12:=stelle;
          end;
        end;
      end;
      if vari12 > 0 then
      begin
        zeichen:=Btextstring;
        ausgabe:=Copy(zeichen,vari12,laenge-(vari12-1));
      end else begin
        ausgabe:='';
      end;
    end;
    laenge:=Length(ausgabe);
    if laenge > 0 then
    begin
      if ((Art = 2) or (Art = 3)) then
      begin
        vari12:=0;
        for stelle:=laenge downto 1 do
        begin
          zeichen:=Copy(ausgabe,stelle,1);
          if (vari12 = 0) then
          begin
            if (zeichen <> ' ') then
            begin
              vari12:=stelle;
            end;
          end;
        end;
        if (vari12 > 0) then
        begin
          zeichen:=ausgabe;
          ausgabe:=Copy(zeichen,1,vari12);
        end else begin
          ausgabe:='';
        end;
      end;
    end else begin
      ausgabe:='';
    end;
  end;
  Result:=ausgabe;
end;

function TForm1.Blankdazu(Pos, Lan: integer; var Htextstring: string): string;
  var laenge: integer;
  var stelle: integer;
  var ausgabe: string;
begin
  laenge:=Length(Htextstring);
  ausgabe:=Htextstring;
  if laenge < Lan then
  begin
    while laenge < Lan do
    begin
      if Pos = 1 then
      begin
        ausgabe:=' '+ausgabe;
      end else begin
        ausgabe:=ausgabe+' ';
      end;
      laenge:=Length(ausgabe);
    end;
  end else begin
    if laenge > Lan then
    begin
      if Pos = 1 then
      begin
        stelle:=laenge-Lan+1;
        ausgabe:=Copy(Htextstring,stelle,Lan);
      end else begin
        ausgabe:=Copy(Htextstring,1,Lan);
      end;
    end;
  end;
  Result:=ausgabe;
end;

procedure TForm1.MainMenu1DrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; AState: TOwnerDrawState);
  var h1: string;
begin
  AState:=AState;
  ACanvas.Font.Name := 'Courier';
  ACanvas.Font.Size := 16;
  ACanvas.Font.Style := [fsBold];
  ACanvas.Font.Color := $008000FF;
  ACanvas.Brush.Color := clBtnShadow;
  ACanvas.Rectangle(ARect);
  ACanvas.FillRect(ARect.Left,ARect.Top,ARect.Right,ARect.Bottom);
  h1:=(Sender as TMenuItem).Caption;
  ACanvas.TextOut(ARect.Left+10, ARect.Top, h1);
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  var h1: string;
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  if (abbruch) then
  begin
    CanClose:=true;
  end else begin
    Form1.Cursor:=crDefault;
    Form1.Refresh;
    CanClose:=false;
    if (Closestat = 0) then
    begin
      Closestat:=1;
      h1:='';
      if Updperson then h1:='Personen wurden noch nicht gespeichert';
      if h1 = '' then
      begin
        if Updprojekt then h1:=h1+'Projekt wurde noch nicht gespeichert';
      end else begin
        if Updprojekt then h1:=h1+'  Projekt wurde noch nicht gespeichert';
      end;
      if h1 = '' then
      begin
        if Updvorgang then h1:=h1+'Vorgänge wurden noch nicht gespeichert';
      end else begin
        if Updvorgang then h1:=h1+'  Vorgänge wurden noch nicht gespeichert';
      end;
      if h1 = '' then
      begin
        h1:=h1+'Programm wirklich beenden ?';
      end else begin
        h1:=h1+'  Programm wirklich beenden ?';
      end;
      JaNein:=messagedlg(h1, mtConfirmation, [mbYes, mbNo], 0);
      Closestat:=0;
      if (JaNein = mrYes) then
      begin
        CanClose:=true;
      end else begin
        mtasts:=1;
      end;
    end;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
  var heute: string;
  var h1: integer;
  var h3: string;
  var Rec: LongRec;
  var hx1: integer;
  var hx2: integer;
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  Label2.Caption:='';
  Closestat:=0;
  mtasts:=1;
  ialpha:='';
  inummer:=0;
  inumkom:=0;
  ikomma:=0;
  iart:=1;
  izeich:=3;
  istell:=0;
  mlauf:=0;
  BUser:='                                                     ';
  try
    BUser:=GetEnvironmentVariable('USERNAME');
  except
    BUser:='unknown';
  end;
  heute:=FormatDateTime('DD.MM.YYYY',now);
  h1:=StrLen(PChar(heute));
  if (h1 = 10) then
  begin
    h3:=Copy(heute, 1, 2);
    Tag:=StrToInt(h3);
    h3:=Copy(heute, 4, 2);
    Monat:=StrToInt(h3);
    h3:=Copy(heute, 7, 4);
    Jahr:=StrToInt(h3);
  end;
  abbruch:=false;
  sta:=1;
  Rec:=LongRec(GetFileVersion(ParamStr(0)));
  hx1:=Rec.Bytes[2];
  hx2:=Rec.Bytes[2];
  vers:=IntToStr(hx1)+'.'+IntToStr(hx2);
  Form1.Caption:='  '+Application.Title+'            Version '+vers+'                   <'+BUser+'>';
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
end;

procedure TForm1.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  Form1.Cursor:=crDefault;
  Form1.Refresh;
  CloseAction:=caFree;
end;

procedure TForm1.FormActivate(Sender: TObject);
  var h1: integer;
  var h2: integer;
  var h3: integer;
  var h4: Boolean;
  var h5: string;
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  Label2.Caption:='';
  if sta = 1 then
  begin
    Form1.Caption:='  '+Application.Title+'            Version '+vers+'                   <'+BUser+'>';
    Closestat:=0;
    mtasts:=1;
    ialpha:='';
    inummer:=0;
    inumkom:=0;
    ikomma:=0;
    iart:=1;
    izeich:=3;
    istell:=0;
    mlauf:=0;
    sta:=2;
    Updperson:=False;
    Updprojekt:=False;
    Updvorgang:=False;
    Indexpfad:='';
    Personpfad:='';
    Projektpfad:='';
    Projektfile:='';
    Projektname:='';
    Projektnummer:=0;
    Projektstatus:=' ';
    Projektstart:=0;
    Projektende:=0;
    GraphikstartX:=0;
    GraphikstartY:=0;
    ScrollposaltX:=1;
    ScrollposaltY:=1;
    SetLength(Persontab, Personmax+1);
    SetLength(Vorgangtab, Vorgangmax+1);
    PersonenLoe;
    ProjektLoe;
    VorgangLoe;
    AnzeigeLoe1;
    AnzeigeLoe2;
    AnzeigeLoe3;
    Scollbarinit;
    Bildladen;
    FileListBox1.Visible:=False;
    Image2.Visible:=False;
    MenuItem1.Visible:=False;
    MenuItem5.Visible:=False;
    MenuItem7.Visible:=False;
    MenuItem6.Visible:=False;
    MenuItem8.Visible:=False;
    MenuItem9.Visible:=False;
    MenuItem2.Visible:=False;
    MenuItem10.Visible:=False;
    MenuItem11.Visible:=False;
    MenuItem12.Visible:=False;
    MenuItem13.Visible:=False;
    MenuItem14.Visible:=False;
    MenuItem15.Visible:=False;
    MenuItem3.Visible:=False;
    MenuItem16.Visible:=False;
    MenuItem17.Visible:=False;
    MenuItem18.Visible:=False;
    MenuItem19.Visible:=False;
    MenuItem20.Visible:=False;
    MenuItem4.Visible:=True;
    h4:=SelectDirectoryDialog1.Execute;
    if h4 then
    begin
      Indexpfad:=SelectDirectoryDialog1.FileName+'\';
      h5:=Indexpfad+'Projekt.index';
      if not FileExists(h5) then
      begin
        MenuItem4.Visible:=False;
        Frame1:=TFrame1.Create(self);
        Frame1.parent:=self;
        Frame1.Align:=alClient;
        Frame1.Visible:=True;
        Frame1.Top:=0;
        Frame1.Left:=0;
        Frame1.Height:=864;
        Frame1.Width:=1595;
        Frame1.Name:='Frame1';
        Form1.ActiveControl:=Frame1.Edit1;
      end else begin
        Menuefunktionen;
        MenuItem6.Visible:=False;
      end;
    end else begin
      JaNein:=messagedlg('Verzeichnis für Projektindexfile nicht ausgewählt, Programm-Abbruch', mtInformation, [mbOk], 0);
      abbruch:=True;
      close;
    end;
    vpx:=56;
    vpy:=68;
    h1:=Printers.Printer.Printers.Count;
    h2:=Printers.Printer.PrinterIndex;
    if h1 > 0 then
    begin
      for h3:=0 to h1-1 do
      begin
        if h3 = h2 then
        begin
          drucker:=Printers.Printer.Printers.Strings[h3];
          druckn:=h3;
        end;
      end;
    end;
  end;
end;

procedure TForm1.MainMenu1MeasureItem(Sender: TObject; ACanvas: TCanvas;
  var AWidth, AHeight: Integer);
begin
  ACanvas:=ACanvas;
  AHeight:=32;
  AWidth:=383;
end;

procedure TForm1.MenuItem4Click(Sender: TObject);
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  Label2.Caption:='';
  if (Closestat = 0) then
  begin
    if (mtasts = 1) then
    begin
      mtasts:=0;
      mlauf:=0;
      close;
    end;
  end;
end;

procedure TForm1.MenuItem5Click(Sender: TObject);
  var h1: Boolean;
  var h2: string;
  var h3: string;
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  Label2.Caption:='';
  if (mtasts = 1) then
  begin
    h1:=True;
    if Updperson then
    begin
      h1:=False;
      JaNein:=messagedlg('Personen wurden noch nicht gespeichert, trotzdem neues Personenfile anlegen?', mtConfirmation, [mbYes, mbNo], 0);
      if (JaNein = mrYes) then
      begin
        h1:=True;
      end;
    end;
    if h1 then
    begin
      mtasts:=0;
      h1:=SelectDirectoryDialog1.Execute;
      if h1 then
      begin
        Personpfad:=SelectDirectoryDialog1.FileName+'\';
        Updperson:=False;
        PersonenLoe;
        Menuefunktionen;
        mtasts:=1;
      end else begin
        mtasts:=1;
      end;
    end;
  end;
end;

procedure TForm1.MenuItem6Click(Sender: TObject);
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  Label2.Caption:='';
  if (mtasts = 1) then
  begin
    Frame2:=TFrame2.Create(self);
    Frame2.parent:=self;
    Frame2.Align:=alClient;
    Frame2.Visible:=True;
    Frame2.Top:=0;
    Frame2.Left:=0;
    Frame2.Height:=864;
    Frame2.Width:=1595;
    Frame2.Name:='Frame1';
    Frame2.Timer1.Enabled:=True;
  end;
end;

procedure TForm1.MenuItem7Click(Sender: TObject);
  var h1: Boolean;
  var h2: integer;
  var h3: integer;
  var h4: integer;
  var h5: string;
  var h6: string;
  var h7: string;
  var h8: integer;
  var hb: integer;
  var hc: integer;
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  Label2.Caption:='';
  hb:=0;
  if (mtasts = 1) then
  begin
    h1:=True;
    if Updperson then
    begin
      h1:=False;
      JaNein:=messagedlg('Personen wurden noch nicht gespeichert, trotzdem Personenfile laden?', mtConfirmation, [mbYes, mbNo], 0);
      if (JaNein = mrYes) then
      begin
        h1:=True;
      end;
    end;
    if h1 then
    begin
      mtasts:=0;
      OpenDialog1.DefaultExt:='mita';
      OpenDialog1.Filter:='Personen|*.mita';
      OpenDialog1.Title:='Personen';
      h1:=OpenDialog1.Execute;
      if h1 then
      begin
        Updperson:=False;
        PersonenLoe;
        h2:=Length(OpenDialog1.FileName);
        if h2 > 5 then
        begin
          h3:=0;
          for h4:=1 to h2 do
          begin
            h5:=Copy(OpenDialog1.FileName,h4,1);
            if h5 = '\' then h3:=h4;
          end;
          if h3 > 0 then
          begin
            if h3 < h2 then
            begin
              Personpfad:=Copy(OpenDialog1.FileName,1,h3);
              if FileExists(OpenDialog1.FileName) then
              begin
                assignFile(Fileadresse, OpenDialog1.FileName);
                Reset(Fileadresse);
                While Not EOF(Fileadresse) do
                begin
                  readLn(Fileadresse, h6);
                  h5:=Blankweg(1, h6);
                  h2:=Length(h5);
                  if h2 > 0 then
                  begin
                    h6:='';
                    h4:=0;
                    h8:=0;
                    for h3:=1 to h2 do
                    begin
                      if h4 = 0 then
                      begin
                        h7:=Copy(h5,h3,1);
                        if h7 = ';' then
                        begin
                          h4:=h3;
                        end else begin
                          h6:=h6+h7;
                        end;
                      end;
                      if h4 > 0 then
                      begin
                        hc:=Length(h6);
                        if h8 = 16 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Partnernachname:=ialpha;
                            h8:=17;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 15 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Partnervorname:=ialpha;
                            h8:=16;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 14 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Partnertelefon:=ialpha;
                            h8:=15;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 13 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Partneranrede:=ialpha;
                            h8:=14;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 12 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Firmenname:=ialpha;
                            h8:=13;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 11 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Land:=ialpha;
                            h8:=12;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 10 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Email:=ialpha;
                            h8:=11;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 9 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Handy:=ialpha;
                            h8:=10;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 8 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Telefon:=ialpha;
                            h8:=9;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 7 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Ort:=ialpha;
                            h8:=8;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 6 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Plz:=ialpha;
                            h8:=7;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 5 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Hausnummer:=ialpha;
                            h8:=6;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 4 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Strasse:=ialpha;
                            h8:=5;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 3 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Nachname:=ialpha;
                            h8:=4;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 2 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Vorname:=ialpha;
                            h8:=3;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 1 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            Persontab[hb].Art:=ialpha;
                            h8:=2;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                        if h8 = 0 then
                        begin
                          if h4 < h2 then
                          begin
                            ialpha:=Blankweg(3, h6);
                            try
                              h8:=StrToInt(ialpha);
                            except
                              h8:=0;
                            end;
                            if ((h8 > 0) and (h8 <= Personmax)) then
                            begin
                              Persontab[h8].Nummer:=h8;
                              Persontab[h8].Satzart:=True;
                              personzaehl:=personzaehl+1;
                              hb:=h8;
                            end;
                            h8:=1;
                            h6:='';
                            h4:=0;
                          end;
                        end;
                      end;
                    end;
                    if h8 = 16 then
                    begin
                      if h4 < h2 then
                      begin
                        ialpha:=Blankweg(3, h6);
                        Persontab[hb].Partnernachname:=ialpha;
                        h8:=17;
                        h6:='';
                        h4:=0;
                      end;
                    end;
                  end;
                end;
                CloseFile(Fileadresse);
              end else begin
                Label2.Caption:='File '+OpenDialog1.FileName+' nicht vorhanden, Abbruch';
              end;
            end else begin
              Label2.Caption:='Filename nicht zulässig, Abbruch';
            end;
          end else begin
            Label2.Caption:='Filename nicht zulässig, Abbruch';
          end;
        end else begin
          Label2.Caption:='Filename nicht zulässig, Abbruch';
        end;
        Menuefunktionen;
      end;
      mtasts:=1;
    end;
  end;
end;

procedure TForm1.MenuItem8Click(Sender: TObject);
  var h1: Boolean;
  var h2: string;
  var h3: string;
  var h4: integer;
  var h5: string;
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  Label2.Caption:='';
  if (mtasts = 1) then
  begin
    JaNein:=messagedlg('Person/Firma speichern (alte Daten werden überschrieben)?', mtConfirmation, [mbYes, mbNo], 0);
    if (JaNein = mrYes) then
    begin
      h1:=True;
      if Personpfad = '' then
      begin
        h1:=SelectDirectoryDialog1.Execute;
        if h1 then
        begin
          Personpfad:=SelectDirectoryDialog1.FileName+'\';
        end;
      end;
      if h1 then
      begin
        h2:=Personpfad+'Personen.mita';
        if FileExists(h2) then
        begin
          DeleteFile(h2);
        end;
        h3:=Personpfad+'Personen.mita';
        assignFile(Fileadresse, h3);
        Rewrite(Fileadresse);
        for h4:=1 to Personmax do
        begin
          if Persontab[h4].Satzart then
          begin
            h5:=IntToStr(Persontab[h4].Nummer)+';'+Persontab[h4].Art+';'+Persontab[h4].Vorname+';'+Persontab[h4].Nachname+';';
            h5:=h5+Persontab[h4].Strasse+';'+Persontab[h4].Hausnummer+';'+Persontab[h4].Plz+';'+Persontab[h4].Ort+';';
            h5:=h5+Persontab[h4].Telefon+';'+Persontab[h4].Handy+';'+Persontab[h4].Email+';'+Persontab[h4].Land+';';
            h5:=h5+Persontab[h4].Firmenname+';'+Persontab[h4].Partneranrede+';'+Persontab[h4].Partnertelefon+';';
            h5:=h5+Persontab[h4].Partnervorname+';'+Persontab[h4].Partnernachname;
            writeln(Fileadresse, h5);
          end;
        end;
        CloseFile(Fileadresse);
        Updperson:=False;
      end;
    end;
    mtasts:=1;
  end;
end;

procedure TForm1.MenuItem9Click(Sender: TObject);
  var h1: Boolean;
  var h2: string;
  var h3: integer;
  var h4: string;
  var h5: string;
  var h6: string;
  var h7: string;
  var h8: string;
  var h9: string;
  var ha: string;
  var hb: string;
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  Label2.Caption:='';
  if (mtasts = 1) then
  begin
    mtasts:=0;
    format:=poLandscape;
    Printers.Printer.Orientation:=format;
    h1:=PrinterSetupDialog1.Execute;
    if h1 then
    begin
      drucker:=Printers.Printer.Printers.Strings[Printer.PrinterIndex];
      druckn:=Printer.PrinterIndex;
      Screen.Cursor:=crHourGlass;
      Form1.Cursor:=crHourGlass;
      Form1.Refresh;
      Printers.Printer.PrinterIndex:=druckn;
      Printers.Printer.Canvas.Brush.Style:=bsSolid;
      Printers.Printer.Canvas.Font.Charset:=0;
      Printers.Printer.Canvas.Font.Name:='Courier';
      Printers.Printer.Canvas.Font.Style:=[fsBold];
      Printers.Printer.Canvas.Font.Size:=8;
      Printers.Printer.Canvas.Font.Pitch:=fpFixed;
      Printers.Printer.Canvas.Font.Color:=clBlack;
      Printers.Printer.Copies:=1;
      px:=vpx+vpx;
      py:=vpy;
      seite:=0;
      zeile:=99;
      Printers.Printer.BeginDoc;
        if zeile > 64 then Kopfperson;
        for h3:=1 to Personmax do
        begin
          if Persontab[h3].Satzart then
          begin
            if zeile > 64 then Kopfperson;
            h2:=IntToStr(Persontab[h3].Nummer);
            h4:=Blankdazu(1, 4, h2);
            h2:=Persontab[h3].Art;
            h5:=Blankdazu(2, 6, h2);
            h2:=Persontab[h3].Vorname;
            h6:=Blankdazu(2, 25, h2);
            hb:=' ';
            if Persontab[h3].Art = 'Person' then hb:=hb+' ';
            h2:=Persontab[h3].Nachname;
            h7:=Blankdazu(2, 25, h2);
            h2:=Persontab[h3].Strasse;
            h8:=Blankdazu(2, 30, h2);
            h2:=Persontab[h3].Hausnummer;
            h9:=Blankdazu(2, 5, h2);
            h2:=Persontab[h3].Firmenname;
            ha:=Blankdazu(2, 30, h2);
            Printers.Printer.Canvas.TextOut(px, py, '  '+h4+'    '+h5+'   '+h6+hb+h7+' '+h8+'  '+h9+'      '+ha);
            py:=py+vpy;
            zeile:=zeile+1;
            h2:=Persontab[h3].Plz;
            h4:=Blankdazu(2, 10, h2);
            h2:=Persontab[h3].Ort;
            h5:=Blankdazu(2, 30, h2);
            h2:=Persontab[h3].Telefon;
            h6:=Blankdazu(2, 20, h2);
            h2:=Persontab[h3].Handy;
            h7:=Blankdazu(2, 20, h2);
            h2:=Persontab[h3].Email;
            h8:=Blankdazu(2, 30, h2);
            h2:=Persontab[h3].Land;
            h9:=Blankdazu(2, 30, h2);
            Printers.Printer.Canvas.TextOut(px, py, h4+' '+h5+' '+h6+' '+h7+' '+h8+' '+h9);
            py:=py+vpy;
            zeile:=zeile+1;
            if Persontab[h3].Art = 'Firma' then
            begin
              if ((Persontab[h3].Partneranrede <> '') and (Persontab[h3].Partneranrede <> 'ohne')) then
              begin
                h2:=Persontab[h3].Partneranrede;
                h4:=Blankdazu(2, 5, h2);
                h2:=Persontab[h3].Partnervorname;
                h5:=Blankdazu(2, 25, h2);
                h2:=Persontab[h3].Partnernachname;
                h6:=Blankdazu(2, 25, h2);
                h2:=Persontab[h3].Partnertelefon;
                h7:=Blankdazu(2, 20, h2);
                Printers.Printer.Canvas.TextOut(px, py, 'Partner: '+h4+'   '+h5+' '+h6+' '+h7);
                py:=py+vpy;
                zeile:=zeile+1;
              end;
            end;
            Printers.Printer.Canvas.TextOut(px, py, '--------------------------------------------------------------------------------------------------------------------------------------------------------------------');
            py:=py+vpy;
            zeile:=zeile+1;
          end;
        end;
      Printers.Printer.EndDoc;
      Screen.Cursor:=crDefault;
      Form1.Cursor:=crDefault;
      Form1.Refresh;
      mtasts:=1;
    end;
  end;
end;
// Nummer Personenart Vorname                   Nachname                  Strasse                        Hausnummer Firmenname
//   1234    123456   123456789.123456789.12345 123456789.123456789.12345 123456789.123456789.123456789. 12345      123456789.123456789.123456789.
// PLZ        Ort                            Telefon              Handy                EMail                          Land
// 123456789. 123456789.123456789.123456789. 123456789.123456789. 123456789.123456789. 123456789.123456789.123456789. 123456789.123456789.123456789.
// Partner: Anrede Vorname                   Nachname                  Telefon
//          1234   123456789.123456789.12345 123456789.123456789.12345 123456789.123456789.
// --------------------------------------------------------------------------------------------------------------------------------------------------------------------

procedure TForm1.ScrollBar1Change(Sender: TObject);
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  Label2.Caption:='';
  if (mtasts = 1) then
  begin
    if ScrollBar1.Position <> ScrollposaltX then
    begin
      ScrollposaltX:=ScrollBar1.Position;
      Bildladen;
    end;
  end;
end;

procedure TForm1.ScrollBar2Change(Sender: TObject);
begin
  Label1.Caption:=UTF8Encode(#169)+'LINSOFT               Projektverwaltung                   Datum: '+FormatDateTime('DD.MM.YYYY',now);
  Label2.Caption:='';
  if (mtasts = 1) then
  begin
    if ScrollBar2.Position <> ScrollposaltY then
    begin
      ScrollposaltY:=ScrollBar2.Position;
      Bildladen;
    end;
  end;
end;

end.

