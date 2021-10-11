unit ufrmInfo;

// |===========================================================================|
// | unit ufrmInfo                                                             |
// | 2010 F.BASSO                                                              |
// |___________________________________________________________________________|
// | unit� permettant l'affichage d'information pour le programme fbcreateuser |
// |___________________________________________________________________________|
// | Ce programme est libre, vous pouvez le redistribuer et ou le modifier     |
// | selon les termes de la Licence Publique G�n�rale GNU publi�e par la       |
// | Free Software Foundation .                                                |
// | Ce programme est distribu� car potentiellement utile,                     |
// | mais SANS AUCUNE GARANTIE, ni explicite ni implicite,                     |
// | y compris les garanties de commercialisation ou d'adaptation              |
// | dans un but sp�cifique.                                                   |
// | Reportez-vous � la Licence Publique G�n�rale GNU pour plus de d�tails.    |
// |                                                                           |
// | anbasso@wanadoo.fr                                                        |
// |___________________________________________________________________________|
// | Versions                                                                  |
// |   1.0.0.0 Cr�ation de l'unit�                                             |
// |===========================================================================|


interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls;

type
  TfrmInfo = class(TForm)
    Memo1: TMemo;
    Image1: TImage;
    lblNomProduit: TLabel;
    lblVersion: TLabel;
    lblCopyright: TLabel;
    Timer1: TTimer;
    procedure Timer1Timer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { D�clarations priv�es }
  public
    { D�clarations publiques }
    procedure AfficheMessage(texte: string);
    procedure fermer;
  end;

var
  frmInfo: TfrmInfo;

implementation

{$R *.dfm}

uses ufonctiondiv;

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Public                                                             #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

procedure TfrmInfo.AfficheMessage( texte : string );

//  ___________________________________________________________________________
// | class procedure TFrmSplash.AfficheMessage                                 |
// | _________________________________________________________________________ |
// || Ajoute un message dans la listebox                                      ||
// ||_________________________________________________________________________||
// || Entr�es | texte : string                                                ||
// ||         |   Texte � afficher                                            ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  if not Assigned(FrmInfo) then
    FrmInfo := TFrmInfo.Create(Application);
  frmInfo.Visible:= true;
  FrmInfo.memo1.Lines.Add(texte);
end;

procedure TfrmInfo.fermer;

//  ___________________________________________________________________________
// | class procedure TfrmInfo.fermer                                           |
// | _________________________________________________________________________ |
// || Permet d'initier la fermeture de la fen�tre                             ||
// ||_________________________________________________________________________||
// || Entr�es |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

Begin
  if not Assigned(FrmInfo) then exit;
  frmInfo.Timer1.Enabled := true;
end;

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Evenementiel                                                       #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

procedure TfrmInfo.FormCreate(Sender: TObject);

//  ___________________________________________________________________________
// | class procedure TFrmSplash.AfficheMessage                                 |
// | _________________________________________________________________________ |
// || Permet d'afficher les information de version                            ||
// ||_________________________________________________________________________||
// || Entr�es | Sender: TObject                                               ||
// ||         |   Objet g�n�rant l'�v�nement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  Caption := StrLectureNomProduit;
  lblNomProduit.caption := StrLectureNomProduit;
  lblVersion.caption :=  'Version : ' + strlectureversion;
  lblCopyright.Caption := StrLectureCopyright ;
end;

procedure TfrmInfo.Timer1Timer(Sender: TObject);

//  ___________________________________________________________________________
// | class procedure TFrmSplash.AfficheMessage                                 |
// | _________________________________________________________________________ |
// || permet la fermeture de la fen�tre apres un d�lais                       ||
// ||_________________________________________________________________________||
// || Entr�es | texte : string                                                ||
// ||         |   Texte � afficher                                            ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  timer1.Enabled := false;
  close;
  memo1.Clear ;
end;

end.
