unit ufrmInfo;

// |===========================================================================|
// | unit ufrmInfo                                                             |
// | 2010 F.BASSO                                                              |
// |___________________________________________________________________________|
// | unité permettant l'affichage d'information pour le programme fbcreateuser |
// |___________________________________________________________________________|
// | Ce programme est libre, vous pouvez le redistribuer et ou le modifier     |
// | selon les termes de la Licence Publique Générale GNU publiée par la       |
// | Free Software Foundation .                                                |
// | Ce programme est distribué car potentiellement utile,                     |
// | mais SANS AUCUNE GARANTIE, ni explicite ni implicite,                     |
// | y compris les garanties de commercialisation ou d'adaptation              |
// | dans un but spécifique.                                                   |
// | Reportez-vous à la Licence Publique Générale GNU pour plus de détails.    |
// |                                                                           |
// | anbasso@wanadoo.fr                                                        |
// |___________________________________________________________________________|
// | Versions                                                                  |
// |   1.0.0.0 Création de l'unité                                             |
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
    { Déclarations privées }
  public
    { Déclarations publiques }
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
// || Entrées | texte : string                                                ||
// ||         |   Texte à afficher                                            ||
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
// || Permet d'initier la fermeture de la fenêtre                             ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
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
// || Entrées | Sender: TObject                                               ||
// ||         |   Objet générant l'événement                                  ||
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
// || permet la fermeture de la fenêtre apres un délais                       ||
// ||_________________________________________________________________________||
// || Entrées | texte : string                                                ||
// ||         |   Texte à afficher                                            ||
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
