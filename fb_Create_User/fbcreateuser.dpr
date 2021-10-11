program fbcreateuser;

uses
  Forms,
  UfrmMaitre in 'UfrmMaitre.pas' {frmMaitre},
  Uuser in 'Uuser.pas',
  ufrmUser in 'ufrmUser.pas' {frmUser},
  ufrmGroupe in 'ufrmGroupe.pas' {frmGroupe},
  ufrmSites in 'ufrmSites.pas' {frmSites},
  uEntite in 'uEntite.pas',
  ufrmInfo in 'ufrmInfo.pas' {frmInfo},
  uBanques in 'uBanques.pas',
  ufrmBanques in 'ufrmBanques.pas' {frmBanque},
  ufonctionDiv in 'ufonctionDiv.pas';

{$R *.res}
begin
  Application.Initialize;
  Application.Title := 'fb Create User';
  Application.CreateForm(TfrmMaitre, frmMaitre);
  Application.CreateForm(TfrmGroupe, frmGroupe);
  Application.CreateForm(TfrmBanque, frmBanque);
  Application.CreateForm(TfrmSites, frmSites);
  Application.CreateForm(TfrmUser, frmUser);
  Application.Run;
end.
