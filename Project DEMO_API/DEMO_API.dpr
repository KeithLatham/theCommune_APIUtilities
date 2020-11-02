program DEMO_API;

uses
  Vcl.Forms,
  MAIN_DemoAPI in 'pas\MAIN_DemoAPI.pas' {Form1} ,
  LOGIC_DemoAPI in 'pas\LOGIC_DemoAPI.pas',
  Commune_APIUtilities in 'pas\Commune_APIUtilities.pas';

{$R *.res}

begin Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;

end.
