(*
2019-2020 "Tsusai": Skyslope Ripper, a crappy buggy tool to export transactions
from Skyslope
*)

program SkySlopeRipper;

uses
  Vcl.Forms,
  Main in 'Main.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
