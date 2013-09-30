program AutoCAD_Converter;

{%ToDo 'AutoCAD_Converter.todo'}

uses
  Forms,
  unMain in 'unMain.pas' {fmMain},
  kaaAcadConverter in 'SourceCode\kaaAcadConverter.pas',
  AutoCAD_TLB in 'SourceCode\AutoCAD_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
