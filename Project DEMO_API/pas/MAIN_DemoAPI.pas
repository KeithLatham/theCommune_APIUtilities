unit MAIN_DemoAPI;

interface

uses
    Winapi.Windows,
    Winapi.Messages,
    System.SysUtils,
    System.Variants,
    System.Classes,
    Vcl.Graphics,
    Vcl.Controls,
    Vcl.Forms,
    Vcl.Dialogs,
    Vcl.StdCtrls;

type
    TForm1 = class(TForm)
        Button_Close: TButton;
        Button_Enumerate: TButton;
        ComboBox1: TComboBox;
        Memo1: TMemo;
        procedure Button_CloseClick(Sender: TObject);
        procedure Button_EnumerateClick(Sender: TObject);
        procedure FormCreate(Sender: TObject);
    end;

var
    Form1: TForm1;

implementation

uses
    LOGIC_DemoAPI,
    Commune_APIUtilities;

{$R *.dfm}

procedure TForm1.Button_CloseClick(Sender: TObject);
begin
    close;
end;

procedure TForm1.Button_EnumerateClick(Sender: TObject);
var
    myDrive: string;
    myDriveList: tStringlist;
begin
    myDriveList := tAPIUtility.EnumerateAllDrives;
    try
        tDemo.LOAD_Combo(ComboBox1, myDriveList);
        if Form1.ComboBox1.Items.Count > 0 then tDemo.INFO_AllDrives(ComboBox1, Memo1);
    finally
        myDriveList.free;
    end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
    Memo1.Clear;
    ComboBox1.Clear;
end;

end.
