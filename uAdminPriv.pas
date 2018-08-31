unit uAdminPriv;

interface

uses
  WinSvc, SysUtils, Windows, ShellAPI;

type
  TAdminPriv = class
  public
    class function IsAdmin : Boolean;
    class procedure RunAsAdmin(const AFile: string;
      const AParameters: string = ''; Handle: HWND = 0);
  end;

implementation

{ TAdminPriv }

class function TAdminPriv.IsAdmin: Boolean;
var
  H: SC_HANDLE;
begin
  if Win32Platform <> VER_PLATFORM_WIN32_NT then
    Result := True
  else begin
    H := OpenSCManager(PChar('localhost'), nil, SC_MANAGER_ALL_ACCESS);
    Result := H <> 0;
    if Result then
      CloseServiceHandle(H);
  end;
end;

class procedure TAdminPriv.RunAsAdmin(const AFile, AParameters: string;
  Handle: HWND);
var
  sei: TShellExecuteInfo;
begin                   
  FillChar(sei, SizeOf(sei), 0);

  sei.cbSize := SizeOf(sei);
  sei.Wnd := Handle;
  sei.fMask := SEE_MASK_FLAG_DDEWAIT or SEE_MASK_FLAG_NO_UI;
  sei.lpVerb := 'runas';
  sei.lpFile := PChar(aFile);
  sei.lpParameters := PChar(aParameters);
  sei.nShow := SW_SHOWNORMAL;

  if not ShellExecuteEx(@sei) then
    RaiseLastOSError;
end;

end.
