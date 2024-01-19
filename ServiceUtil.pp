{ Copyright (C) 2024 by Bill Stewart (bstewart at iname.com)

  This program is free software; you can redistribute it and/or modify it under
  the terms of the GNU Lesser General Public License as published by the Free
  Software Foundation; either version 3 of the License, or (at your option) any
  later version.

  This program is distributed in the hope that it will be useful, but WITHOUT
  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
  FOR A PARTICULAR PURPOSE. See the GNU General Lesser Public License for more
  details.

  You should have received a copy of the GNU Lesser General Public License
  along with this program. If not, see https://www.gnu.org/licenses/.

}

{$MODE OBJFPC}
{$MODESWITCH UNICODESTRINGS}

{ The purpose of this unit is to wrap the WindowsServiceAccounts and
  WindowsServices units to share a single password pointer when
  creating/resetting the service and its associated local service account.
}

unit ServiceUtil;

interface

uses
  windows,
  WindowsServiceAccounts,
  WindowsServices,
  WindowsPrivileges;

const
  RANDOM_PASSWORD_LENGTH = 127;
  SERVICE_WAIT_SECONDS   = 120;  // 2 minutes

function LocalServiceAccountExists(var ServiceAccountInfo: TServiceAccountInfo): DWORD;

function AddLocalServiceAccount(var ServiceAccountInfo: TServiceAccountInfo): DWORD;

function TestLocalServiceAccountPrivileges(var ServiceAccountInfo: TServiceAccountInfo;
  out HasPrivileges: Boolean): DWORD;

function GrantLocalServiceAccountPrivileges(var ServiceAccountInfo: TServiceAccountInfo): DWORD;

function RevokeLocalServiceAccountPrivileges(var ServiceAccountInfo: TServiceAccountInfo): DWORD;

function ResetLocalServiceAccount(var ServiceAccountInfo: TServiceAccountInfo): DWORD;

function DisableLocalServiceAccount(var ServiceAccountInfo: TServiceAccountInfo): DWORD;

function LocalServiceExists(var ServiceInfo: TServiceInfo): DWORD;

function AddLocalService(var ServiceInfo: TServiceInfo): DWORD;

function ResetLocalServiceStartType(var ServiceInfo: TServiceInfo): DWORD;

function ResetLocalServiceCommandLine(var ServiceInfo: TServiceInfo): DWORD;

function ResetLocalServiceCredential(var ServiceInfo: TServiceInfo): DWORD;

function RemoveLocalService(var ServiceInfo: TServiceInfo): DWORD;

implementation

var
  ServiceAccountPassword: string;

function CryptBinaryToStringW(pbBinary: Pointer;
  cbBinary: DWORD;
  dwFlags: DWORD;
  pszString: LPWSTR;
  var pcchString: DWORD): BOOL;
  stdcall; external 'crypt32.dll';

procedure GetRandomString(Len: Integer; var S: string);
const
  CRYPT_STRING_BASE64 = $00000001;
  CRYPT_STRING_NOCRLF = $40000000;
var
  Bytes: array of Byte;
  I: Integer;
  Flags, NumChars: DWORD;
  Chars: array of Char;
begin
  if Len <= 0 then
    exit;
  SetLength(Bytes, High(Word));
  Randomize();
  for I := 0 to Length(Bytes) - 1 do
    Bytes[I] := Random(High(Byte) + 1);
  Flags := CRYPT_STRING_BASE64 or CRYPT_STRING_NOCRLF;
  if CryptBinaryToStringW(@Bytes[0],  // const BYTE *pbBinary
    Length(Bytes),                    // DWORD      cbBinary
    Flags,                            // DWORD      dwFlags
    nil,                              // LPWSTR     pszString
    NumChars) then                    // DWORD      pcchString
  begin
    SetLength(Chars, NumChars);
    if CryptBinaryToStringW(@Bytes[0],  // const BYTE *pbBinary
      Length(Bytes),                    // DWORD      cbBinary
      Flags,                            // DWORD      dwFlags
      @Chars[0],                        // LPWSTR     pszString
      NumChars) then                    // PDWORD     pcchString
    begin
      if Len > NumChars then
        Len := NumChars;
      SetLength(S, Len);
      Move(Chars[0], S[1], Len * SizeOf(Char));
    end;
    FillChar(Chars[0], Length(Chars) * SizeOf(Char), 0);
  end;
  FillChar(Bytes[0], Length(Bytes), 0);
end;

procedure WipeString(var S: string);
begin
  if Length(S) > 0 then
  begin
    FillChar(S[1], Length(S) * SizeOf(Char), 0);
  end;
end;

function LocalServiceAccountExists(var ServiceAccountInfo: TServiceAccountInfo): DWORD;
begin
  result := ServiceAccountExists(ServiceAccountInfo);
end;

function AddLocalServiceAccount(var ServiceAccountInfo: TServiceAccountInfo): DWORD;
begin
  ServiceAccountInfo.Password := PChar(ServiceAccountPassword);
  result := AddServiceAccount(ServiceAccountInfo);
end;

function ResetLocalServiceAccount(var ServiceAccountInfo: TServiceAccountInfo): DWORD;
begin
  ServiceAccountInfo.Password := PChar(ServiceAccountPassword);
  result := ResetServiceAccount(ServiceAccountInfo);
end;

function TestLocalServiceAccountPrivileges(var ServiceAccountInfo: TServiceAccountInfo;
  out HasPrivileges: Boolean): DWORD;
var
  Privileges: TStringArray;
begin
  SetLength(Privileges, 1);
  Privileges[0] := 'SeServiceLogonRight';
  result := TestAccountPrivileges('', ServiceAccountInfo.UserName, Privileges, HasPrivileges);
end;

function GrantLocalServiceAccountPrivileges(var ServiceAccountInfo: TServiceAccountInfo): DWORD;
var
  Privileges: TStringArray;
begin
  SetLength(Privileges, 1);
  Privileges[0] := 'SeServiceLogonRight';
  result := AddAccountPrivileges('', ServiceAccountInfo.UserName, Privileges);
end;

function RevokeLocalServiceAccountPrivileges(var ServiceAccountInfo: TServiceAccountInfo): DWORD;
var
  Privileges: TStringArray;
begin
  SetLength(Privileges, 1);
  Privileges[0] := 'SeServiceLogonRight';
  result := RemoveAccountPrivileges('', ServiceAccountInfo.UserName, Privileges);
end;

function DisableLocalServiceAccount(var ServiceAccountInfo: TServiceAccountInfo): DWORD;
begin
  result := DisableServiceAccount(ServiceAccountInfo);
end;

function LocalServiceExists(var ServiceInfo: TServiceInfo): DWORD;
begin
  result := ServiceExists(ServiceInfo);
end;

function AddLocalService(var ServiceInfo: TServiceInfo): DWORD;
begin
  ServiceInfo.Password := PChar(ServiceAccountPassword);
  result := AddService(ServiceInfo);
end;

function ResetLocalServiceStartType(var ServiceInfo: TServiceInfo): DWORD;
begin
  result := SetServiceStartType(ServiceInfo);
end;

function ResetLocalServiceCommandLine(var ServiceInfo: TServiceInfo): DWORD;
begin
  result := SetServiceCommandLine(ServiceInfo);
end;

function ResetLocalServiceCredential(var ServiceInfo: TServiceInfo): DWORD;
begin
  ServiceInfo.Password := PChar(ServiceAccountPassword);
  result := SetServiceCredential(ServiceInfo);
end;

function RemoveLocalService(var ServiceInfo: TServiceInfo): DWORD;
begin
  StopService(ServiceInfo, SERVICE_WAIT_SECONDS);
  result := RemoveService(ServiceInfo);
end;

procedure Init();
begin
  GetRandomString(RANDOM_PASSWORD_LENGTH, ServiceAccountPassword);
end;

procedure Done();
begin
  WipeString(ServiceAccountPassword);
end;

initialization
  Init();

finalization
  Done();

end.
