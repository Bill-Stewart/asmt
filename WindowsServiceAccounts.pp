{ Copyright (C) 2024 by Bill Stewart (bstewart at iname.com)

  This program is free software: you can redistribute it and/or modify it under
  the terms of the GNU General Public License as published by the Free Software
  Foundation, either version 3 of the License, or (at your option) any later
  version.

  This program is distributed in the hope that it will be useful, but WITHOUT
  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
  FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
  details.

  You should have received a copy of the GNU General Public License
  along with this program. If not, see https://www.gnu.org/licenses/.

}

unit WindowsServiceAccounts;

{$MODE OBJFPC}
{$MODESWITCH UNICODESTRINGS}

interface

uses
  windows;

type
  TServiceAccountInfo = record
    UserName: string;
    Description: string;
    Password: PChar;  // Pointer to maintain single copy in memory
  end;

function ServiceAccountExists(var ServiceAccountInfo: TServiceAccountInfo): DWORD;

function AddServiceAccount(var ServiceAccountInfo: TServiceAccountInfo): DWORD;

function ResetServiceAccount(var ServiceAccountInfo: TServiceAccountInfo): DWORD;

function DisableServiceAccount(var ServiceAccountInfo: TServiceAccountInfo): DWORD;

implementation

const
  TIMEQ_FOREVER             = High(DWORD);
  USER_MAXSTORAGE_UNLIMITED = High(DWORD);
  USER_PRIV_USER            = 1;
  UF_SCRIPT                 = $00001;
  UF_ACCOUNTDISABLE         = $00002;
  UF_LOCKOUT                = $00010;
  UF_DONT_EXPIRE_PASSWD     = $10000;

type
  LPCWSTR = PChar;
  NET_API_STATUS = DWORD;

  USER_INFO_2 = record
    usri2_name:           LPWSTR;
    usri2_password:       LPWSTR;
    usri2_password_age:   DWORD;
    usri2_priv:           DWORD;
    usri2_home_dir:       LPWSTR;
    usri2_comment:        LPWSTR;
    usri2_flags:          DWORD;
    usri2_script_path:    LPWSTR;
    usri2_auth_flags:     DWORD;
    usri2_full_name:      LPWSTR;
    usri2_usr_comment:    LPWSTR;
    usri2_parms:          LPWSTR;
    usri2_workstations:   LPWSTR;
    usri2_last_logon:     DWORD;
    usri2_last_logoff:    DWORD;
    usri2_acct_expires:   DWORD;
    usri2_max_storage:    DWORD;
    usri2_units_per_week: DWORD;
    usri2_logon_hours:    PBYTE;
    usri2_bad_pw_count:   DWORD;
    usri2_num_logons:     DWORD;
    usri2_logon_server:   LPWSTR;
    usri2_country_code:   DWORD;
    usri2_code_page:      DWORD;
  end;
  PUSER_INFO_2 = ^USER_INFO_2;

function NetUserAdd(servername: LPCWSTR;
  level: DWORD;
  buf: Pointer;
  out parm_err: DWORD): NET_API_STATUS;
  stdcall; external 'netapi32.dll';

function NetUserGetInfo(servername: LPCWSTR;
  username: LPCWSTR;
  level: DWORD;
  out bufptr: Pointer): NET_API_STATUS;
  stdcall; external 'netapi32.dll';

function NetUserSetInfo(servername: LPCWSTR;
  username: LPCWSTR;
  level: DWORD;
  buf: Pointer;
  out parm_err: DWORD): NET_API_STATUS;
  stdcall; external 'netapi32.dll';

function NetApiBufferFree(Buffer: LPVOID): NET_API_STATUS;
  stdcall; external 'netapi32.dll';

function ServiceAccountExists(var ServiceAccountInfo: TServiceAccountInfo): DWORD;
var
  pUserInfo: PUSER_INFO_2;
begin
  result := NetUserGetInfo(nil,          // LPCWSTR servername
    PChar(ServiceAccountInfo.UserName),  // LPCWSTR username
    2,                                   // DWORD   level
    pUserInfo);                          // LPBYTE  bufptr
  if result = ERROR_SUCCESS then
    NetApiBufferFree(pUserInfo);
end;

function AddServiceAccount(var ServiceAccountInfo: TServiceAccountInfo): DWORD;
var
  pUserInfo: PUSER_INFO_2;
  ParmErr: DWORD;
begin
  GetMem(pUserInfo, SizeOf(USER_INFO_2));
  FillChar(pUserInfo^, SizeOf(USER_INFO_2), 0);
  with pUserInfo^ do
  begin
    usri2_name         := PChar(ServiceAccountInfo.UserName);
    usri2_full_name    := nil;
    usri2_comment      := PChar(ServiceAccountInfo.Description);
    usri2_password     := ServiceAccountInfo.Password;
    // UF_SCRIPT bit required when creating account
    usri2_flags        := UF_SCRIPT or UF_DONT_EXPIRE_PASSWD;
    // Remaining bits are required when creating account
    usri2_priv         := USER_PRIV_USER;
    usri2_acct_expires := TIMEQ_FOREVER;
    usri2_max_storage  := USER_MAXSTORAGE_UNLIMITED;
  end;
  result := NetUserAdd(nil,  // LPCWSTR servername
    2,                       // DWORD   level
    pUserInfo,               // LPBYTE  buf
    ParmErr);                // LPDWORD parm_err
  FreeMem(pUserInfo);
end;

function ResetServiceAccount(var ServiceAccountInfo: TServiceAccountInfo): DWORD;
var
  pUserInfo: PUSER_INFO_2;
  ParmErr: DWORD;
begin
  result := NetUserGetInfo(nil,          // LPCWSTR servername
    PChar(ServiceAccountInfo.UserName),  // LPCWSTR username
    2,                                   // DWORD   level
    pUserInfo);                          // LPBYTE  bufptr
  if result <> ERROR_SUCCESS then
    exit;
  with pUserInfo^ do
  begin
    // Reset password
    usri2_password := ServiceAccountInfo.Password;
    // Enable if disabled
    if (usri2_flags and UF_ACCOUNTDISABLE) <> 0 then
      usri2_flags := usri2_flags and (not UF_ACCOUNTDISABLE);
    // Clear "Account is locked out" state if active
    if (usri2_flags and UF_LOCKOUT) <> 0 then
      usri2_flags := usri2_flags and (not UF_LOCKOUT);
    // Set "password never expires" if not set
    if (usri2_flags and UF_DONT_EXPIRE_PASSWD) = 0 then
      usri2_flags := usri2_flags or UF_DONT_EXPIRE_PASSWD;
    // Do not change logon hours
    usri2_logon_hours := nil;
  end;
  result := NetUserSetInfo(nil,          // LPCWSTR servername
    PChar(ServiceAccountInfo.UserName),  // LPCWSTR username
    2,                                   // DWORD   level
    pUserInfo,                           // LPBYTE  buf
    ParmErr);                            // LPDWORD parm_err
  NetApiBufferFree(pUserInfo);
end;

function DisableServiceAccount(var ServiceAccountInfo: TServiceAccountInfo): DWORD;
var
  pUserInfo: PUSER_INFO_2;
  ParmErr: DWORD;
begin
  result := NetUserGetInfo(nil,          // LPCWSTR servername
    PChar(ServiceAccountInfo.UserName),  // LPCWSTR username
    2,                                   // DWORD   level
    pUserInfo);                          // LPBYTE  bufptr
  if result <> ERROR_SUCCESS then
    exit;
  // If UF_ACCOUNTDISABLE is not set (account is enabled)
  if (pUserInfo^.usri2_flags and UF_ACCOUNTDISABLE) = 0 then
  begin
    with pUserInfo^ do
    begin
      // Disable
      usri2_flags := usri2_flags or UF_ACCOUNTDISABLE;
      // Do not change logon hours
      usri2_logon_hours := nil;
    end;
    result := NetUserSetInfo(nil,          // LPCWSTR servername
      PChar(ServiceAccountInfo.UserName),  // LPCWSTR username
      2,                                   // DWORD   level
      pUserInfo,                           // LPBYTE  buf
      ParmErr);                            // LPDWORD parm_err
    end;
  NetApiBufferFree(pUserInfo);
end;

begin
end.
