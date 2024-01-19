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

unit WindowsServices;

{$MODE OBJFPC}
{$MODESWITCH UNICODESTRINGS}

interface

uses
  windows;

type
  // Enumerated type values must be in ascending order
  TServiceStartType = (
    Auto        = 2,  // SERVICE_AUTO_START
    Demand      = 3,  // SERVICE_DEMAND_START
    Disabled    = 4,  // SERVICE_DISABLED
    DelayedAuto = 5   // SERVICE_AUTO_START + call ChangeServiceConfig2W
  );
  TServiceState = (
    Stopped         = 1,  // SERVICE_STOPPED
    StartPending    = 2,  // SERVICE_START_PENDING
    StopPending     = 3,  // SERVICE_STOP_PENDING
    Running         = 4,  // SERVICE_RUNNING
    ContinuePending = 5,  // SERVICE_CONTINUE_PENDING
    PausePending    = 6,  // SERVICE_PAUSE_PENDING
    Paused          = 7   // SERVICE_PAUSED
  );

  TServiceInfo = record
    Name: string;
    DisplayName: string;
    Description: string;
    CommandLine: string;
    StartType: TServiceStartType;
    UserName: string;
    Password: PChar;  // Pointer to maintain single copy in memory
  end;

function StringToStartType(const S: string; out StartType: TServiceStartType): Boolean;

function ServiceExists(var ServiceInfo: TServiceInfo): DWORD;

function GetServiceState(var ServiceInfo: TServiceInfo; out State: TServiceState): DWORD;

function StartService(var ServiceInfo: TServiceInfo; const TimeoutSecs: DWORD): DWORD;

function StopService(var ServiceInfo: TServiceInfo; const TimeoutSecs: DWORD): DWORD;

function AddService(var ServiceInfo: TServiceInfo): DWORD;

function SetServiceCommandLine(var ServiceInfo: TServiceInfo): DWORD;

function SetServiceStartType(var ServiceInfo: TServiceInfo): DWORD;

function SetServiceCredential(var ServiceInfo: TServiceInfo): DWORD;

function RemoveService(var ServiceInfo: TServiceInfo): DWORD;

implementation

const
  SERVICE_NO_CHANGE = High(DWORD);
  SERVICE_CONFIG_DESCRIPTION = 1;
  SERVICE_CONFIG_DELAYED_AUTO_START_INFO = 3;
  SERVICE_WAIT_INTERVAL_MILLISECONDS = 500;
  SERVICE_TIMEOUT_MAX_SECONDS = 43200;  // 12 hours

type
  SERVICE_DELAYED_AUTO_START_INFO = record
    fDelayedAutostart: Integer;  // Win32 BOOL
  end;

  SERVICE_DESCRIPTION = record
    lpDescription: LPWSTR;
  end;

function ChangeServiceConfig2W(hService: SC_HANDLE;
  dwInfoLevel: DWORD;
  pInfo: LPVOID): BOOL;
  stdcall; external 'advapi32.dll';

function LowercaseString(const S: string): string;
var
  Locale: LCID;
  Len: DWORD;
  pResult: PChar;
begin
  result := '';
  if S = '' then
    exit;
  Locale := MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT);
  Len := LCMapStringW(Locale,  // LCID    Locale
    LCMAP_LOWERCASE,           // DWORD   dwMapFlags
    PChar(S),                  // LPCWSTR lpSrcStr
    -1,                        // int     cchSrc
    nil,                       // LPWSTR  lpDestStr
    0);                        // int     cchDest
  if Len = 0 then
    exit;
  GetMem(pResult, Len * SizeOf(Char));
  if LCMapStringW(Locale,  // LCID    Locale
    LCMAP_LOWERCASE,       // DWORD   dwMapFlags
    PChar(S),              // LPCWSTR lpSrcStr
    -1,                    // int     cchSrc
    pResult,               // LPWSTR  lpDestStr
    Len) > 0 then          // int     cchDest
  begin
    result := string(pResult);
  end;
  FreeMem(pResult);
end;

function StringToStartType(const S: string; out StartType: TServiceStartType): Boolean;
var
  T: string;
begin
  result := true;
  T := LowercaseString(S);
  case T of
    'auto':        StartType := Auto;
    'demand':      StartType := Demand;
    'disabled':    StartType := Disabled;
    'delayedauto': StartType := DelayedAuto;
  else
    result := false;
  end;
end;

function ServiceExists(var ServiceInfo: TServiceInfo): DWORD;
var
  SCManager, Service: SC_HANDLE;
begin
  SCManager := OpenSCManagerW(nil,  // LPCWSTR lpMachineName
    nil,                            // LPCWSTR lpDatabaseName
    SC_MANAGER_CONNECT);            // DWORD   dwDesiredAccess 
  if SCManager = 0 then
  begin
    result := GetLastError();
    exit;
  end;
  Service := OpenServiceW(SCManager,  // SC_HANDLE hSCManager
    PChar(ServiceInfo.Name),          // LPCWSTR   lpServiceName
    SERVICE_QUERY_STATUS);            // DWORD     dwDesiredAccess
  if Service <> 0 then
    result := ERROR_SUCCESS
  else
    result := GetLastError();
  CloseServiceHandle(Service);  // SC_HANDLE hSCObject
  CloseServiceHandle(SCManager);  // SC_HANDLE hSCObject
end;

function GetServiceState(var ServiceInfo: TServiceInfo; out State: TServiceState): DWORD;
var
  SCManager, Service: SC_HANDLE;
  Status: SERVICE_STATUS;
begin
  result := ERROR_SUCCESS;
  SCManager := OpenSCManagerW(nil,  // LPCWSTR lpMachineName
    nil,                            // LPCWSTR lpDatabaseName
    SC_MANAGER_CONNECT);            // DWORD   dwDesiredAccess
  if SCManager = 0 then
  begin
    result := GetLastError();
    exit;
  end;
  Service := OpenServiceW(SCManager,  // SC_HANDLE hSCManager
    PChar(ServiceInfo.Name),          // LPCWSTR   lpServiceName
    SERVICE_QUERY_STATUS);            // DWORD     dwDesiredAccess
  if Service = 0 then
    result := GetLastError();
  if result = ERROR_SUCCESS then
  begin
    if QueryServiceStatus(Service,  // SC_HANDLE        hService
      Status) then                  // LPSERVICE_STATUS lpServiceStatus
    begin
      State := TServiceState(Status.dwCurrentState);
    end
    else
      result := GetLastError();
    CloseServiceHandle(Service);  // SC_HANDLE hSCObject
  end;
  CloseServiceHandle(SCManager);  // SC_HANDLE hSCObject
end;

function StartService(var ServiceInfo: TServiceInfo; const TimeoutSecs: DWORD): DWORD;
var
  SCManager, Service: SC_HANDLE;
  WaitTime: DWORD;
  Status: SERVICE_STATUS;
  State: TServiceState;
begin
  if TimeoutSecs > SERVICE_TIMEOUT_MAX_SECONDS then
  begin
    result := ERROR_INVALID_PARAMETER;
    exit;
  end;
  result := ERROR_SUCCESS;
  SCManager := OpenSCManagerW(nil,  // LPCWSTR lpMachineName
    nil,                            // LPCWSTR lpDatabaseName
    SC_MANAGER_CONNECT);            // DWORD   dwDesiredAccess
  if SCManager = 0 then
  begin
    result := GetLastError();
    exit;
  end;
  Service := OpenServiceW(SCManager,  // SC_HANDLE hSCManager
    PChar(ServiceInfo.Name),          // LPCWSTR   lpServiceName
    SERVICE_ALL_ACCESS);              // DWORD     dwDesiredAccess
  if Service <> 0 then
  begin
    if StartServiceW(Service,  // SC_HANDLE hService
      0,                       // DWORD     dwNumServiceArgs
      nil) then                // LPCWSTR   *lpServiceArgVectors
    begin
      if TimeoutSecs > 0 then
      begin
        WaitTime := 0;
        repeat
          Sleep(SERVICE_WAIT_INTERVAL_MILLISECONDS);
          Inc(WaitTime, SERVICE_WAIT_INTERVAL_MILLISECONDS);
          if QueryServiceStatus(Service,  // SC_HANDLE        hService
            Status) then                  // LPSERVICE_STATUS lpServiceStatus
          begin
            State := TServiceState(Status.dwCurrentState);
          end
          else
          begin
            result := GetLastError();
            break;
          end;
          if WaitTime > TimeoutSecs * 1000 then
          begin
            result := ERROR_SERVICE_REQUEST_TIMEOUT;
            break;
          end;
        until State = Running;
      end;
    end
    else
      result := GetLastError();
    CloseServiceHandle(Service);  // SC_HANDLE hSCObject
  end
  else
    result := GetLastError();
  CloseServiceHandle(SCManager);  // SC_HANDLE hSCObject
end;

function StopService(var ServiceInfo: TServiceInfo; const TimeoutSecs: DWORD): DWORD;
var
  SCManager, Service: SC_HANDLE;
  Status: SERVICE_STATUS;
  WaitTime: DWORD;
  State: TServiceState;
begin
  if TimeoutSecs > SERVICE_TIMEOUT_MAX_SECONDS then
  begin
    result := ERROR_INVALID_PARAMETER;
    exit;
  end;
  SCManager := OpenSCManagerW(nil,  // LPCWSTR lpMachineName
    nil,                            // LPCWSTR lpDatabaseName
    SC_MANAGER_CONNECT);            // DWORD   dwDesiredAccess
  if SCManager = 0 then
  begin
    result := GetLastError();
    exit;
  end;
  result := ERROR_SUCCESS;
  Service := OpenServiceW(SCManager,  // SC_HANDLE hSCManager
    PChar(ServiceInfo.Name),          // LPCWSTR   lpServiceName
    SERVICE_ALL_ACCESS);              // DWORD     dwDesiredAccess
  if Service <> 0 then
  begin
    if ControlService(Service,  // SC_HANDLE        hService
      SERVICE_CONTROL_STOP,     // DWORD            dwControl
      Status) then              // LPSERVICE_STATUS lpServiceStatus
    begin
      if TimeoutSecs > 0 then
      begin
        WaitTime := 0;
        repeat
          Sleep(SERVICE_WAIT_INTERVAL_MILLISECONDS);
          Inc(WaitTime, SERVICE_WAIT_INTERVAL_MILLISECONDS);
          if QueryServiceStatus(Service,  // SC_HANDLE        hService
            Status) then                  // LPSERVICE_STATUS lpServiceStatus
          begin
            State := TServiceState(Status.dwCurrentState);
          end
          else
          begin
            result := GetLastError();
            break;
          end;
          if WaitTime > TimeoutSecs * 1000 then
          begin
            result := ERROR_SERVICE_REQUEST_TIMEOUT;
            break;
          end;
        until State = Stopped;
      end;
    end
    else
      result := GetLastError();
    CloseServiceHandle(Service);  // SC_HANDLE hSCObject
  end
  else
    result := GetLastError();
  CloseServiceHandle(SCManager);  // SC_HANDLE hSCObject
end;

function AddService(var ServiceInfo: TServiceInfo): DWORD;
var
  SCManager, Service: SC_HANDLE;
  StartType: TServiceStartType;
  SDASI: SERVICE_DELAYED_AUTO_START_INFO;
  SD: SERVICE_DESCRIPTION;
begin
  result := ERROR_SUCCESS;
  SCManager := OpenSCManagerW(nil,  // LPCWSTR lpMachineName
    nil,                            // LPCWSTR lpDatabaseName
    SC_MANAGER_CREATE_SERVICE);     // DWORD   dwDesiredAccess
  if SCManager = 0 then
  begin
    result := GetLastError();
    exit;
  end;
  if ServiceInfo.StartType = DelayedAuto then
    StartType := Auto
  else
    StartType := ServiceInfo.StartType;
  Service := CreateServiceW(SCManager,  // SC_HANDLE hSCManager
    PChar(ServiceInfo.Name),            // LPCWSTR   lpServiceName
    PChar(ServiceInfo.DisplayName),     // LPCWSTR   lpDisplayName
    SERVICE_ALL_ACCESS,                 // DWORD     dwDesiredAccess
    SERVICE_WIN32_OWN_PROCESS,          // DWORD     dwServiceType
    DWORD(StartType),                   // DWORD     dwStartType
    SERVICE_ERROR_NORMAL,               // DWORD     dwErrorControl
    PChar(ServiceInfo.CommandLine),     // LPCWSTR   lpBinaryPathName
    nil,                                // LPCWSTR   lpLoadOrderGroup
    nil,                                // LPDWORD   lpdwTagId
    nil,                                // LPCWSTR   lpDependencies
    PChar(ServiceInfo.UserName),        // LPCWSTR   lpServiceStartName
    ServiceInfo.Password);              // LPCWSTR   lpPassword
  if Service <> 0 then
  begin
    if ServiceInfo.StartType = DelayedAuto then
    begin
      FillChar(SDASI, SizeOf(SDASI), 0);
      SDASI.fDelayedAutostart := 1;
      if not ChangeServiceConfig2W(Service,      // SC_HANDLE hService
        SERVICE_CONFIG_DELAYED_AUTO_START_INFO,  // DWORD     dwInfoLevel
        @SDASI) then                             // LPVOID    lpInfo
      begin
        result := GetLastError();
      end;
    end;
  end
  else
    result := GetLastError();

  if result = ERROR_SUCCESS then
  begin
    FillChar(SD, SizeOf(SD), 0);
    SD.lpDescription := PChar(ServiceInfo.Description);
    if not ChangeServiceConfig2W(Service,  // SC_HANDLE hService
      SERVICE_CONFIG_DESCRIPTION,          // DWORD     dwInfoLevel
      @SD) then                            // LPVOID    lpInfo
    begin
      result := GetLastError();
    end;
  end;

  CloseServiceHandle(Service);  // SC_HANDLE hSCObject
  CloseServiceHandle(SCManager);  // SC_HANDLE hSCObject
end;

function SetServiceCommandLine(var ServiceInfo: TServiceInfo): DWORD;
var
  SCManager, Service: SC_HANDLE;
begin
  result := ERROR_SUCCESS;
  SCManager := OpenSCManagerW(nil,  // LPCWSTR lpMachineName
    nil,                            // LPCWSTR lpDatabaseName
    SC_MANAGER_CONNECT);            // DWORD   dwDesiredAccess
  if SCManager = 0 then
  begin
    result := GetLastError();
    exit;
  end;
  Service := OpenServiceW(SCManager,  // SC_HANDLE hSCManager
    PChar(ServiceInfo.Name),          // LPCWSTR   lpServiceName
    SERVICE_CHANGE_CONFIG);           // DWORD     dwDesiredAccess
  if Service <> 0 then
  begin
    if not ChangeServiceConfigW(Service,  // SC_HANDLE hService
      SERVICE_NO_CHANGE,                  // DWORD     dwServiceType
      SERVICE_NO_CHANGE,                  // DWORD     dwStartType
      SERVICE_NO_CHANGE,                  // DWORD     dwErrorControl
      PChar(ServiceInfo.CommandLine),     // LPCWSTR   lpBinaryPathName
      nil,                                // LPCWSTR   lpLoadOrderGroup
      nil,                                // LPDWORD   lpdwTagId
      nil,                                // LPCWSTR   lpDependencies
      nil,                                // LPCWSTR   lpServiceStartName
      nil,                                // LPCWSTR   lpPassword
      nil) then                           // LPCWSTR   lpDisplayName
    begin
      result := GetLastError();
    end;
    CloseServiceHandle(Service);  // SC_HANDLE hSCObject
  end
  else
    result := GetLastError();
  CloseServiceHandle(SCManager);  // SC_HANDLE hSCObject
end;

function SetServiceStartType(var ServiceInfo: TServiceInfo): DWORD;
var
  SCManager, Service: SC_HANDLE;
  StartType: TServiceStartType;
  SDASI: SERVICE_DELAYED_AUTO_START_INFO;
begin
  result := ERROR_SUCCESS;
  SCManager := OpenSCManagerW(nil,  // LPCWSTR lpMachineName
    nil,                            // LPCWSTR lpDatabaseName
    SC_MANAGER_CONNECT);            // DWORD   dwDesiredAccess
  if SCManager = 0 then
  begin
    result := GetLastError();
    exit;
  end;
  Service := OpenServiceW(SCManager,  // SC_HANDLE hSCManager
    PChar(ServiceInfo.Name),          // LPCWSTR   lpServiceName
    SERVICE_CHANGE_CONFIG);           // DWORD     dwDesiredAccess
  if Service <> 0 then
  begin
    if ServiceInfo.StartType = DelayedAuto then
      StartType := Auto
    else
      StartType := ServiceInfo.StartType;
    if ChangeServiceConfigW(Service,  // SC_HANDLE hService
      SERVICE_NO_CHANGE,              // DWORD     dwServiceType
      DWORD(StartType),               // DWORD     dwStartType
      SERVICE_NO_CHANGE,              // DWORD     dwErrorControl
      nil,                            // LPCWSTR   lpBinaryPathName
      nil,                            // LPCWSTR   lpLoadOrderGroup
      nil,                            // LPDWORD   lpdwTagId
      nil,                            // LPCWSTR   lpDependencies
      nil,                            // LPCWSTR   lpServiceStartName
      nil,                            // LPCWSTR   lpPassword
      nil) then                       // LPCWSTR   lpDisplayName
    begin
      if ServiceInfo.StartType = DelayedAuto then
      begin
        FillChar(SDASI, SizeOf(SDASI), 0);
        SDASI.fDelayedAutostart := 1;
        if not ChangeServiceConfig2W(Service,      // SC_HANDLE hService
          SERVICE_CONFIG_DELAYED_AUTO_START_INFO,  // DWORD     dwInfoLevel
          @SDASI) then                             // LPVOID    lpInfo
        begin
          result := GetLastError();
        end;
      end;
    end
    else
      result := GetLastError();
    CloseServiceHandle(Service);  // SC_HANDLE hSCObject
  end
  else
    result := GetLastError();
  CloseServiceHandle(SCManager);  // SC_HANDLE hSCObject
end;

function SetServiceCredential(var ServiceInfo: TServiceInfo): DWORD;
var
  SCManager, Service: SC_HANDLE;
begin
  result := ERROR_SUCCESS;
  SCManager := OpenSCManagerW(nil,  // LPCWSTR lpMachineName
    nil,                            // LPCWSTR lpDatabaseName
    SC_MANAGER_CONNECT);            // DWORD   dwDesiredAccess
  if SCManager = 0 then
  begin
    result := GetLastError();
    exit;
  end;
  Service := OpenServiceW(SCManager,  // SC_HANDLE hSCManager
    PChar(ServiceInfo.Name),          // LPCWSTR   lpServiceName
    SERVICE_CHANGE_CONFIG);           // DWORD     dwDesiredAccess
  if Service <> 0 then
  begin
    if not ChangeServiceConfigW(Service,  // SC_HANDLE hService
      SERVICE_NO_CHANGE,                  // DWORD     dwServiceType
      SERVICE_NO_CHANGE,                  // DWORD     dwStartType
      SERVICE_NO_CHANGE,                  // DWORD     dwErrorControl
      nil,                                // LPCWSTR   lpBinaryPathName
      nil,                                // LPCWSTR   lpLoadOrderGroup
      nil,                                // LPDWORD   lpdwTagId
      nil,                                // LPCWSTR   lpDependencies
      PChar(ServiceInfo.UserName),        // LPCWSTR   lpServiceStartName
      ServiceInfo.Password,               // LPCWSTR   lpPassword
      nil) then                           // LPCWSTR   lpDisplayName
    begin
      result := GetLastError();
    end;
    CloseServiceHandle(Service);  // SC_HANDLE hSCObject
  end
  else
    result := GetLastError();
  CloseServiceHandle(SCManager);  // SC_HANDLE hSCObject
end;

function RemoveService(var ServiceInfo: TServiceInfo): DWORD;
var
  SCManager, Service: SC_HANDLE;
begin
  result := ERROR_SUCCESS;
  SCManager := OpenSCManagerW(nil,  // LPCWSTR lpMachineName
    nil,                            // LPCWSTR lpDatabaseName
    SC_MANAGER_CONNECT);            // DWORD   dwDesiredAccess
  if SCManager = 0 then
  begin
    result := GetLastError();
    exit;
  end;
  Service := OpenServiceW(SCManager,  // SC_HANDLE hSCManager
    PChar(ServiceInfo.Name),          // LPCWSTR   lpServiceName
    SERVICE_ALL_ACCESS);              // DWORD     dwDesiredAccess
  if Service <> 0 then
  begin
    if not DeleteService(Service) then  // SC_HANDLE hService
      result := GetLastError();
    CloseServiceHandle(Service);  // SC_HANDLE hSCObject
  end
  else
    result := GetLastError();
  CloseServiceHandle(SCManager);  // SC_HANDLE hSCObject
end;

begin
end.
