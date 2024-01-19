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

program asmt;

{ asmt = "automatic service management tool" (for lack of a better name)

  --init:
  * If local service account doesn't exist:
      * Create it with long, random password
      * Set "Password never expires"
      * Grant required privileges/rights
    else
      * Reset the password to long, random password
      * Set "Password never expires" if not set
      * Enable the account if disabled
      * Grant required privileges/rights
  * If service doesn't exist:
      * Create it to start using service account
    else
      * Set the credentials to match service account
      * Set the command line and service start type

  --reset:
  If local service account exists:
    * Reset the password to long, random password
    * If password expiration is enabled, disable it
    * Enable the account if disabled
    * Grant required privileges/rights
  If service exists:
    * Reset the credentials to match service account

  --remove:
  If local service account exists:
    * Disable it
    * Revoke required privileges/rights
  If service exists:
    * Stop the service if it is running
    * Remove the service

  Why this tool?
  * Useful to run an application as a service using a local service account
    that's not a member of any groups (least privilege)
  * Privileges/rights (e.g., SeServiceLogonRight) granted and revoked
    automatically
  * Password generated randomly and managed automatically

  I initially wrote this as a tool for a Syncthing Windows-based installer
  if the user wants to run it as a service. It might be useful for other
  similar types of programs.

}

{$MODE OBJFPC}
{$MODESWITCH UNICODESTRINGS}
{$R *.res}

// wargcv and wgetopts: https://github.com/Bill-Stewart/wargcv
// WindowsPrivileges: https://github.com/Bill-Stewart/PrivMan
uses
  windows,
  wargcv,
  wgetopts,
  FileUtil,
  ServiceUtil,
  WindowsMessages,
  WindowsPrivileges,
  WindowsServiceAccounts,
  WindowsServices;

const
  PROGRAM_NAME = 'asmt';
  PROGRAM_COPYRIGHT = 'Copyright (C) 2024 by Bill Stewart';

type
  // Groupings of command line parameters
  TParams = (
    Help,
    Init,
    Reset,
    Remove);
  TParamSet = set of TParams;
  TRequiredParams = (
    Name,
    CommandLine,
    Account);
  TRequiredParamSet = set of TRequiredParams;
  TCommandLine = object
    ParamSet: TParamSet;
    RequiredParamSet: TRequiredParamSet;
    Error: DWORD;
    ServiceInfo: TServiceInfo;
    ServiceAccountInfo: TServiceAccountInfo;
    procedure InitServiceInfo();
    procedure InitServiceAccountInfo();
    procedure Parse();
  end;

procedure Usage();
begin
  WriteLn(PROGRAM_NAME, ' ', GetFileVersion(ParamStr(0)), ' - ', PROGRAM_COPYRIGHT);
  WriteLn('This is free software and comes with ABSOLUTELY NO WARRANTY.');
  WriteLn();
  WriteLn('asmt - automatic service management tool');
  WriteLn();
  WriteLn('SYNOPSIS');
  WriteLn();
  WriteLn('Create/reset/remove a local service account and create/reset/remove a service');
  WriteLn('that uses the service account to log on.');
  WriteLn();
  WriteLn();
  WriteLn('USAGE');
  WriteLn();
  WriteLn(PROGRAM_NAME, ' --init --name=<servicename> [--displayname="<displayname>"]');
  WriteLn('  [--description="<servicedescription>"] --commandline="<commandline>"');
  WriteLn('  [--starttype=<starttype>] --account=<serviceaccountname>');
  WriteLn('  [--accountdescription="<serviceaccountdescription>"]');
  WriteLn();
  WriteLn(PROGRAM_NAME, ' --reset --name=<servicename> --account=<serviceaccountname>');
  WriteLn();
  WriteLn(PROGRAM_NAME, ' --remove --name=<servicename> --account=<serviceaccountname>');
  WriteLn();
  WriteLn('COMMENTS');
  WriteLn('* --init creates/resets a local service account and associated service');
  WriteLn('* --reset resets password of the service account and updates the service');
  WriteLn('* --remove disables the service account and stops/removes the service');
  WriteLn('* <starttype> is one of: Auto,Demand,Disabled,DelayedAuto');
  WriteLn('* <commandline> can include embedded " by doubling them (i.e., "")');
  WriteLn('* All commands require administrator privilege/elevation');
end;

procedure TCommandLine.InitServiceInfo();
begin
  with ServiceInfo do
  begin
    Name := '';
    DisplayName := '';
    Description := '';
    CommandLine := '';
    StartType := Auto;
    UserName := '';
    Password := nil;
  end;
end;

procedure TCommandLine.InitServiceAccountInfo();
begin
  with ServiceAccountInfo do
  begin
    UserName := '';
    Description := '';
    Password := nil;
  end;
end;

procedure TCommandLine.Parse();
var
  Opts: array[1..12] of TOption;
  Opt: Char;
  I: Integer;
  LongOptName: string;
begin
  with Opts[1] do
  begin
    Name := 'help';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[2] do
  begin
    Name := 'init';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[3] do
  begin
    Name := 'name';
    Has_arg := Optional_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[4] do
  begin
    Name := 'displayname';
    Has_arg := Optional_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[5] do
  begin
    Name := 'description';
    Has_arg := Optional_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[6] do
  begin
    Name := 'commandline';
    Has_arg := Optional_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[7] do
  begin
    Name := 'starttype';
    Has_arg := Optional_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[8] do
  begin
    Name := 'account';
    Has_arg := Optional_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[9] do
  begin
    Name := 'accountdescription';
    Has_arg := Optional_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[10] do
  begin
    Name := 'reset';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[11] do
  begin
    Name := 'remove';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  with Opts[12] do
  begin
    Name := '';
    Has_arg := No_Argument;
    Flag := nil;
    Value := #0;
  end;
  ParamSet := [];
  RequiredParamSet := [];
  Error := ERROR_SUCCESS;
  OptErr := false;
  InitServiceInfo();
  InitServiceAccountInfo();
  repeat
    Opt := GetLongOpts('', @Opts[1], I);
    if Opt = #0 then
    begin
      LongOptName := Opts[I].Name;
      case LongOptName of
        'help':
        begin
          Include(ParamSet, Help);
        end;
        'init':
        begin
          Include(ParamSet, Init);
        end;
        'name':
        begin
          if OptArg <> '' then
          begin
            Include(RequiredParamSet, Name);
            ServiceInfo.Name := OptArg;
          end;
        end;
        'displayname':
        begin
          if OptArg <> '' then
            ServiceInfo.DisplayName := OptArg;
        end;
        'description':
        begin
          if OptArg <> '' then
            ServiceInfo.Description := OptArg;
        end;
        'commandline':
        begin
          if OptArg <> '' then
          begin
            Include(RequiredParamSet, CommandLine);
            ServiceInfo.CommandLine := OptArg;
          end;
        end;
        'starttype':
        begin
          if not StringToStartType(OptArg, ServiceInfo.StartType) then
            Error := ERROR_INVALID_PARAMETER;
        end;
        'account':
        begin
          if OptArg <> '' then
          begin
            Include(RequiredParamSet, Account);
            ServiceInfo.UserName := '.\' + OptArg;
            ServiceAccountInfo.UserName := OptArg;
          end;
        end;
        'accountdescription':
        begin
          if OptArg <> '' then
            ServiceAccountInfo.Description := OptArg;
        end;
        'reset':
        begin
          Include(ParamSet, Reset);
        end;
        'remove':
        begin
          Include(ParamSet, Remove);
        end;
      end;
    end;
  until Opt = EndOfOptions;
  if Error <> ERROR_SUCCESS then
    exit;
  if Help in ParamSet then
    exit;
  // PopCnt returns number of elements in set
  if PopCnt(DWORD(ParamSet)) <> 1 then
  begin
    WriteLn('Must specify one of: --init, --reset, or --remove');
    Error := ERROR_INVALID_PARAMETER;
    exit;
  end;
  if not (RequiredParamSet >= [Name,Account]) then
  begin
    WriteLn('Must specify --name and --account');
    Error := ERROR_INVALID_PARAMETER;
    exit;
  end;
  if Init in ParamSet then
  begin
    if not (RequiredParamSet >= [CommandLine]) then
    begin
      WriteLn('Must Specify --commandline');
      Error := ERROR_INVALID_PARAMETER;
      exit;
    end;
  end;
end;

var
  RC: DWORD;
  CmdLine: TCommandLine;
  ServiceInfo: TServiceInfo;
  ServiceAccountInfo: TServiceAccountInfo;
  HasPrivileges: Boolean;

begin
  RC := ERROR_SUCCESS;

  CmdLine.Parse();

  if (ParamCount = 0) or (Help in CmdLine.ParamSet) then
  begin
    Usage();
    exit;
  end;

  if CmdLine.Error <> ERROR_SUCCESS then
  begin
    RC := CmdLine.Error;
    WriteLn(GetWindowsMessage(RC, true));
    ExitCode := Integer(RC);
    exit;
  end;

  ServiceInfo := CmdLine.ServiceInfo;
  ServiceAccountInfo := CmdLine.ServiceAccountInfo;

  if Init in CmdLine.ParamSet then
  begin
    RC := LocalServiceAccountExists(ServiceAccountInfo);
    if RC = ERROR_SUCCESS then
    begin
      WriteLn('Local service account "', ServiceAccountInfo.UserName, '" exists');
      RC := ResetLocalServiceAccount(ServiceAccountInfo);
      if RC = ERROR_SUCCESS then
        WriteLn('Reset local service account "', ServiceAccountInfo.UserName, '"')
      else
        WriteLn('Failed to reset local service account "', ServiceAccountInfo.UserName);
    end
    else
    begin
      WriteLn('Local service account "', ServiceAccountInfo.UserName, '" does not exist');
      RC := AddLocalServiceAccount(ServiceAccountInfo);
      if RC = ERROR_SUCCESS then
        WriteLn('Created local service account "', ServiceAccountInfo.UserName, '"')
      else
        WriteLn('Failed to create local service account "', ServiceAccountInfo.UserName, '"')
    end;
    if RC = ERROR_SUCCESS then
    begin
      RC := TestLocalServiceAccountPrivileges(ServiceAccountInfo, HasPrivileges);
      if RC = ERROR_SUCCESS then
      begin
        if HasPrivileges then
          WriteLn('Local service account "', ServiceAccountInfo.UserName ,'" has required privileges/rights')
        else
        begin
          WriteLn('Local service account "', ServiceAccountInfo.UserName ,'" does not have required/privileges/rights');
          RC := GrantLocalServiceAccountPrivileges(ServiceAccountInfo);
          if RC = ERROR_SUCCESS then
            WriteLn('Granted local service account "', ServiceAccountInfo.UserName, '" required privileges/rights')
          else
            WriteLn('Failed to grant local service account "', ServiceAccountInfo.UserName, '" required privileges/rights');
        end;
      end
      else
        WriteLn('Unable to determine if local service account "', ServiceAccountInfo.UserName, '" has required privileges/rights');
    end;
    if RC = ERROR_SUCCESS then
    begin
      RC := LocalServiceExists(ServiceInfo);
      if RC = ERROR_SUCCESS then
      begin
        WriteLn('Service "', ServiceInfo.Name, '" exists');
        RC := ResetLocalServiceCredential(ServiceInfo);
        if RC = ERROR_SUCCESS then
          RC := ResetLocalServiceCommandLine(ServiceInfo);
        if RC = ERROR_SUCCESS Then
          RC := ResetLocalServiceStartType(ServiceInfo);
        if RC = ERROR_SUCCESS then
          WriteLn('Reset service "', ServiceInfo.Name, '"')
        else
          WriteLn('Failed to reset service "', ServiceInfo.Name, '"')
      end
      else
      begin
        WriteLn('Service "', ServiceInfo.Name, '" does not exist');
        RC := AddLocalService(ServiceInfo);
        if RC = ERROR_SUCCESS then
          WriteLn('Created service "', ServiceInfo.Name, '"')
        else
          WriteLn('Failed to create service "', ServiceInfo.Name, '"')
      end;
    end;
  end
  else if Reset in CmdLine.ParamSet then
  begin
    RC := ServiceAccountExists(ServiceAccountInfo);
    if RC = ERROR_SUCCESS then
    begin
      WriteLn('Local service account "', ServiceAccountInfo.UserName, '" exists');
      RC := ResetLocalServiceAccount(ServiceAccountInfo);
      if RC = ERROR_SUCCESS then
      begin
        WriteLn('Reset local service account "', ServiceAccountInfo.UserName, '"');
        RC := ServiceExists(ServiceInfo);
        if RC = ERROR_SUCCESS then
        begin
          WriteLn('Service "', ServiceInfo.Name, '" exists');
          RC := ResetLocalServiceCredential(ServiceInfo);
          if RC = ERROR_SUCCESS then
            WriteLn('Reset service "', ServiceInfo.Name, '"')
          else
            WriteLn('Failed to reset service "', ServiceInfo.Name, '"');
        end
        else
          WriteLn('Service "', ServiceInfo.Name, '" does not exist');
      end
      else
        WriteLn('Failed to reset local service account "', ServiceAccountInfo.UserName);
    end
    else
      WriteLn('Local service account "', ServiceAccountInfo.UserName, '" does not exist');
  end
  else if Remove in CmdLine.ParamSet then
  begin
    RC := LocalServiceAccountExists(ServiceAccountInfo);
    if RC = ERROR_SUCCESS then
    begin
      WriteLn('Local service account "', ServiceAccountInfo.UserName, '" exists');
      RC := DisableLocalServiceAccount(ServiceAccountInfo);
      if RC = ERROR_SUCCESS then
        WriteLn('Disabled local service account "', ServiceAccountInfo.UserName, '"')
      else
        WriteLn('Failed to disable local service account "', ServiceAccountInfo.UserName, '"');
      RC := TestLocalServiceAccountPrivileges(ServiceAccountInfo, HasPrivileges);
      if RC = ERROR_SUCCESS then
      begin
        if HasPrivileges then
        begin
          WriteLn('Local service account "', ServiceAccountInfo.UserName, '" has assigned privileges/rights');
          RC := RevokeLocalServiceAccountPrivileges(ServiceAccountInfo);
          if RC = ERROR_SUCCESS then
            WriteLn('Revoked assigned privileges/rights from local service account "', ServiceAccountInfo.UserName, '"')
          else
            WriteLn('Failed to revoke assigned privileges/rights from local service account "', ServiceAccountInfo.UserName, '"');
        end
        else
          WriteLn('Local service account "', ServiceAccountInfo.UserName, '" does not have assigned privileges/rights');
      end
      else
        WriteLn('Unable to determine if local service account "', ServiceAccountInfo.UserName, '" has assigned privileges/rights');
    end
    else
      WriteLn('Local service account "', ServiceAccountInfo.UserName, '" does not exist');
    RC := LocalServiceExists(ServiceInfo);
    if RC = ERROR_SUCCESS then
    begin
      WriteLn('Service "', ServiceInfo.Name, '" exists');
      RC := RemoveLocalService(ServiceInfo);
      if RC = ERROR_SUCCESS then
        WriteLn('Removed service "', ServiceInfo.Name, '"')
      else
        WriteLn('Failed to removed service "', ServiceInfo.Name, '"');
    end
    else
      WriteLn('Service "', ServiceInfo.Name, '" does not exist');
  end;

  ExitCode := Integer(RC);
  if RC = 0 then
    WriteLn(GetWindowsMessage(RC))
  else
    WriteLn(GetWindowsMessage(RC, true));

end.
