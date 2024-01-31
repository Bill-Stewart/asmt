# asmt

**asmt** ("automatic service management tool") is a Windows console (text-based, command-line) program that creates/resets a service and a local service account that runs the service.

## AUTHOR

Bill Stewart - bstewart at iname dot com

## LICENSE

**asmt** is covered by the GNU Public License (GPL). See the file `LICENSE` for details.

## USAGE

Due to that fact that **asmt** manages local accounts and services, it requires administrator privilege (i.e., "Run as administrator"). Command-line parameters (options starting with `--`) are case-sensitive.

---

### Initialize a Service

The `--init` option creates a service and service account if they don't exist, or resets them if they do exist.

`asmt` `--init` `--name=`_servicename_ [`--displayname="`_displayname_`"`] [`--description="`_servicedescription_`"`] `--commandline="`_commandline_`"` `--starttype=`_starttype_ `--account=`_serviceaccountname_ [`--accountdescription="`_serviceaccountdescription_`"`]

Where:

* _servicename_ is the service's name
* _displayname_ is the service's display name
* _servicedescription_ is the service's description
* _commandline_ is the command the service runs
* _starttype_ is the service startup type
* _serviceaccountname_ is the local service account's username
* _serviceaccountdescription_ is the local service account's description

Comments:

* The `--name`, `--commandline`, and `--account` parameters are required, and all of the rest are optional.

* The `--starttype` parameter's argument must be one of the following: `Auto`, `Demand`, `Disabled`, or `DelayedAuto`. If you omit it, the default service startup type is `Auto`.

* The _commandline_ parameter can contain embedded `"` characters by doubling them (i.e., `""`).

The `--init` parameter uses the following logic:

* If the local service account doesn't exist:
  * Creates the local service account with a long, random password
  * Sets the "Password never expires" option
  * Grants the account the "Log on as a service" right
* Otherwise (the local service account exists):
  * Resets the account's password with a long, random password
  * Sets the "Password never expires" option if it's not set
  * Enables the account if it is disabled
  * Grants "Log on as a service" right if needed
* If the service doesn't exist:
  * Creates it with to run with the specified service account and the values specified
* Otherwise (the service exists):
  * Sets the credentials to match the service account
  * Sets the command line and start type to the values specified

---

### Reset a Service

The `--reset` option resets the password of a specified service account and its associated service.

`asmt` `--reset` `--name=`_servicename_ `--account=`_serviceaccountname_

The `--reset` option does the following:

* Resets the service account's password to a long, random password
* Sets the service account's "Password never expires" option if it's not set
* Enables the service account if it's disabled
* Updates the service to start using the service account

---

### Remove a Service

The `--remove` option removes the service and disables its associated service account.

`asmt` `--remove` `--name=`_servicename_ `--account=`_serviceaccountname_

The `--remove` option does the following:

* Disables the service account
* Revokes the "Log on as a service" right
* Stops the service
* Removes the service

> **NOTE:** The `--remove` option does not delete the service account because the account may have permissions assigned over one or more objects in the file system, registry, etc. Deleting the account would "orphan" the account name in these access control lists (ACLs). If you are certain you have removed the service account from all ACLs where it was given permissions, you can manually delete the service account.

---

## VERSION HISTORY

### 0.0.4 (2024-01-31)

* Version bump due to updated WindowsPrivileges source (bug fix).

### 0.0.3 (2024-01-31)

* Version bump due to updated WindowsPrivileges source.

### 0.0.2 (2024-01-22)

* Validated same license (GPL) at top of source files.

* Version bump due to updated WindowsPrivileges source.

### 0.0.1 (2024-01-19)

* Initial version.
