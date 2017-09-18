*********************************************************************************************************
* Created by Ioan Popovici   | Requires PowerShell 3.0, Windows 2012 or better                          *
* ===================================================================================================== *
* Modified by   |    Date    | Revision | Comments                                                      *
* _____________________________________________________________________________________________________ *
* Ioan Popovici | 2017-02-14 | v1.0     | First version                                                 *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This PADT script is used to install SCCM Client on Workgroup Machines.
.DESCRIPTION
    This PADT script is used to install SCCM Client on Workgroup Machines. Client Certificate must be generated
    first on a machine with access to VISMA CA.
.NOTES
    Should work on domain joined machines, but in that case the certificate can be issued trough GPO much easier.
.NOTES
    To do:                                                                                               
    * Better Error Handling.                                                                                
    * Eliminate some hardcoded variables.                                                                    
    * A Web Certificate Enrollment point should eliminate the need for manual password entry or generating.
    * the Client Certificate on another machine, however there are some security concerns.  

## Installation:
* Run Deploy-Application.exe elevated on a machine that resides in a DATAKRAFTVERK Domain or Subdomain.
Machine must be 2012 or better, SCCM agent should be installed and working. When prompted input the Hostname
of the target machine and a certificate password. Remember the password you will need it!
* Copy the whole installation folder to the target machine.
* Run Deploy-Application.exe elevated the target machine. When prompted enter the certificate password from step 1.

## Uninstallation:
* Run Deploy-Application.exe "uninstall" command elevated

## Troubleshooting:
* If you get a Certificate Subject Name Mismatch Error, it means you entered the target NetBiosName wrong at step 1
or you are not running the program on the intended machine.
* Check the logs

## Logs:
* C:\Windows\ccmsetup\Logs
* C:\Windows\Logs\Software

## Tools:
* Check the tools folder
