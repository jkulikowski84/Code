CLS

function Set-Privilege {
    [OutputType('System.Boolean')]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet(
            'SeAssignPrimaryTokenPrivilege', 'AssignPrimaryToken',
            'SeAuditPrivilege', 'Audit',
            'SeBackupPrivilege', 'Backup',
            'SeChangeNotifyPrivilege', 'ChangeNotify',
            'SeCreateGlobalPrivilege', 'CreateGlobal',
            'SeCreatePagefilePrivilege', 'CreatePagefile',
            'SeCreatePermanentPrivilege', 'CreatePermanent',
            'SeCreateSymbolicLinkPrivilege', 'CreateSymbolicLink',
            'SeCreateTokenPrivilege', 'CreateToken',
            'SeDebugPrivilege', 'Debug',
            'SeEnableDelegationPrivilege', 'EnableDelegation',
            'SeImpersonatePrivilege', 'Impersonate',
            'SeIncreaseBasePriorityPrivilege', 'IncreaseBasePriority',
            'SeIncreaseQuotaPrivilege', 'IncreaseQuota',
            'SeIncreaseWorkingSetPrivilege', 'IncreaseWorkingSet',
            'SeLoadDriverPrivilege', 'LoadDriver',
            'SeLockMemoryPrivilege', 'LockMemory',
            'SeMachineAccountPrivilege', 'MachineAccount',
            'SeManageVolumePrivilege', 'ManageVolume',
            'SeProfileSingleProcessPrivilege', 'ProfileSingleProcess',
            'SeRelabelPrivilege', 'Relabel',
            'SeRemoteShutdownPrivilege', 'RemoteShutdown',
            'SeRestorePrivilege', 'Restore',
            'SeSecurityPrivilege', 'Security',
            'SeShutdownPrivilege', 'Shutdown',
            'SeSyncAgentPrivilege', 'SyncAgent',
            'SeSystemEnvironmentPrivilege', 'SystemEnvironment',
            'SeSystemProfilePrivilege', 'SystemProfile',
            'SeSystemtimePrivilege', 'SystemTime',
            'SeTakeOwnershipPrivilege', 'TakeOwnership',
            'SeTcbPrivilege', 'Tcb', 'TrustedComputingBase',
            'SeTimeZonePrivilege', 'TimeZone',
            'SeTrustedCredManAccessPrivilege', 'TrustedCredManAccess',
            'SeUndockPrivilege', 'Undock',
            'SeUnsolicitedInputPrivilege', 'UnsolicitedInput'
        )]
        [Alias('PrivilegeName')]
        [string[]]
        $Name,

        [switch]
        $Disable
    )

    begin {
        $signature = '[DllImport("ntdll.dll", EntryPoint = "RtlAdjustPrivilege")]
        public static extern IntPtr SetPrivilege(int Privilege, bool bEnablePrivilege, bool IsThreadPrivilege, out bool PreviousValue);
 
        [DllImport("advapi32.dll")]
        public static extern bool LookupPrivilegeValue(string host, string name, out long pluid);'
        Add-Type -MemberDefinition $signature -Namespace AdjPriv -Name Privilege

        $getPrivilegeConstant = {
            param($str)

            if ($str -eq 'TrustedComputingBase') {
                return 'SeTcbPrivilege'
            } elseif ($str -match '^Se.*Privilege$') {
                return $str
            } else {
                "Se${str}Privilege"
            }
        }
    }

    process {
        foreach ($priv in $Name) {
            [long]$privId = $null
            $null = [AdjPriv.Privilege]::LookupPrivilegeValue($null, (& $getPrivilegeConstant $priv), [ref]$privId)
            ![bool][long][AdjPriv.Privilege]::SetPrivilege($privId, !$Disable, $false, [ref]$null)
        }
    }
}

#Set the necessary Privilege
if ($MyInvocation.InvocationName -ne '.') 
{
    #Set-Privilege -Name SeBackupPrivilege, SeRestorePrivilege
    #Set-Privilege #-Name SeTakeOwnershipPrivilege -Disable
   # Set-Privilege -Name SeRestorePrivilege
   # Set-Privilege -Name se
}

