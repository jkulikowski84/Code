CLS

$properties = @('givenname', 'initials', 'mobile', 'telephonenumber', 'sn', 'displayname', 'company', 'title', 'mail', 'department', 'samaccountname')

Clear-Variable AA, AB, AAC, AC, BA, BB, BBC, BC, CA, CB, CCA, CC -Force -Confirm:$False -ErrorAction SilentlyContinue

Write-output "`nDSQuery"

(Measure-Command {
$AA = (dsquery * -filter "(&(samAccountType=805306368)(|(mobile=*)(telephonenumber=*)))" -limit 0 -attr $properties).trim() | select -skip 1
}).TotalMilliseconds

Write-output "`nDSQuery with Select-Object"

(Measure-Command {
$AB = ((dsquery * -filter "(&(samAccountType=805306368)(|(mobile=*)(telephonenumber=*)))" -limit 0 -attr $properties).trim() | select -skip 1) | Select-Object @{Name="First Name";Expression={$_.givenname}}, @{Name="Middle Initial";Expression={$_.initials}}, @{Name="Last Name";Expression={$_.sn}}, @{Name="Display Name";Expression={$_.displayname}}, @{Name="SamAccountName";Expression={$_.samaccountname}}, @{Name="Email";Expression={$_.mail}}, @{Name="Mobile";Expression={$_.mobile}}, @{Name="Telephone Number";Expression={$_.telephonenumber}}, @{Name="Title";Expression={$_.title}}, @{Name="Dept";Expression={$_.department}}, @{Name="Company";Expression={$_.company}}
}).TotalMilliseconds

Write-output "`nDSQuery with foreach loop"

(Measure-Command {
$AAC = (dsquery * -filter "(&(samAccountType=805306368)(|(mobile=*)(telephonenumber=*)))" -limit 0 -attr $properties).trim() | select -skip 1

$AC = foreach($i in $AAC)
    {
        [pscustomobject]@{
            FirstName           = [string] $i.givenname
            MiddleName          = [string] $i.initials 
            LastName            = [string] $i.sn 
            DisplayName         = [string] $i.displayname 
            SamAccountName      = [string] $i.samaccountname 
            Email               = [string] $i.mail 
            Mobile              = [string] $i.mobile 
            TelephoneNumber     = [string] $i.telephonenumber 
            Title               = [string] $i.title 
            Dept                = [string] $i.department 
            Company             = [string] $i.company 
        }
    }
}).TotalMilliseconds

Write-output "`nADSISEARCHER"

(Measure-Command {
$searcher = [adsisearcher]::new()
$searcher.Sort.PropertyName = "sn"
$searcher.PageSize = 10000
$searcher.Filter = "(&(samAccountType=805306368)(|(mobile=*)(telephonenumber=*)))"
$searcher.PropertiesToLoad.AddRange($properties)
$BA = ($searcher.FindAll().Properties)
}).TotalMilliseconds

Write-output "`nADSISEARCHER with select-object"

(Measure-Command {
$searcher = [adsisearcher]::new()
$searcher.Sort.PropertyName = "sn"
$searcher.PageSize = 10000
$searcher.Filter = "(&(samAccountType=805306368)(|(mobile=*)(telephonenumber=*)))"
$searcher.PropertiesToLoad.AddRange($properties)
$BB = ($searcher.FindAll().Properties) | Select-Object @{Name="First Name";Expression={$_.givenname}}, @{Name="Middle Initial";Expression={$_.initials}}, @{Name="Last Name";Expression={$_.sn}}, @{Name="Display Name";Expression={$_.displayname}}, @{Name="SamAccountName";Expression={$_.samaccountname}}, @{Name="Email";Expression={$_.mail}}, @{Name="Mobile";Expression={$_.mobile}}, @{Name="Telephone Number";Expression={$_.telephonenumber}}, @{Name="Title";Expression={$_.title}}, @{Name="Dept";Expression={$_.department}}, @{Name="Company";Expression={$_.company}}
}).TotalMilliseconds

Write-output "`nADSISEARCHER with foreach loop"

(Measure-Command {
$searcher = [adsisearcher]::new()
$searcher.Sort.PropertyName = "sn"
$searcher.PageSize = 10000
$searcher.Filter = "(&(samAccountType=805306368)(|(mobile=*)(telephonenumber=*)))"
$searcher.PropertiesToLoad.AddRange($properties)
$BBC = ($searcher.FindAll().Properties)

$BC = foreach($i in $BBC)
    {
        [pscustomobject]@{
            FirstName           = [string] $i.givenname
            MiddleName          = [string] $i.initials 
            LastName            = [string] $i.sn 
            DisplayName         = [string] $i.displayname 
            SamAccountName      = [string] $i.samaccountname 
            Email               = [string] $i.mail 
            Mobile              = [string] $i.mobile 
            TelephoneNumber     = [string] $i.telephonenumber 
            Title               = [string] $i.title 
            Dept                = [string] $i.department 
            Company             = [string] $i.company 
        }
    }
}).TotalMilliseconds

Write-output "`nGet-ADObject"

(Measure-Command {
$CA = (Get-ADObject -LDAPFilter "(&(samAccountType=805306368)(|(mobile=*)(telephonenumber=*)))" -Properties $properties)

}).TotalMilliseconds

Write-output "`nGet-ADObject with Select-Object" 

(Measure-Command {
$CB = (Get-ADObject -LDAPFilter "(&(samAccountType=805306368)(|(mobile=*)(telephonenumber=*)))" -Properties $properties) | Select-Object @{Name="First Name";Expression={$_.givenname}}, @{Name="Middle Initial";Expression={$_.initials}}, @{Name="Last Name";Expression={$_.sn}}, @{Name="Display Name";Expression={$_.displayname}}, @{Name="SamAccountName";Expression={$_.samaccountname}}, @{Name="Email";Expression={$_.mail}}, @{Name="Mobile";Expression={$_.mobile}}, @{Name="Telephone Number";Expression={$_.telephonenumber}}, @{Name="Title";Expression={$_.title}}, @{Name="Dept";Expression={$_.department}}, @{Name="Company";Expression={$_.company}}

}).TotalMilliseconds

Write-output "`nGet-ADObject with foreach loop"

(Measure-Command {
$CCA = (Get-ADObject -LDAPFilter "(&(samAccountType=805306368)(|(mobile=*)(telephonenumber=*)))" -Properties $properties)

    $CC = foreach($i in $CCA)
    {
        [pscustomobject]@{
            FirstName           = [string] $i.givenname
            MiddleName          = [string] $i.initials 
            LastName            = [string] $i.sn 
            DisplayName         = [string] $i.displayname 
            SamAccountName      = [string] $i.samaccountname 
            Email               = [string] $i.mail 
            Mobile              = [string] $i.mobile 
            TelephoneNumber     = [string] $i.telephonenumber 
            Title               = [string] $i.title 
            Dept                = [string] $i.department 
            Company             = [string] $i.company 
        }
    }
}).TotalMilliseconds
