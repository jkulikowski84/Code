CLS
# note: uses Get-Files function for comparision

	function Get-Files 
	{
		[cmdletbinding()]
		param (
			[parameter(ValueFromPipeline=$true)]
			[string[]]$Path = $PWD,
			[string[]]$Include,
			[string[]]$ExcludeDirs,
			[string[]]$ExcludeFiles,
			[switch]$Recurse,
			[switch]$FullName,
			[switch]$Directory,
			[switch]$File,
			[ValidateSet('Robocopy', 'Dir', 'EnumerateFiles', 'AlphaFS')]
			[string]$Method = 'Robocopy',
			[string]$AlphaFSdllPath = "$env:USERPROFILE\Documents\WindowsPowerShell\AlphaFS.dll"
		)
    
		begin 
		{
			Write-Warning 'robocopy does not use same encoding as powershell for special characters. use alphafs'
			Write-Warning 'enumeratefiles will download everything from OneDrive'
			Write-Warning 'none of these show "size on disk." use get-filesizeondisk for OneDrive'
			
			if ($Directory -and $File) 
			{
				throw 'Cannot use both -Directory and -File at the same time.'
			}

			$Path = (Resolve-Path $Path).ProviderPath

			function CreateFolderObject 
			{
				# commenting out to deal with constrained mode
				#$name = New-Object System.Text.StringBuilder
				$name = ''
				#$null = $name.Append((Split-Path $matches.FullName -Leaf))
				$name += $(Split-Path $matches.FullName -Leaf)
				if (-not $name.ToString().EndsWith('\')) 
				{
					$null += '\'
				}
				Write-Output $(new-object psobject -prop @{
					FullName = $matches.FullName
					DirectoryName = $($matches.FullName.substring(0, $matches.fullname.lastindexof('\')))
					Name = $name.ToString()
					Size = $null
					Extension = '[Directory]'
					DateModified = $null
				})
			}
		}

		process 
		{
			if ($Method -eq 'Robocopy') 
			{
				$params = '/L', '/NJH', '/BYTES', '/FP', '/NC', '/TS', <#'/XJ',#> '/R:0', '/W:0'
			
				if ($Recurse) {$params += '/E'}
				if ($Include) {$params += $Include}
				if ($ExcludeDirs) {$params += '/XD', ('"' + ($ExcludeDirs -join '" "') + '"')}
				if ($ExcludeFiles) {$params += '/XF', ('"' + ($ExcludeFiles -join '" "') + '"')}
			
				foreach ($dir in $Path) 
				{
					# https://stackoverflow.com/a/30244061/4589490
					if ($dir.contains(' ')) 
					{
						$dir = '"' + $dir + ' "'
					}

					foreach ($line in $(robocopy $dir 'c:\tmep' $params)) 
					{
						# folder
						if (!$File -and $line -match '\s+\d+\s+(?<FullName>.*\\)$') 
						{
							if ($Include) 
							{
								if ($matches.FullName -like "*$($include.replace('*',''))*") 
								{
									if ($FullName) 
									{
										Write-Output $( $matches.FullName )
									} 
									else 
									{
										Write-Output $( CreateFolderObject )
									}
								}
							} 
							else 
							{
								if ($FullName) 
								{
									Write-Output $( $matches.FullName )
								} 
								else 
								{
									Write-Output $( CreateFolderObject )
								}
							}
						} 
						
						# file
						elseif (!$Directory -and $line -match '(?<Size>\d+)\s(?<Date>\S+\s\S+)\s+(?<FullName>.*[^\\])$') 
						{
							if ($FullName) 
							{
								Write-Output $( $matches.FullName )
							} 
							else 
							{
								# [System.IO.FileInfo]$matches.fullname
								$name = Split-Path $matches.FullName -Leaf
								
								Write-Output $(new-object psobject -prop @{
									FullName = $matches.FullName
									DirectoryName = Split-Path $matches.FullName
									Name = $name
									Size = [int64]$matches.Size
									Extension = $(if ($name.IndexOf('.') -ne -1) {'.' + $name.split('.')[-1]} else {'[None]'})
									DateModified = $matches.Date
								})
							}
						} 
						else 
						{
							# Uncomment to see all lines that were not matched in the regex above.
							#Write-host "[NOTMATCHED] $line"
						}
					}
				}
			} 
			elseif ($Method -eq 'Dir') 
			{
				$params = @('/a-d', '/-c') # ,'/TA' for last access time instead of date modified (default)
            
				if ($Recurse) { $params += '/S' }
            
				foreach ($dir in $Path) 
				{
					foreach ($line in $(cmd /c dir $dir $params)) 
					{
						switch -Regex ($line) 
						{
							# folder
							'Directory of (?<Folder>.*)' 
							{
								$CurrentDir = $matches.Folder
                            
								if (-not $CurrentDir.EndsWith('\')) 
								{
									$CurrentDir = "$CurrentDir\"
								}
							}

							# file
							'(?<Date>.* [ap]m) +(?<Size>.*?) (?<Name>.*)' 
							{
								if ($FullName) 
								{
									Write-Output $( $CurrentDir + $matches.Name )
								} 
								else 
								{
									[System.IO.FileInfo]($CurrentDir + $matches.Name)
									<#
									Write-Output $([pscustomobject]@{
										Folder = $CurrentDir
										Name = $Matches.Name
										Size = $Matches.Size
										LastWriteTime = [datetime]$Matches.Date
									})
									#>
								}
							}
						}
					}
				}
			} 
			elseif ($Method -eq 'AlphaFS') 
			{
				ipmo $AlphaFSdllPath
            
				if ($Recurse) 
				{
					$searchOption = 'AllDirectories'
				} 
				else 
				{
					$searchOption = 'TopDirectoryOnly'
				}
            
				foreach ($dir in $Path) 
				{
					if ($FullName) 
					{
						Write-Output $( [Alphaleonis.Win32.Filesystem.Directory]::EnumerateFiles($dir, '*.*', $searchOption) )
					} 
					else 
					{
						[Alphaleonis.Win32.Filesystem.Directory]::EnumerateFiles($dir, '*.*', $searchOption) | % {
							Write-Output $( [Alphaleonis.Win32.Filesystem.File]::GetFileSystemEntryInfo($_) | select *, @{n='Extension';e={if ($_.filename.contains('.')) {$_.filename -replace '.*(\.\w+)$', '$1'}}} )
						}
					}
				}
			} 
			elseif ($Method -eq 'EnumerateFiles') 
			{
				if ($Recurse) 
				{
					$searchOption = 'AllDirectories'
				} 
				else 
				{
					$searchOption = 'TopDirectoryOnly'
				}
            
				foreach ($dir in $Path) 
				{
					if ($FullName) 
					{
						Write-Output $( [System.IO.Directory]::EnumerateFiles($dir, '*.*', $searchOption) | % {$_} )
					} 
					else 
					{
						[System.IO.Directory]::EnumerateFiles($dir, '*.*', $searchOption) | % {
							Write-Output $([System.IO.FileInfo]$_)
						}
					}
				}
			}
		}
	}

	$path = "\\pwvs01ldps0001\Fileserver\Software"

	$a1 = New-Object System.Collections.ArrayList
	$a2 = New-Object System.Collections.ArrayList
	$a3 = New-Object System.Collections.ArrayList
	$a4 = New-Object System.Collections.ArrayList
	$a5 = New-Object System.Collections.ArrayList

	$null = 1..3 | % {
		$a1.Add(
			(Measure-Command {
				# the built-in Get-ChildItem is probably still the fastest
				Get-ChildItem $path -File -Recurse
			}).Ticks
		)
		$a2.Add(
			(Measure-Command {
				# Checks hidden files by default
				Get-Files $path -Method EnumerateFiles -File -Recurse
			}).Ticks
		)
		$a3.Add(
			(Measure-Command {
				# somehow, cmd /c dir is faster than robocopy
				Get-Files $path -Method Dir -File -Recurse
			}).Ticks
		)
		$a4.Add(
			(Measure-Command {
				# robocopy is the slowest. faster than dir, but only when getting just the name
				Get-Files $path -Method Robocopy -File -Recurse
			}).Ticks
		)
		$a5.Add(
			(Measure-Command {
				# also pretty slow, but, like robocopy, supports long file names
				Get-Files $path -Method AlphaFS -File -Recurse
			}).Ticks
		)
	}

($a1 | measure -Sum).sum
($a2 | measure -Sum).sum
($a3 | measure -Sum).sum
($a4 | measure -Sum).sum
($a5 | measure -Sum).sum
