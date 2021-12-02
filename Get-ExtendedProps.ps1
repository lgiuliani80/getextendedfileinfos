param(
    [Parameter(Mandatory=$true)][string] $inputFileList
)

$D_TITLE            = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9} 2"
$D_AUTHORS          = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9} 4"
$D_AUTHOR_LAST_SAVE = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9} 8"
$D_LAST_PRINT       = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9} 11"
$D_CREATION_DATE    = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9} 12"
$D_LAST_SAVE        = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9} 13"

Write-Host "[$(Get-Date -Format o)] Starting..."

$sh = New-Object -Com Shell.Application
$files = Get-Content $inputFileList -Encoding UTF8
$i = 0

Write-Host "[$(Get-Date -Format o)] Input file read. Grouping..."

$files | % { 
    [pscustomobject]@{ Path = (Split-Path $_); N = (Split-Path -Leaf $_); FullName = $_ } 
} | Group-Object {
    $_.Path
} | %{
    Write-Host "[$(Get-Date -Format o)] Opening shell folder $($_.Name)..."
    
    $ns = $sh.NameSpace($_.Name)
    $items = $ns.Items()

    Write-Host "[$(Get-Date -Format o)] Shell folder '$($_.Name)' opened. Examining files..."

    $_.Group | %{
        $si = $items.Item($_.N)
        Write-Host "[$(Get-Date -Format o)] Item '$($_.N)' opened"
        $d = dir $_.FullName
        $acl = Get-Acl $_.FullName
        Write-Host "[$(Get-Date -Format o)] Item '$($_.N)': file system data read"
        [pscustomobject]@{ 
            Filename = $_.FullName
            FileCreationDate = $d.CreationTime
            FileLastWriteDate = $d.LastWriteTime
            FileLastAccessDate = $d.LastAccessTime
            FileOwner = $acl.Owner
            ShellModifyDate = $si.ModifyDate
            DetailsOwner = $si.ExtendedProperty("owner")
            DetailsAuthors = $si.ExtendedProperty($D_AUTHORS) -join ', '
            DetailsTitle = $si.ExtendedProperty($D_TITLE)
            DetailsAuthorLastSave = $si.ExtendedProperty($D_AUTHOR_LAST_SAVE)
            DetailsLastPrint = $si.ExtendedProperty($D_LAST_PRINT)
            DetailsLastSave = $si.ExtendedProperty($D_LAST_SAVE)
            DetailsCreationDate = $si.ExtendedProperty($D_CREATION_DATE)
        }
        Write-Host "[$(Get-Date -Format o)] Item '$($_.N)' completed"

        $i += 1
        Write-Progress -Activity "Analyzing files" -CurrentOperation "'$($_.FullName)' examined" -PercentComplete (100 * $i / $files.Count)
    }
}
