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
$cnt = (Get-Content $inputFileList -ReadCount 1 | Measure-Object -Line).Lines

$sh = New-Object -Com Shell.Application
$i = 0
$prevFolder = ""

Get-Content $inputFileList -Encoding UTF8 | %{
    $folder = Split-Path $_
    $fn = Split-Path $_ -Leaf

    if ($folder -ne $prevFolder) {
        Write-Host "[$(Get-Date -Format o)] Opening shell folder $folder..."
        $ns = $sh.NameSpace($folder)
        $items = $ns.Items()
    }

    $si = $items.Item($fn)
    Write-Host "[$(Get-Date -Format o)] Item '$fn' opened"
    
    $d = dir $_
    $acl = Get-Acl $_
    Write-Host "[$(Get-Date -Format o)] Item '$fn': file system data read"

    [pscustomobject]@{ 
        Filename = $fn
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
    Write-Host "[$(Get-Date -Format o)] Item '$fn' completed"

    $i += 1
    Write-Progress -Activity "Analyzing files" -CurrentOperation "'$_' examined" -PercentComplete (100 * $i / $cnt)

    $prevFolder = $folder
}
