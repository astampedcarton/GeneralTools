#Extract the figure information
function Get-RtfPictHex {
    param (
        [string]$Path
    )

    $content = Get-Content -Path $Path -Raw
    $pattern = '{\\pict[\s\S]*?}'  # non-greedy pict block
    $matches = [regex]::Matches($content, $pattern)

    #Keep only the hex data
    $hexList = @()
    foreach ($match in $matches) {
        $hex = ($match.Value -replace '[^\da-fA-F]', '').ToLower()
        $hexList += $hex
    }
    return $hexList
}

function Compare-RtfImages {
    param (
        [string]$oldfile,
        [string]$newfile
    )

    #Extract the picture portion from the RTF.
    #Because we only want to check if there's a differences or not
    #we compare the whole results.
    $hex1 = Get-RtfPictHex -Path $oldfile
    $hex2 = Get-RtfPictHex -Path $newfile

    #Extract and pad the file information and folder information
    $oldfname = Split-Path -Path $oldfile -Leaf
    $newfname = Split-Path -Path $newfile -Leaf

    $oldfname = $oldfname.Padright(30)
    $newfname = $newfname.Padright(30)

    $oldpath  = Split-Path -Path $oldfile -Parent
    $newpath  = Split-Path -Path $newfile -Parent

    Write-Host "`nOld Loc : $oldpath"
    Write-Host "New Loc : $newpath"
    if ($hex1 -ne $hex2){
      $result = "Re-compare figure".Padright(20)
    } else {
      $result = "No difference".Padright(20)
    }
    Write-Host "Old File: $oldfname New File: $newfname : $result"
}

clear
# Example usage:
Compare-RtfImages -oldfile "C:\Temp\figures\figure1\fig-simple.rtf" -newfile "C:\Temp\figures\figure2\fig-simple.rtf"
Compare-RtfImages -oldfile "C:\Temp\figures\figure1\fig-simple2.rtf" -newfile "C:\Temp\figures\figure2\fig-simple2.rtf"
Compare-RtfImages -oldfile "C:\Temp\figures\figure1\fig-simple3.rtf" -newfile "C:\Temp\figures\figure2\fig-simple3.rtf"
