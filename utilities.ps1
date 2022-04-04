

function Make-Twoway {

# Prompts for two directories
# Creates links form directory 1 to directory 2 and vice versa
# TODO: Add check that directories are dirs, not files

# Prompt for directory 1
$dir1 = Read-Host -Prompt 'What is the first directory you want to link?'
# Prompt for directory 2
$dir2 = Read-Host -Prompt 'What is the second directory you want to link?'

$a = New-Object -ComObject WScript.Shell
$a.Popup("Script will create mutual links between: `n" + $dir1 + "`n and `n" + $dir2, 0, "Proceed?", 32+3)

$base1 = Split-Path $dir1 -Leaf
$base2 = Split-Path $dir2 -Leaf

$name1 = Join-Path -Path $dir1 -ChildPath $base2
$name2 = Join-Path -Path $dir2 -ChildPath $base1

$sh = New-Object -ComObject WScript.Shell

$shortcut = $sh.CreateShortcut($name1 + ".lnk")
$shortcut.TargetPath = $dir2
$shortcut.Save()

$shortcut = $sh.CreateShortcut($name2 + ".lnk")
$shortcut.TargetPath = $dir1
$shortcut.Save()

}

function Make-Hardlink {

# Prompts for a source file and a target file
# Creates a hard link at the target location to source file
# TODO: Add file checking?

# Prompt for source file
$original = Read-Host -Prompt 'What is the file you want to create a hard link on?'
# Prompt for target file
$newhardlink = Read-Host -Prompt 'What is the path where you want the new hard link to exist'

# Trim illegal characters
$original = $original.replace('`n', '').replace('`r', '').replace('`"', '')
$newhardlink = $newhardlink.replace('`n', '').replace('`r', '').replace('`"', '')

$a = New-Object -ComObject WScript.Shell
$answer = $a.Popup("Script will create hardlink to: `n " + $original + '`n' + $newhardlink, 0, 'Proceed?', 32+3)

# Execute the hardlink
New-Item -ItemType HardLink -Path $newhardlink -Target $original

}

function Strip-ExcelPassword {
    param ()
    
    # Mostly taken from https://stackoverflow.com/questions/42860894/remove-known-excel-passwords-with-powershell
    
    # Get Current EXCEL Process ID's so they are not affected but the scripts cleanup
    # SilentlyContinue in case there are no active Excels
    $currentExcelProcessIDs = (Get-Process excel -ErrorAction SilentlyContinue).Id

    $a = Get-Date
    $ErrorActionPreference = "SilentlyContinue"

    CLS

    # Paths
    $c = Get-Location
    $encrypted_Path = $c.Path
    $decrypted_Path = $c.Path
    $processed_Path = 
    $password_Path = 
    
    Write-Host 'Working on folder: ' + $encrypted_Path

    # Load Password Cache
    $arrPasswords = Get-Content -Path $password_Path

    # Load File List
    $arrFiles = Get-ChildItem $encrypted_Path

    # Create counter to display progress
    [int] $count = ($arrfiles.count - 1)

    # New Excel Object
    $ExcelObj = $null
    $ExcelObj = New-Object -ComObject Excel.Application
    $ExcelObj.Visible = $false

    # Loop through each file
    $arrFiles | % {
        $file = get-item -path $_.fullname
        # Display current file
        write-host "`n Processing" $file.name -f "DarkYellow"
        write-host "`n Items remaining: " $count `n

        # Excel xlsx
        if ($file.Extension -like "*.xls*") {

            # Loop through password cache
            $arrPasswords | % {
                $passwd = $_

                # Attempt to open file
                $Workbook = $ExcelObj.Workbooks.Open($file.fullname, 1, $false, 5, $passwd)
                $Workbook.Activate()

                # if password is correct, remove $passwd from array and save new file without password to $decrypted_Path
                if ($Workbook.Worksheets.count -ne 0) 
                {   
                    $Workbook.Password = $null
                    $savePath = $decrypted_Path + $file.Name
                    write-host "Decrypted: " $file.Name -f "DarkGreen"
                    $Workbook.SaveAs($savePath)

                    # Added to keep Excel process memory utilization in check
                    $ExcelObj.Workbooks.close()

                    # Move original file to $processed_Path
                    move-item $file.fullname -Destination $processed_Path -Force

                }
                else {
                    # Close Document
                    $ExcelObj.Workbooks.Close()
                }
            }

        }



        $count--
        # Next File
    }
    # Close Document and Application
    $ExcelObj.Workbooks.close()
    $ExcelObj.Application.Quit()

    Write-host "`nProcessing Complete!" -f "Green"
    Write-host "`nFiles w/o a matching password can be found in the Encrypted folder."
    Write-host "`nTime Started   : " $a.ToShortTimeString()
    Write-host "Time Completed : " $(Get-Date).ToShortTimeString()
    Write-host "`nTotal Duration : " 
    NEW-TIMESPAN –Start $a –End $(Get-Date)

    # Remove any stale Excel processes created by this script's execution
    Get-Process excel -ErrorAction SilentlyContinue | Where-Object { $currentExcelProcessIDs -notcontains $_.id } | Stop-Process
}