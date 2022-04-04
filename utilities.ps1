

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