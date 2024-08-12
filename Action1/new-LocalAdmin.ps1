<#
.SYNOPSIS
This script creates or updates a local administrator account on each workstation with a unique, randomly generated password.
The password is verified before being written back to Action1 RMM as a custom attribute.
Additionally, it ensures that the account is hidden from the login screen.
It can be run on a schedule to rotate the local administrator account on a workstation that already has an existing entry.

.DESCRIPTION
- Generates a random password with a specified length and complexity.
- Creates a new local administrator account if it does not exist, or updates the password if it does.
- Ensures the password does not expire.
- Verifies the password by attempting a simple operation.
- Updates Action1 RMM with the new password only after successful verification.
- Hides the administrator account from the login screen by modifying the registry.

.NOTES
- You can change the username of the local administrator account by modifying the `$username` variable.
- Ensure that the value entered for the `$customField` variable matches an existing custom attribute field in your Action1 RMM instance.
- Adjust the password length and complexity by modifying the `$minLength`, `$maxLength`, and `$nonAlphaChars` variables.

Author: Anthony Murdoch
Date: 2024-07-09

.EXAMPLE
This script does not take any parameters. Simply run it to create or update the local administrator account and update the custom attribute field in Action1 RMM.
#>

# Password length and complexity settings
$minLength = 15
$maxLength = 20
$length = Get-Random -Minimum $minLength -Maximum $maxLength
$nonAlphaChars = 5

# Action1 instance specific settings
$username = "LocalAdministrator"
$customField = "Local Admin Password"

function New-RandomPassword {
    param (
        [int]$length,
        [int]$nonAlphaChars
    )

    $chars = @()
    $chars += [char[]](48..57)  # 0-9
    $chars += [char[]](65..90)  # A-Z
    $chars += [char[]](97..122) # a-z
    $specialChars = [char[]]("!@#$%^&*()-_=+[]{}|;:,.<>?")

    $passwordChars = -join ((1..($length - $nonAlphaChars)) | ForEach-Object { $chars | Get-Random })
    $passwordChars += -join ((1..$nonAlphaChars) | ForEach-Object { $specialChars | Get-Random })

    # Shuffle the password characters
    $password = -join ($passwordChars.ToCharArray() | Get-Random -Count $passwordChars.Length)
    return $password
}

$password = New-RandomPassword -length $length -nonAlphaChars $nonAlphaChars

$group = "Administrators"
$keyPath = "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"

try {
    $user = Get-LocalUser -Name $username -ErrorAction Stop
    Write-Host "Setting password for existing local user $username."
    $user | Set-LocalUser -Password (ConvertTo-SecureString -AsPlainText $password -Force)
} catch {
    Write-Host "Creating new local user $username."
    $user = New-LocalUser -Name $username -Password (ConvertTo-SecureString -AsPlainText $password -Force) -FullName "Local Administrator" -PasswordNeverExpires
    Add-LocalGroupMember -Group $group -Member $user.Name
}

Write-Host "Setting password for $username to never expire."
& WMIC USERACCOUNT WHERE "Name='$username'" SET PasswordExpires=FALSE

# verify the (new) password - I know, this is convoluted but believe me, I tried the other 4+ methods you can think of and they all failed
# I'm not sure whether that's something to do with the permissions of Action1's scriptrunner, but after a lot of trial and error, this is the only way I could get it to work

$testFilePath = "$env:TEMP\admin_test_file.txt"
$testContent = "This is a test file for admin verification."

try {
    # Create the test file
    Set-Content -Path $testFilePath -Value $testContent -Force

    # Set permissions so only the new admin account and SYSTEM can read it
    icacls $testFilePath /inheritance:r
    icacls $testFilePath /grant:r "${env:COMPUTERNAME}\${username}:(R)"
    icacls $testFilePath /grant:r "SYSTEM:(F)"

    # Try to read the file
    $readContent = Get-Content -Path $testFilePath -ErrorAction Stop

    if ($readContent -eq $testContent) {
        Write-Host "Password verification successful. Updating Action1 custom field."
        Action1-Set-CustomAttribute "$customField" "$password"
    } else {
        throw "File content does not match expected value."
    }
} catch {
    Write-Host "Password verification failed. Error details: $_"
    exit 1
} finally {
    # Clean up: remove the test file
    Remove-Item -Path $testFilePath -Force -ErrorAction SilentlyContinue
}

Write-Host "Hiding the admin user from login screen"
New-Item -Path "$keyPath" -Name SpecialAccounts -Force | Out-Null
New-Item -Path "$keyPath\SpecialAccounts" -Name UserList -Force | Out-Null
New-ItemProperty -Path "$keyPath\SpecialAccounts\UserList" -Name $username -Value 0 -PropertyType DWord -Force | Out-Null

Write-Host "Account setup and password update completed."
