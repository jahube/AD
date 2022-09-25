
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$path = "C:\Skripte"
$Zip = "chilkatdotnet48-9.5.0-x64.zip"
$DLL = "chilkatdotnet48-9.5.0-x64\ChilkatDotNet48.dll"

$mailbox = "testuser87@pshell.site"
$password = "Test1234567="
$mailbox_server = "WIN-V38U3L9KFM9.pshell.site"


If (!(Test-Path $path)) { mkdir $path }
If (!(Test-Path "$temp\$Zip")) { wget "https://chilkatdownload.com/9.5.0.91/$Zip" -OutFile "$temp\$Zip" }
If (!(Test-Path $path)) { mkdir $path }
If (!(Test-Path "$path\$Dll")) { Expand-Archive -LiteralPath "$temp\$Zip" -DestinationPath $path }

Add-Type -Path  "$path\$DLL"

$Assembly = [System.Reflection.Assembly]::LoadFrom("$path\$DLL");

# This example assumes the Chilkat API to have been previously unlocked.
# See Global Unlock Sample for sample code.

$imap = New-Object Chilkat.Imap

# Use TLS
$imap.Ssl = $true
$imap.Port = 993
$success = $imap.Connect($mailbox_server)
if ($success -ne $true) {
    $($imap.LastErrorText)
    exit
}

# Authenticate
$success = $imap.Login($mailbox,$password)
if ($success -ne $true) {
    $($imap.LastErrorText)
    exit
}

# Get the list of capabilities:
$caps = $imap.Capability()
$("Capabilities: " + $caps)

# Here is an example of the string returned:
# * CAPABILITY IMAP4rev1 UNSELECT IDLE NAMESPACE QUOTA ID XLIST CHILDREN X-GM-EXT-1 
# UIDPLUS COMPRESS=DEFLATE ENABLE MOVE CONDSTORE ESEARCH UTF8=ACCEPT APPENDLIMIT=35882577
# LIST-EXTENDED LIST-STATUS

# Chilkat v9.5.0.58 introduces the HasCapability method to
# check to see if a particular capability exists:
if ($imap.HasCapability("QUOTA",$caps) -eq $true) {
    $("IMAP server supports the QUOTA extension.")
}

if ($imap.HasCapability("IDLE",$caps) -eq $true) {
    $("IMAP server supports IDLE.")
}

# Disconnect from the IMAP server.
$success = $imap.Disconnect()
