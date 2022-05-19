$caminho = Get-Location
Import-Module "$caminho\Files\7Zip4Powershell\7Zip4PowerShell.dll"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
<#
.SYNOPSIS
    Show GUI Messagebox
.DESCRIPTION
    Show GUI Messagebox and wait for user input or timeout
.PARAMETER Message
    Message to show
.PARAMETER Title
    Messagebox title
.PARAMETER Buttons
    Messagebox buttons
.PARAMETER Icon
    Messagebox Icon
.PARAMETER Timeout
    Messagebox show timeout in seconds. After that it will autoclose
.EXAMPLE
    PS C:\> Show-MessageBox -Message 'My Message' -Title 'MB title' -Buttons = 'YesNo' -Icon Information
    SHow Messagebox with message, title, Yes/No buttons and Information icon
.EXAMPLE
    PS C:\> Show-MessageBox -Message 'Everything Lost !' -Title 'This is the end' -Icon Exclamation -Timeout 10
    SHow Messagebox with message, title, Exclamation icon with 10 sec timeout
.OUTPUTS
    System.Windows.Forms.DialogResult. If timed out, return No
.LINK
    https://stackoverflow.com/a/26418199
    https://docs.microsoft.com/ru-ru/dotnet/api/system.threading.tasks.task.delay?view=netframework-4.5    
.NOTES
    Author: Max Kozlov
    Idea from stack overflow
    Required Net 4.5+
    TODO: set focus on messagebox window
#>
function Show-MessageBox {
[CmdletBinding()]
param(
    [Parameter(Mandatory, Position=0)]
    [string]$Message,
    [Parameter(Position=1)]
    [string]$Title = '',
    [ValidateSet('OK','OKCancel','AbortRetryIgnore','YesNoCancel','YesNo','RetryCancel')]
    [Parameter(Position=2)]
    [string]$Buttons = 'OK',
    [ValidateSet('None','Hand','Error','Stop','Question','Exclamation','Warning','Asterisk','Information')]
    [Parameter(Position=3)]
    [string]$Icon = 'None',
    [Parameter(Position=4)]
    [int]$Timeout = 0
)
    Add-Type -Assembly System.Windows.Forms
    $w = $null
    if ($Timeout) {
        $cancel = New-Object System.Threading.CancellationTokenSource
        $w = New-Object System.Windows.Forms.Form
        $w.Size = New-Object System.Drawing.Size (0,0)
        [System.Action[System.Threading.Tasks.Task]]$action = {
            param($t)
            Write-Debug "Want to Close $($task.Status)"
            $w.Close()
        }
        $task = [System.Threading.Tasks.Task]::Delay(
            [timespan]::FromSeconds($Timeout), $cancel.Token
        ).ContinueWith($action,
            [System.Threading.Tasks.TaskScheduler]::FromCurrentSynchronizationContext()
        )
        Write-Debug "Before $($task.Status)"
    }
    #$w.TopMost = $true
    [System.Windows.Forms.MessageBox]::Show($w, $Message, $Title, $Buttons, $Icon)
    if ($Timeout) {
        Write-Debug "After $($task.Status)"
        if ($task.Status -ne 'RanToCompletion') {
            Write-Debug "Do Cancel"
            $cancel.Cancel()
            $task.Wait()
            $cancel.Dispose()
        }
        $task.Dispose()
    }
}
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'Fiddler Trace (*.saz)|*.saz'
}
$null = $FileBrowser.ShowDialog()

[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$title = 'Fiddler Password'
$msg   = 'What is the Fiddler file password?'

$unpackingPassword = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title) 

$guid=[guid]::NewGuid().guid
New-Item -Path $env:TEMP\$guid -ItemType Directory |Out-Null
Expand-7Zip -ArchiveFileName $Filebrowser.Filename -TargetPath "$env:TEMP\$guid" -Password $unpackingPassword




$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select Diagnostic'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please select a diagnostic:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

#[void] $listBox.Items.Add('NSPI')
[void] $listBox.Items.Add('Fiddler HTTPS Decryption')
#[void] $listBox.Items.Add('atl-dc-003')
#[void] $listBox.Items.Add('atl-dc-004')
#[void] $listBox.Items.Add('atl-dc-005')
#[void] $listBox.Items.Add('atl-dc-006')
#[void] $listBox.Items.Add('atl-dc-007')

$form.Controls.Add($listBox)

#$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $listBox.SelectedItem

    if ($x -eq "NSPI")
    {
    Select-String -Path "$env:TEMP\$guid\raw\*.txt" -Pattern "Error" -List | Out-GridView


}
if ($x -eq "Fiddler HTTPS Decryption")
    {
   $Decrypted= Select-String -Path "$env:TEMP\$guid\*.htm" -Pattern "HTTPS" -List 
 
   if ($Decrypted -eq $null)
   {
   
   Show-MessageBox -Message 'ERROR: Fiddler trace was NOT properly collected.' -Title 'FiddlerChecker' -Buttons 'OK' -Icon Information
   }

   else

   {
   Show-MessageBox -Message 'OK: Fiddler trace was properly collected.' -Title 'FiddlerChecker' -Buttons 'OK' -Icon Information
   }


}

}
Remove-Item "$env:TEMP\$guid" -Recurse
Read-Host "Press any key to exit"

