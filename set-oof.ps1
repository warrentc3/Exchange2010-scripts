## Quick and dirty set out of office
## Version 1.0

Function Test-Credential { 
    [OutputType([Bool])] 
## sourced from https://dotps1.github.io
param (
    [Parameter(
        Mandatory = $true
    )]
    [Alias(
        'PSCredential'
    )]
    [ValidateNotNull()]
    [PSCredential]
    $Credential,

    [Parameter()]
    [String]
    $Domain = $Credential.GetNetworkCredential().Domain
)

[System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") |
    Out-Null

$principalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext(
    [System.DirectoryServices.AccountManagement.ContextType]::Domain, $Domain
)

$networkCredential = $Credential.GetNetworkCredential()

Write-Output -InputObject $(
    $principalContext.ValidateCredentials(
        $networkCredential.UserName, $networkCredential.Password
    )
)

$principalContext.Dispose()
}


## Form designed with help from poshgui.com
Add-Type -AssemblyName System.Windows.Forms

if($creds = $host.ui.PromptForCredential('Need credentials', 'Please enter your user name and password.','', "")){}else{exit}
$credtest = Test-Credential -Credential $creds

while ($credtest -eq $false) 
    {
    return [System.Windows.Forms.MessageBox]::Show("Incorrect username or password.")
    if($creds = $host.ui.PromptForCredential('Need credentials', 'Please enter your user name and password.','', "")){}else{exit}
    $credtest = Test-Credential -Credential $creds
    }


Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Set Out Of Office"
$Form.TopMost = $true
$Form.Width = 500
$Form.Height = 400

$email = New-Object system.windows.Forms.TextBox
$email.Text = "enter here"
$email.BackColor = "#ffffff"
$email.ForeColor = "#000000"
$email.Width = 250
$email.Height = 20
$email.location = new-object system.drawing.point(35,47)
$email.Font = "Lucida Console,10"
$Form.controls.Add($email)

$oofmessage = New-Object system.windows.Forms.RichTextBox
$oofmessage.Text = "enter here"
$oofmessage.Width = 400
$oofmessage.Height = 200
$oofmessage.location = new-object system.drawing.point(35,105)
$oofmessage.Font = "Lucida Console,10"
$Form.controls.Add($oofmessage)

$emaillabel = New-Object system.windows.Forms.Label
$emaillabel.Text = "E-Mail Address:"
$emaillabel.AutoSize = $true
$emaillabel.Width = 25
$emaillabel.Height = 10
$emaillabel.location = new-object system.drawing.point(25,24)
$emaillabel.Font = "Lucida Console,10"
$Form.controls.Add($emaillabel)

$labeloof = New-Object system.windows.Forms.Label
$labeloof.Text = "Out Of Office Message:"
$labeloof.AutoSize = $true
$labeloof.Width = 25
$labeloof.Height = 10
$labeloof.location = new-object system.drawing.point(25,85)
$labeloof.Font = "Lucida Console,10"
$Form.controls.Add($labeloof)

$button6 = New-Object system.windows.Forms.Button
$button6.BackColor = "#2b00ff"
$button6.Text = "Submit"
$button6.ForeColor = "#ffffff"
$button6.Width = 90
$button6.Height = 30
$button6.Add_MouseClick({
$user = $email.Text
$oof = $oofmessage.Text
$max = New-PSSessionOption -MaximumRedirection 1
Invoke-Command -ConfigurationName Microsoft.Exchange -ConnectionUri https://mailserver.com/PowerShell -Authentication Kerberos -Credential $creds -ScriptBlock {param($user,$oof);Set-MailboxAutoReplyConfiguration -Identity $user -AutoReplyState Enabled -InternalMessage $oof -ExternalMessage $oof -ExternalAudience All} -ArgumentList $user,$oof -AllowRedirection -SessionOption $max
return [System.Windows.Forms.MessageBox]::Show("Out Of Office message has been set for $user.")
$oofmessage.Text = " "
$email.Text = " "
})
$button6.location = new-object system.drawing.point(200,320)
$button6.Font = "Lucida Console,10"
$Form.controls.Add($button6)

[void]$Form.ShowDialog()
$Form.Dispose()
