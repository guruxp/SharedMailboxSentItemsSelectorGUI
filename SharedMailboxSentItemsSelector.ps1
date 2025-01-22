#Connect to Exchange Online
Connect-ExchangeOnline

# Load the required assemblies
Add-Type -AssemblyName PresentationFramework

# WPF Window
$window = New-Object Windows.Window
$window.Title = "Shared mailbox sent items management"
$window.Width = 600
$window.Height = 550

# Buttons
$Checkbutton = New-Object Windows.Controls.Button
$Checkbutton.Content = "1 - Check connection"
$Checkbutton.Width = 200
$Checkbutton.Height = 50
$Checkbutton.Margin = "50,30,50,30"

$GetStatebutton = New-Object Windows.Controls.Button
$GetStatebutton.Content = "3 - Get sent items state"
$GetStatebutton.Width = 200
$GetStatebutton.Height = 50
$GetStatebutton.Margin = "50,20,50,20"

$FlipSendAsValueButton = New-Object Windows.Controls.Button
$FlipSendAsValueButton.Content = "Change SendAs value"
$FlipSendAsValueButton.Width = 200
$FlipSendAsValueButton.Height = 50
$FlipSendAsValueButton.Margin = "5,10,5,20"
$FlipSendAsValueButton.IsEnabled = $false

$FlipSendOnBehalfValueButton = New-Object Windows.Controls.Button
$FlipSendOnBehalfValueButton.Content = "Change SendOnBehalf value"
$FlipSendOnBehalfValueButton.Width = 200
$FlipSendOnBehalfValueButton.Height = 50
$FlipSendOnBehalfValueButton.Margin = "5,10,5,20"
$FlipSendOnBehalfValueButton.IsEnabled = $false

# Text fields
$SharedMailbox = New-Object Windows.Controls.TextBox
$SharedMailbox.Text = "INPUTEXISTINGMAILBOXHERE"
$SharedMailbox.Width = 200
$SharedMailbox.Height = 25
$SharedMailbox.Margin = "50,5,50,30"

# Labels
$SharedMailboxSendAsState = New-Object Windows.Controls.Label
$SharedMailboxSendAsState.Width = 300
$SharedMailboxSendAsState.Height = 50
$SharedMailboxSendAsState.Margin = "50,5,50,5"

$MailboxSMTPLabel = New-Object Windows.Controls.Label
$MailboxSMTPLabel.Width = 200
$MailboxSMTPLabel.Height = 50
$MailboxSMTPLabel.Margin = "50,10,50,5"
$MailboxSMTPLabel.Content = "2 - Specify the SMTP address `n     of the shared mailbox:"

# Stack panel
$stackPanel = New-Object Windows.Controls.StackPanel

# Add the button to the window
$stackPanel.Children.Add($Checkbutton)
$stackPanel.Children.Add($MailboxSMTPLabel)
$stackPanel.Children.Add($SharedMailbox)
$stackPanel.Children.Add($GetStatebutton)
$stackPanel.Children.Add($SharedMailboxSendAsState)

# Create a horizontal stack panel for the buttons
$buttonPanel = New-Object Windows.Controls.StackPanel
$buttonPanel.Orientation = "Horizontal"
$buttonPanel.HorizontalAlignment = "Center"
$buttonPanel.Children.Add($FlipSendAsValueButton)
$buttonPanel.Children.Add($FlipSendOnBehalfValueButton)

# Add the horizontal stack panel to the main stack panel
$stackPanel.Children.Add($buttonPanel)

$window.Content = $stackPanel

# Buttons behavior
$Checkbutton.Add_Click({
if (Get-Mailbox INPUTEXISTINGMAILBOXHERE -ErrorAction SilentlyContinue){
    $Checkbutton.Content = "Success"
    $Checkbutton.Background = "LightGreen"
} else {
    $Checkbutton.Content = "Fail"
    $Checkbutton.Background = "LightRed"
}
})

$GetStatebutton.Add_Click({
 if (Get-Mailbox $SharedMailbox.Text -RecipientTypeDetails SharedMailbox -ErrorAction SilentlyContinue){
    $SMBXSendAsValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSentAsEnabled
    $SMBXSOBOValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSendOnBehalfEnabled
    $SharedMailboxSendAsState.Content = "SendAs: " + $SMBXSendAsValue  + "`nSendOnBehalfOf: " + $SMBXSOBOValue
    $FlipSendAsValueButton.IsEnabled = $true     
    $FlipSendOnBehalfValueButton.IsEnabled = $true

} else {
    $FlipSendAsValueButton.IsEnabled = $false     
    $FlipSendOnBehalfValueButton.IsEnabled = $false
    $SharedMailboxSendAsState.Content = "Not allowed: This is not a Shared Mailbox"
}
})

$FlipSendOnBehalfValueButton.Add_Click({
$SMBXSOBOValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSendOnBehalfEnabled
$SMBXSendAsValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSentAsEnabled
if ($SMBXSOBOValue -eq 'True'){
    Set-Mailbox $SharedMailbox.Text -MessageCopyForSendOnBehalfEnabled:$False
    $SMBXSOBOValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSendOnBehalfEnabled
    $SharedMailboxSendAsState.Content = "SendAs: " + $SMBXSendAsValue  + "`nSendOnBehalfOf: " + $SMBXSOBOValue
} else {
    Set-Mailbox $SharedMailbox.Text -MessageCopyForSendOnBehalfEnabled:$True
    $SMBXSOBOValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSendOnBehalfEnabled
    $SharedMailboxSendAsState.Content = "SendAs: " + $SMBXSendAsValue  + "`nSendOnBehalfOf: " + $SMBXSOBOValue
}
})

$FlipSendAsValueButton.Add_Click({
$SMBXSendAsValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSentAsEnabled
$SMBXSOBOValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSendOnBehalfEnabled
if ($SMBXSendAsValue -eq 'True') {
    $SMBXSendAsValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSentAsEnabled
    Set-Mailbox $SharedMailbox.Text -MessageCopyForSentAsEnabled:$False
    $SMBXSendAsValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSentAsEnabled
    $SMBXSOBOValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSendOnBehalfEnabled
    $SharedMailboxSendAsState.Content = "SendAs: " + $SMBXSendAsValue  + "`nSendOnBehalfOf: " + $SMBXSOBOValue
} else {
    Set-Mailbox $SharedMailbox.Text -MessageCopyForSentAsEnabled:$True
    $SMBXSendAsValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSentAsEnabled
    $SMBXSOBOValue = (Get-Mailbox $SharedMailbox.Text).MessageCopyForSendOnBehalfEnabled
    $SharedMailboxSendAsState.Content = "SendAs: " + $SMBXSendAsValue  + "`nSendOnBehalfOf: " + $SMBXSOBOValue
}
})
# Show the window
$window.Background = "LightGray"
$window.ShowDialog()