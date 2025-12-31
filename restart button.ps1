##example
function Run-Checks {
        # your code to test ports, firewall, services
        Write-Host "Running checks..."}
        # initial 
        runRun-Checks
        # refresh loop
        do {
            $choice = Read-Host "Press R to refresh or Q to quit"
            if ($choice -eq "R") { Run-Checks }
        } while ($choice -ne "Q")

## example2
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object Windows.Forms.Form
$form.Text = "Admin Toolbox"
$form.Size = '300,200'
$button = New-Object Windows.Forms.Button
$button.Text = "Refresh"
$button.Location = '100,50'
$button.Add_Click({    Run-Checks   # call your function again
    })
$form.Controls.Add($button)
$form.ShowDialog()