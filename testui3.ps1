Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#------------------ Helper Functions ------------------

function Add-LabelTextboxPair {
    param (
        [System.Windows.Forms.Form]$form,
        [string]$labelText,
        [int]$x,
        [int]$y
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Location = New-Object System.Drawing.Point($x, $y)
    $label.Size = New-Object System.Drawing.Size(120, 20)
    $form.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point(($x + 130), $y)
    $textbox.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($textbox)

    return $textbox
}

function Add-AccessDropdown {
    param (
        [System.Windows.Forms.Form]$form,
        [int]$x,
        [int]$y
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Access Level:"
    $label.Location = New-Object System.Drawing.Point($x, $y)
    $label.Size = New-Object System.Drawing.Size(120, 20)
    $form.Controls.Add($label)

    $dropdown = New-Object System.Windows.Forms.ComboBox
    $dropdown.Location = New-Object System.Drawing.Point(($x + 130), $y)
    $dropdown.Size = New-Object System.Drawing.Size(200, 20)
    $dropdown.DropDownStyle = "DropDownList"
    $dropdown.Items.AddRange(@("None", "Reviewer", "Author", "Editor", "Owner"))
    $form.Controls.Add($dropdown)

    $tooltip = New-Object System.Windows.Forms.ToolTip
    $tooltip.SetToolTip($dropdown, "None = No access, Reviewer = View only, Author = Add items, Editor = Full edit, Owner = Full control")

    return $dropdown
}

function Write-AuditLog {
    param (
        [string]$action,
        [string]$target,
        [string]$user,
        [string]$access
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $log = "$timestamp | $action | Target: $target | User: $user | Access: $access"
    Add-Content -Path "$env:USERPROFILE\CalendarPermissionAudit.log" -Value $log
}

#------------------ UI Functions ------------------

function UserToUserUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "User to User Calendar Access"
    $form.Size = New-Object System.Drawing.Size(400, 300)
    $form.StartPosition = "CenterScreen"

    $ownerBox = Add-LabelTextboxPair -form $form -labelText "Calendar Owner:" -x 20 -y 30
    $userBox = Add-LabelTextboxPair -form $form -labelText "Accessing User:" -x 20 -y 70
    $accessDropdown = Add-AccessDropdown -form $form -x 20 -y 110

    $btnSubmit = New-Object System.Windows.Forms.Button
    $btnSubmit.Text = "Apply"
    $btnSubmit.Location = New-Object System.Drawing.Point(150, 160)
    $btnSubmit.Size = New-Object System.Drawing.Size(80, 30)
    $form.Controls.Add($btnSubmit)

    $btnSubmit.Add_Click({
        $owner = $ownerBox.Text
        $user = $userBox.Text
        $access = $accessDropdown.SelectedItem

        if (-not $owner -or -not $user -or -not $access) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all fields.")
            return
        }

        try {
            Get-Mailbox -Identity $owner | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Mailbox '$owner' does not exist.")
            return
        }

        try {
            Add-MailboxFolderPermission -Identity ($owner + ":\Calendar") -User $user -AccessRights $access
            Write-AuditLog -action "UserToUser" -target $owner -user $user -access $access
            [System.Windows.Forms.MessageBox]::Show("Permission applied successfully.")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)")
        }
    })

    $form.ShowDialog()
}

function GroupToMailboxUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Group to Mailbox Calendar Access"
    $form.Size = New-Object System.Drawing.Size(400, 250)
    $form.StartPosition = "CenterScreen"

    $groupBox = Add-LabelTextboxPair -form $form -labelText "Distribution Group:" -x 20 -y 30
    $mailboxBox = Add-LabelTextboxPair -form $form -labelText "Mailbox:" -x 20 -y 70

    $btnSubmit = New-Object System.Windows.Forms.Button
    $btnSubmit.Text = "Apply"
    $btnSubmit.Location = New-Object System.Drawing.Point(150, 120)
    $btnSubmit.Size = New-Object System.Drawing.Size(80, 30)
    $form.Controls.Add($btnSubmit)

    $btnSubmit.Add_Click({
        $group = $groupBox.Text
        $mailbox = $mailboxBox.Text

        if (-not $group -or -not $mailbox) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all fields.")
            return
        }

        try {
            Get-Mailbox -Identity $mailbox | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Mailbox '$mailbox' does not exist.")
            return
        }

        try {
            $users = Get-DistributionGroupMember -Identity $group
            foreach ($user in $users) {
                Add-MailboxFolderPermission -Identity ($mailbox + ":\Calendar") -User $user.PrimarySmtpAddress -AccessRights Editor
                Write-AuditLog -action "GroupToMailbox" -target $mailbox -user $user.PrimarySmtpAddress -access "Editor"
            }
            [System.Windows.Forms.MessageBox]::Show("Permissions applied to group members.")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)")
        }
    })

    $form.ShowDialog()
}

function EveryoneToMailboxUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Everyone to Single Mailbox Calendar"
    $form.Size = New-Object System.Drawing.Size(400, 220)
    $form.StartPosition = "CenterScreen"

    $targetBox = Add-LabelTextboxPair -form $form -labelText "Target User:" -x 20 -y 30
    $accessDropdown = Add-AccessDropdown -form $form -x 20 -y 70

    $btnSubmit = New-Object System.Windows.Forms.Button
    $btnSubmit.Text = "Apply"
    $btnSubmit.Location = New-Object System.Drawing.Point(150, 120)
    $btnSubmit.Size = New-Object System.Drawing.Size(80, 30)
    $form.Controls.Add($btnSubmit)

    $btnSubmit.Add_Click({
        $target = $targetBox.Text
        $access = $accessDropdown.SelectedItem

        if (-not $target -or -not $access) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all fields.")
            return
        }

        try {
            Get-Mailbox -Identity $target | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Mailbox '$target' does not exist.")
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show("Apply permissions to all users?", "Confirm", "YesNo")
        if ($confirm -ne "Yes") { return }

        try {
            $users = Get-Mailbox | Select -ExpandProperty Alias
            foreach ($user in $users) {
                Add-MailboxFolderPermission -Identity ($user + ":\Calendar") -User $target -AccessRights $access
                Write-AuditLog -action "EveryoneToMailbox" -target $user -user $target -access $access
            }
            [System.Windows.Forms.MessageBox]::Show("Permissions applied to all users.")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)")
        }
    })

    $form.ShowDialog()
}

function CheckCalendarPermissionsUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Check Calendar Permissions"
    $form.Size = New-Object System.Drawing.Size(650, 500)
    $form.StartPosition = "CenterScreen"

    $mailboxBox = Add-LabelTextboxPair -form $form -labelText "Mailbox Alias:" -x 20 -y 30
    $searchBox = Add-LabelTextboxPair -form $form -labelText "Search Filter:" -x 20 -y 70

    $btnCheck = New-Object System.Windows.Forms.Button
    $btnCheck.Text = "Check"
    $btnCheck.Location = New-Object System.Drawing.Point(150, 110)
    $btnCheck.Size = New-Object System.Drawing.Size(80, 30)
    $form.Controls.Add($btnCheck)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = "Search"
    $btnSearch.Location = New-Object System.Drawing.Point(250, 110)
    $btnSearch.Size = New-Object System.Drawing.Size(80, 30)
    $form.Controls.Add($btnSearch)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "Export"
    $btnExport.Location = New-Object System.Drawing.Point(350, 110)
    $btnExport.Size = New-Object System.Drawing.Size(80, 30)
    $form.Controls.Add($btnExport)

    $outputBox = New-Object System.Windows.Forms.TextBox
    $outputBox.Multiline = $true
    $outputBox.ScrollBars = "Vertical"
    $outputBox.Location = New-Object System.Drawing.Point(20, 160)
    $outputBox.Size = New-Object System.Drawing.Size(600, 280)
    $outputBox.ReadOnly = $true
    $form.Controls.Add($outputBox)

    $global:cachedPermissions = $null

    $btnCheck.Add_Click({
        $mailbox = $mailboxBox.Text
        if (-not $mailbox) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a mailbox alias.")
            return
        }

        try {
            Get-Mailbox -Identity $mailbox | Out-Null
        } catch {
            $outputBox.Text = "Mailbox '$mailbox' does not exist."
            return
        }

        try {
            $calendar = $mailbox + ":\Calendar"
            $global:cachedPermissions = Get-MailboxFolderPermission -Identity $calendar
            $output = $cachedPermissions | Select Name, User, AccessRights | Format-Table -AutoSize | Out-String
            $outputBox.Text = $output
        } catch {
            $outputBox.Text = "Error: $($_.Exception.Message)"
        }
    })

    $btnSearch.Add_Click({
        $filter = $searchBox.Text
        if (-not $global:cachedPermissions) {
            [System.Windows.Forms.MessageBox]::Show("Please run 'Check' first.")
            return
        }

        $filtered = $global:cachedPermissions | Where-Object {
            $_.User -like "*$filter*" -or $_.AccessRights -like "*$filter*" -or $_.Name -like "*$filter*"
        }

        if ($filtered.Count -eq 0) {
            $outputBox.Text = "No results match '$filter'."
        } else {
            $outputBox.Text = $filtered | Select Name, User, AccessRights | Format-Table -AutoSize | Out-String
        }
    })

    $btnExport.Add_Click({
        $mailbox = $mailboxBox.Text
        if (-not $mailbox) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a mailbox alias before exporting.")
            return
        }

        $path = "$env:USERPROFILE\CalendarPermissions_$mailbox.txt"
        $outputBox.Text | Out-File -FilePath $path -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Exported to $path")
    })

    $form.ShowDialog()
}

function Show-Form {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Multi Calendar Permissions Tool"
    $form.Size = New-Object System.Drawing.Size(500, 220)
    $form.StartPosition = "CenterScreen"

    $labelOp = New-Object System.Windows.Forms.Label
    $labelOp.Text = "Select Operation:"
    $labelOp.Location = New-Object System.Drawing.Point(20, 20)
    $labelOp.Size = New-Object System.Drawing.Size(120, 20)
    $form.Controls.Add($labelOp)

    $comboOp = New-Object System.Windows.Forms.ComboBox
    $comboOp.Location = New-Object System.Drawing.Point(150, 20)
    $comboOp.Size = New-Object System.Drawing.Size(300, 20)
    $comboOp.DropDownStyle = "DropDownList"
    $comboOp.Items.AddRange(@(
        "User to User Calendar Access",
        "Group to Mailbox Calendar Access",
        "Everyone to Single Mailbox Calendar",
        "Check Calendar Permissions",
        "Exit"
    ))
    $form.Controls.Add($comboOp)

    $btnGo = New-Object System.Windows.Forms.Button
    $btnGo.Text = "Go"
    $btnGo.Location = New-Object System.Drawing.Point(200, 60)
    $btnGo.Size = New-Object System.Drawing.Size(80, 30)
    $form.Controls.Add($btnGo)

    $btnGo.Add_Click({
        switch ($comboOp.SelectedIndex) {
            0 { UserToUserUI }
            1 { GroupToMailboxUI }
            2 { EveryoneToMailboxUI }
            3 { CheckCalendarPermissionsUI }
            4 {
                Disconnect-ExchangeOnline -Confirm:$false
                $form.Close()
            }
        }
    })

    $form.ShowDialog()
}

#------------------ Launch ------------------
Connect-ExchangeOnline
Show-Form
Disconnect-ExchangeOnline -Confirm:$false