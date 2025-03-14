# Create a report to list unread emails sent in the last 24 hours

Add-Type -AssemblyName System.Windows.Forms


# Create Outlook COM Object
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")


# Access the Inbox
$Inbox = $Namespace.GetDefaultFolder(6) # 6 = Inbox


# Load configuration from JSON file
$Config = Get-Content "Outlook_Automation/config.json" | ConvertFrom-Json
$Important_Senders_Names = $Config.important_names


# Get emails received in the last 24 hours
$Last24Hours = (Get-Date).AddHours(-24)


function UnreadEmails {

    # Fetch unread emails sent in the last 24 hours
    $RecentEmails = $Inbox.Items.Restrict("[ReceivedTime] >= '" + $Last24Hours.ToString("g") + "'")
    $UnreadEmails = $RecentEmails | Where-Object { $_.UnRead -eq $true }


    # Sort emails based on user importance
    $SortedEmails = $UnreadEmails | Sort-Object { if ($Important_Senders_Names -contains $_.SenderName) { 0 } else { 1 } }
    

    # Create GUI Window
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Unread Email Report"
    $Form.Size = New-Object System.Drawing.Size(1500, 550)


    # Create a table
    $DataGridView = New-Object System.Windows.Forms.DataGridView
    $DataGridView.Size = New-Object System.Drawing.Size(1450, 400)
    $DataGridView.Location = New-Object System.Drawing.Point(10, 10)
    $DataGridView.AutoSizeColumnsMode = "Fill"
    $DataGridView.ColumnCount = 3
    $DataGridView.Columns[0].Name = "Received Time"
    $DataGridView.Columns[1].Name = "Sender"
    $DataGridView.Columns[2].Name = "Subject"


    # Populate the table with email data
    foreach ($Email in $SortedEmails) {
    $rowIndex = $DataGridView.Rows.Add($Email.ReceivedTime, $Email.SenderName, $Email.Subject)
    

    # Highlight emails sent by important senders
        if ($Important_Senders_Names -contains $Email.SenderName) {

            $DataGridView.Rows[$rowIndex].DefaultCellStyle.BackColor = "LightYellow"
            $DataGridView.Rows[$rowIndex].DefaultCellStyle.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

        }
    }

    # Create OK button to close the window
    $Button = New-Object System.Windows.Forms.Button
    $Button.Text = "OK"
    $Button.Size = New-Object System.Drawing.Size(80, 30)
    $Button.Location = New-Object System.Drawing.Point(710, 440)
    $Button.Add_Click({ $Form.Close() })
        
    # Add components to the form
    $Form.Controls.Add($DataGridView)
    $Form.Controls.Add($Button)

    # Show the window
    $Form.ShowDialog()

}

UnreadEmails

