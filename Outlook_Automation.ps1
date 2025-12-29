# Outlook Automation - Create rules to set high importance to emails sent by specific users and move emails to "important" and "Work" folders

# Outlook COM Object
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")


# Load configurations
$Config = Get-Content "Outlook_Automation/config.json" | ConvertFrom-Json

$Important_Senders_Mails = $Config.important_mails
$Important_Senders_Names = $Config.important_names
$Keywords = $Config.keywords

# Access the inbox folder
$Inbox = $Namespace.GetDefaultFolder(6)


# Get time
$TimeNow = Get-Date 
$TimeYesterday = $TimeNow.AddHours(-24)


# Get emails send in the last 24 hours
$RecentEmails = $Inbox.Items.Restrict("[ReceivedTime] >= '" + $TimeYesterday.ToString("g") + "'")


# Set high importance to emails with specific keywords
function Mark-Important {
    
    if ($RecentEmails.Count -gt 0) {
        
        $RecentEmails | ForEach-Object {

            if ($Important_Senders_Mails -contains $_.SenderEmailAddress -or 
               ($Important_Senders_Names -contains $_.SenderName) -or
               ($Keywords -match $_.Subject)) {
                
                $_.Importance = 2
                $_.Save()

            }
        }
    } 
    
    else {

        Write-Host "No emails found in Inbox."
    }
}


# Move necessary emails to important folder
function MoveTo-SubFolder-Important {

    $FolderName = "important"
    $SubFolder = $Inbox.Folders | Where-Object { $_.Name -eq $FolderName}

    if (-not $SubFolder) {

    $SubFolder = $Inbox.Folders.add($FolderName)

    } 
    
    # add any other important keywords
    $RecentEmails | Where-Object {$_.Subject -match "Reminder" -or $_.Subject -match "urgent" -or $_.Subject -match "deadline"} | ForEach-Object {

        Write-Host $_.SenderName 
        $CopiedEmail = $_.Copy()
        $CopiedEmail.Move($SubFolder)
        Write-Host "Moved to important successfully"
    
    }
} 


# Move emails send by work staff to "Work" folder
function MoveTo-SubFolder-Work {

    # added place holders 
    $FolderName = "...Work..."
    $SubFolder = $Inbox.Folders | Where-Object { $_.Name -eq $FolderName}

    if (-not $SubFolder) {

    $SubFolder = $Inbox.Folders.add($FolderName)
    Write-Host "folder created"

    } 
   
    # added placeholders for important senders and places. Replace with real names or get from the config file
    $RecentEmails | Where-Object {$_.SenderName -match "...Sender 1..." -or $_.SenderName -match "...Important Sender 2..." -or $_.Subject -match "...Workplace..."} | ForEach-Object {

    Write-Host $_.SenderName 
    $CopiedEmail = $_.Copy()
    $CopiedEmail.Move($SubFolder)
    Write-Host "Moved to ... Center successfully"
    
    }  
} 

Mark-Important
MoveTo-SubFolder-Important
MoveTo-SubFolder-Work