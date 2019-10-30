# Mark all emails in certain folders as read and purge non-inbox messages older than 7 days

# Get Oulook folders

function recurse_folders {
    Param($top_folder,$folders_to_mark)
    ForEach ($folder in $top_folder.Folders) {
        if ($folders_to_mark -contains $folder.Name.ToString()) {
        $unread_emails = @($folder.Items.Restrict("[unread] = true"))
            Write-Host ("Marking " + $folder.Name + " as read...")
            For ($i = 0; $i -lt $unread_emails.Count; $i++) {
                 $unread_emails[$i].UnRead = $false
                
            }
        }
        recurse_folders $folder $folders_to_mark
    }

}

$outlook_object = New-Object -ComObject Outlook.Application

$inbox = $outlook_object.Application.GetNamespace("MAPI").GetDefaultFolder(6)

# List of folders to mark as read
$folders_to_mark = @("Folder1","folder2","FOLDER 3")

recurse_folders $inbox $folders_to_mark

