Function Export-MicrosoftTodo {
    param(
        $exportFilename = "microsoft_todo_backup.xml"
    )
    $ErrorActionPreference = "Stop"
    # Disabling Progress Bar for better Invoke-WebRequest performance
    $ProgressPreference = 'SilentlyContinue'
    $graphBaseUri = "https://graph.microsoft.com/beta"

    # Ask for token
    $accessToken = Read-Host -AsSecureString -Prompt "Paste OAuth2 Token"

    # User information
    $me = Invoke-RestMethod -Uri ($graphBaseUri + "/me") -Authentication OAuth -Token $accessToken
    "Backing up Microsoft To Do data..."
    "User: $($me.displayName) / $($me.userPrincipalName)"
    "Output file: $exportFilename"

    # List lists - https://docs.microsoft.com/en-us/graph/api/todo-list-lists
    "Getting lists..."
    $lists = Invoke-RestMethod -Uri ($graphBaseUri + "/me/todo/lists") -Authentication OAuth -Token $accessToken | 
        Select-Object -ExpandProperty value
    "Got $($lists.count) lists."

    # List tasks - https://docs.microsoft.com/en-us/graph/api/todotasklist-list-tasks
    $tasks = @()
    foreach ($list in $lists) {
        "Getting tasks in list: $($list.displayName)..."
        $results = @()
        # /me/todo/lists/{todoTaskListId}/tasks
        $uri = ($graphBaseUri + "/me/todo/lists/" + $list.id + "/tasks")
        while ($uri) {
            $response = Invoke-WebRequest -Uri $uri -Authentication OAuth -Token $accessToken
            # Invoke-RestMethod / ConvertFrom-Json mangles the response, which I resent,
            # so we're using an alternative parser and storing the JSON as a string
            # https://stackoverflow.com/a/58169326/12055271
            $json = [Newtonsoft.Json.JsonConvert]::DeserializeObject($response.Content)
            foreach ($task in $json.value) {
                # Don't need ID - Graph API can generate a new one
                $task.Remove("id") | Out-Null
                # Don't need ETag
                $task.Remove("@odata.etag") | Out-Null
                $results += $task.ToString()
            }
            # ConvertFrom-JSON can be trusted to look for next page link
            $uri = ($response.Content | ConvertFrom-Json).'@odata.nextLink'
            # Loop if there's another page
        }
        "Got $($results.Count) tasks."
        $tasks += [PSCustomObject]@{
            "list_id" = $list.id
            "tasks"   = $results
        }
    }
    "Total tasks: $($tasks.tasks.Count)"

    "Exporting to XML..."
    [PSCustomObject]@{
        "about" = [PSCustomObject]@{
            "displayName"   = $me.displayName
            "UPN"           = $me.userPrincipalName
            "backupCreated" = Get-Date
            "scriptVersion" = "0.1"
        }
        "lists" = $lists
        "tasks" = $tasks
    } | Export-Clixml -Path $exportFilename -Verbose
}


Function Import-MicrosoftTodo {
    param(
        $importFilename = "microsoft_todo_backup.xml"
    )
    $ErrorActionPreference = "Stop"
    $graphBaseUri = "https://graph.microsoft.com/beta"

    # Load XML and ask for token
    $sourceData = Import-Clixml -Path $importFilename
    $accessToken = Read-Host -AsSecureString -Prompt "Paste OAuth2 Token"

    "Loaded $importFilename. Backup details:"
    "List count: $($sourceData.lists.Count)"
    "Task count: $($sourceData.tasks.tasks.Count)"
    $sourceData.about | Format-List

    $me = Invoke-RestMethod -Uri ($graphBaseUri + "/me") -Authentication OAuth -Token $accessToken
    "Restoring to user: $($me.displayName) / $($me.userPrincipalName)"

    # Old school. Should use ConfirmPreference etc.
    $confirm = Read-Host "Continue? (Y/N)"
    if ($confirm -ne "y") {
        "Exiting."
        break
    }

    #region Create any missing lists in target account
    $targetLists = Invoke-RestMethod -Uri ($graphBaseUri + "/me/todo/lists") -Authentication OAuth -Token $accessToken | 
        Select-Object -ExpandProperty value
    "Got $($targetLists.Count) lists at target."

    $toCreate = Compare-Object $targetLists $sourceData.lists -Property displayName |
        Where-Object { $_.SideIndicator -eq "=>" }
    "Need to create $($toCreate.Count) lists."

    if ($toCreate) {
        foreach ($list in $toCreate) {
            # Create todoTaskList - https://docs.microsoft.com/en-us/graph/api/todo-post-lists
            $params = @{
                "Method"         = "Post"
                "Uri"            = ($graphBaseUri + "/me/todo/lists")
                "Authentication" = "OAuth"
                "Token"          = $accessToken
                "Body"           = @{
                    "displayName" = $list.displayName
                } | ConvertTo-Json
                # utf-8 makes emojis work. Life priorities are correct.
                "ContentType"    = "application/json; charset=utf-8"
            }
            Invoke-RestMethod @params
        }
        # Get fresh copy of target lists
        $targetLists = Invoke-RestMethod -Uri ($graphBaseUri + "/me/todo/lists") -Authentication OAuth -Token $accessToken | 
            Select-Object -ExpandProperty value
    }
    ""
    #endregion

    #region Add tasks
    foreach ($group in $sourceData.tasks) {
        # XML tasks are grouped by list
        # Lookup displayName. Important that this is case sensitive
        $listDisplayName = $sourceData.lists | Where-Object { $_.id -ceq $group.list_id } | 
            Select-Object -ExpandProperty displayName
        "Processing list: $listDisplayName..."
        $taskCount = $group.tasks.Count
        "Tasks to add: $taskCount"
        # Match with target list
        $targetListId = $targetLists | Where-Object { $_.displayName -eq $listDisplayName } | 
            Select-Object -ExpandProperty id
        $i = 0
        foreach ($task in $group.tasks) {
            $i++
            $progressTask = ($task | ConvertFrom-Json).title
            Write-Progress -Activity "Adding tasks to $listDisplayName" -CurrentOperation $progressTask -PercentComplete ($i / $taskCount * 100)
            # Create todoTask - https://docs.microsoft.com/en-us/graph/api/todotasklist-post-tasks
            # Nested for loops go brrr...
            $params = @{
                # POST /me/todo/lists/{todoTaskListId}/tasks
                "Method"            = "Post"
                "Uri"               = ($graphBaseUri + "/me/todo/lists/" + $targetListId + "/tasks")
                "Authentication"    = "OAuth"
                "Token"             = $accessToken
                "Body"              = $task
                "ContentType"       = "application/json; charset=utf-8"
                "MaximumRetryCount" = 2
                "RetryIntervalSec"  = 5
            }
            Invoke-RestMethod @params | Out-Null
        }
    }
    #endregion
    "Finished!"
}