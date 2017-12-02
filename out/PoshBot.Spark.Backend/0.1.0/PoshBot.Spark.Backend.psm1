using module PoshBot

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '', Scope='Class', Target='*')]
class SparkBackend : Backend {

    # The types of message that we care about from Spark
    # All othere will be ignored
    [string[]]$MessageTypes = @("conversation.activity")

    SparkBackend ([string]$Token) {
        Import-Module PSSpark -Verbose:$false -ErrorAction Stop

        $config = [ConnectionConfig]::new()
        $secToken = $Token | ConvertTo-SecureString -AsPlainText -Force
        $config.Credential = New-Object System.Management.Automation.PSCredential('asdf', $secToken)
        $conn = [SparkConnection]::New()
        $conn.Config = $config
        $this.Connection = $conn
    }

    # Connect to Spark
    [void]Connect() {
        $this.LogInfo('Connecting to backend')
        $this.LogInfo('Listening for the following message types. All others will be ignored', $this.MessageTypes)
        $this.Connection.Connect()
        $this.BotId = $this.GetBotIdentity()
    }

    # Receive a message from the websocket
    [Message[]]ReceiveMessage() {
        $messages = New-Object -TypeName System.Collections.ArrayList
        try {
            # Read the output stream from the receive job and get any messages since our last read
            $jsonResult = $this.Connection.ReadReceiveJob()

            if($null -ne $jsonResult -and $jsonResult -ne [string]::Empty) {
                #Write-Debug -Message "[SparkBackend:ReceiveMessage] Received `n$jsonResult"
                $this.LogDebug('Received message', $jsonResult)

                $sparkMessages = @($jsonResult | ConvertFrom-Json)
                foreach($sparkMessage in $sparkMessages) {

                    # We only care about certain message types from Spark
                    if($sparkMessage.data.eventType -in $this.MessageTypes) {
                        $msg = [Message]::new()

                        # Set the message type and optionally the subtype
                        switch($sparkMessage.data.eventType) {
                            'conversation.activity' {
                                $msg.Type = [MessageType]::Message
                                $sparkMessage = Get-SparkMessage -MessageID $sparkMessage.data.activity.id
                            }
                        }

                        $this.LogDebug("Message type is [$($msg.Type)`:$($msg.Subtype)]")

                        if($sparkMessage.Text)     { $msg.Text = $sparkMessage.Text }
                        if($sparkMessage.RoomID)   { $msg.To   = $sparkMessage.RoomID }
                        if($sparkMessage.UserID)   { $msg.From = $sparkMessage.UserID }

                        $processed = $this._ProcessMentions($msg.Text)
                        $msg.Text = $processed

                        if ($msg.From) {
                            $msg.FromName = $this.UserIdToUsername($msg.From)
                        }

                        # Resolve channel name
                        if ($msg.To -and $msg.To -notmatch '^D') {
                            $msg.ToName = $this.ChannelIdToName($msg.To)
                        }

                        if($sparkMessage.Created) {
                            $msg.Time = $sparkMessage.Created
                        } else {
                            $msg.Time = (Get-Date).ToUniversalTime()
                        }

                        # ** Important safety tip, don't cross the streams **
                        # Only return messages that didn't come from the bot
                        # else we'd cause a feedback loop with the bot processing
                        # it's own responses
                        if(-not $this.MsgFromBot($msg.From)) {
                            $messages.Add($msg) > $null
                        }
                    } else {
                        $this.LogDebug("Message type is [$($sparkMessage.Type)]. Ignoring")
                    }
                }
            }
        } catch {
            Write-Error $_
        }

        return $messages
    }

    # Send a Slack ping
    [void]Ping() {

    }

    # Send a message back to Slack
    [void]SendMessage([Response]$Response) {
        if ($Response.Data.Text.Count -gt 0) {
            foreach ($t in $Response.Data.Text) {
                $t = "```````n" + $t
                $this.LogDebug("Sending response back to Spark channel [$($Response.To)]", $t)
                Send-SparkMessage -Token $this.Connection.Config.Credential.GetNetworkCredential().Password -RoomID $Response.To -MarkdownText $t -Verbose:$false
            }
        }
        
    }


    # Add a reaction to an existing chat message
    [void]AddReaction([Message]$Message, [ReactionType]$Type, [string]$Reaction) {

    }

    # Remove a reaction from an existing chat message
    [void]RemoveReaction([Message]$Message, [ReactionType]$Type, [string]$Reaction) {

    }

    # Resolve a channel name to an Id
    [string]ResolveChannelId([string]$ChannelName) {
        if ($ChannelName -match '^#') {
            $ChannelName = $ChannelName.TrimStart('#')
        }

        $channelId = Get-SparkRoom -Name $ChannelName | Select-Object -ExpandProperty RoomID
        $this.LogDebug("Resolved channel [$ChannelName] to [$channelId]")
        return $channelId
    }

    # Populate the list of users the Slack team
    [void]LoadUsers() {
    
    }

    # Populate the list of channels in the Slack team
    [void]LoadRooms() {
        
    }

    # Get the bot identity Id
    [string]GetBotIdentity() {
        $id = $this.Connection.LoginData.userId
        $id = Get-SparkUser -UserID $id | Select-Object -ExpandProperty UserID
        $this.LogVerbose("Bot identity is [$id]")
        return $id
    }

    # Determine if incoming message was from the bot
    [bool]MsgFromBot([string]$From) {
        $frombot = ($this.BotId -eq $From)
        if ($fromBot) {
            $this.LogDebug("Message is from bot [From: $From == Bot: $($this.BotId)]. Ignoring")
        } else {
            $this.LogDebug("Message is not from bot [From: $From <> Bot: $($this.BotId)]")
        }
        return $fromBot
    }

    # Get a user by their Id
    [SparkPerson]GetUser([string]$UserId) {
        $user = Get-SparkUser -UserID $UserId

        $person = [SparkPerson]::new()
        $person.Id = $user.UserID
        $person.Nickname = $user.NickName
        $person.FullName = $user.Name
        $person.FirstName = $user.FirstName
        $person.LastName = $user.LastName
        $person.Email = $user.Email
        $person.Type = $user.Type
        $person.Status = $user.Status
        $person.Avatar = $user.Avatar
        $person.LastActivity = $user.LastActivity
        $person.Created = $user.Created

        return $person
    }

    # Get a user Id by their name
    [string]UsernameToUserId([string]$Username) {
        $Username = $Username.TrimStart('@')
        $id = Get-SparkUser -Name $Username | Select-Object -ExpandProperty UserID
        return $id
    }

    # Get a user name by their Id
    [string]UserIdToUsername([string]$UserId) {
        $name = $null
        $name = Get-SparkUser -UserID $UserId | Select-Object -ExpandProperty Name
        return $name
    }

    # Get the channel name by Id
    [string]ChannelIdToName([string]$ChannelId) {
        $name = $null
        $name = Get-SparkRoom -RoomID $ChannelId | Select-Object -ExpandProperty Name
        return $name
    }

    # Remove extra characters that Slack decorates urls with
    hidden [string] _SanitizeURIs([string]$Text) {
        $sanitizedText = $Text -replace '<([^\|>]+)\|([^\|>]+)>', '$2'
        $sanitizedText = $sanitizedText -replace '<(http([^>]+))>', '$1'
        return $sanitizedText
    }

    # Break apart a string by number of characters
    hidden [System.Collections.ArrayList] _ChunkString([string]$Text) {
        $chunks = [regex]::Split($Text, "(?<=\G.{$($this.MaxMessageLength)})", [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $this.LogDebug("Split response into [$($chunks.Count)] chunks")
        return $chunks
    }

    # Resolve a reaction type to an emoji
    hidden [string]_ResolveEmoji([ReactionType]$Type) {
        $emoji = [string]::Empty
        Switch ($Type) {
            'Success'        { return 'U+2714' }
            'Failure'        { return 'U+2757' }
            'Processing'     { return 'U+2699' }
            'Warning'        { return 'U+26A0' }
            'ApprovalNeeded' { return 'U+1F510' }
            'Cancelled'      { return 'U+1F6AB' }
            'Denied'         { return 'U+274C' }
        }
        return $emoji
    }

    hidden [string]_UnicodeToString([string]$UnicodeChars) {
        $UnicodeChars = $UnicodeChars -replace 'U\+', '';
    
        $UnicodeArray = @();
        foreach ($UnicodeChar in $UnicodeChars.Split(' ')) {
            $Int = [System.Convert]::ToInt32($UnicodeChar, 16);
            $UnicodeArray += [System.Char]::ConvertFromUtf32($Int);
        }
    
        return $UnicodeArray -join [String]::Empty;
    }
    
    # Strips bot username from text
    hidden [string]_ProcessMentions([string]$Text) {

        $processed = ($Text -split "^sparky\s")[1]

        return $processed
    }
}
class SparkConnection : Connection {

    [System.Net.WebSockets.ClientWebSocket]$WebSocket
    [pscustomobject]$LoginData
    [string]$UserName
    [string]$WebSocketUrl
    [bool]$Connected
    [object]$ReceiveJob = $null

    SparkConnection() {
        $this.WebSocket = New-Object System.Net.WebSockets.ClientWebSocket
    }

    # Connect to Spark and start receiving messages
    [void]Connect() {
        if($null -eq $this.ReceiveJob -or $this.ReceiveJob.State -ne 'Running') {
            $this.LogDebug('Connecting to Spark Real Time API')
            $this.RtmConnect()
            $this.StartReceiveJob()
        } else {
            $this.LogDebug([LogSeverity]::Warning, 'Receive job is already running')
        }
    }

    # Log in to Spark with the bot token and get a URL to connect to via websockets
    [void]RtmConnect() {
        $token = $this.Config.Credential.GetNetworkCredential().Password
        $url = New-SparkWebSocket -Token $token | Select-Object -ExpandProperty url
        $headers = @{ "Authorization" = "Bearer $token" }
        try {
            $r = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -Verbose:$false
            $this.LoginData = $r
            if($r) {
                $this.LogInfo('Successfully authenticated to Spark Real Time API')
                $this.WebSocketUrl = $r.webSocketUrl
                $this.UserName = Get-SparkUser -UserID $r.userId | Select-Object -ExpandProperty Name
            } else {
                throw $r
            }
        } catch {
            $this.LogInfo([LogSeverity]::Error, 'Error connecting to Spark Real Time API', [ExceptionFormatter]::Summarize($_))
        }
    }

    # Setup the websocket receive job
    [void]StartReceiveJob() {
        $recv = {
            [cmdletbinding()]
            param(
                [parameter(mandatory)]
                $url,
                [parameter(mandatory)]
                $token
            )

            # Connect to websocket
            Write-Verbose "[SparkBackend:ReceiveJob] Connecting to websocket at [$($url)]"
            [System.Net.WebSockets.ClientWebSocket]$webSocket = New-Object System.Net.WebSockets.ClientWebSocket
            $ct = New-Object System.Threading.CancellationToken
            $task = $webSocket.ConnectAsync($url, $ct)
            $buffer = (New-Object System.Byte[] 4096)
            $taskResult = $null
            
            while(-not $task.IsCompleted) { Start-Sleep -Milliseconds 100 }
            
            $Body = @{
                id = [guid]::NewGuid().guid
                type = "authorization"
                data = @{
                    token = "Bearer $Token"
                }
            } | ConvertTo-Json
            
            $Array = @()
            $Body.ToCharArray() | ForEach { $Array += [byte]$_ }
            $Body = New-Object System.ArraySegment[byte]  -ArgumentList @(,$Array)
            
            $Conn = $webSocket.SendAsync($Body, [System.Net.WebSockets.WebSocketMessageType]::Text, [System.Boolean]::TrueString, $ct)
            
            while(-not $Conn.IsCompleted) { Start-Sleep -Milliseconds 100 }
            
            # Receive messages and put on output stream so the backend can read them
            
            while($webSocket.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
                do {
                    $taskResult = $webSocket.ReceiveAsync($buffer, $ct)
                    while(-not $taskResult.IsCompleted) { Start-Sleep -Milliseconds 100 }
                } until($taskResult.Result.Count -lt 4096)

                $jsonResult = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $taskResult.Result.Count)
            
                if(-not [string]::IsNullOrEmpty($jsonResult)) {
                    $jsonResult
                }
            }
            $socketStatus = [pscustomobject]@{
                State = $webSocket.State
                CloseStatus = $webSocket.CloseStatus
                CloseStatusDescription = $webSocket.CloseStatusDescription
            }
            $socketStatusStr = ($socketStatus | Format-List | Out-String).Trim()
            Write-Warning -Message "Websocket state is [$($webSocket.State.ToString())].`n$socketStatusStr"
        }
        try {
            $this.ReceiveJob = Start-Job -Name ReceiveRtmMessages -ScriptBlock $recv -ArgumentList $this.WebSocketUrl,$this.Config.Credential.GetNetworkCredential().Password -ErrorAction Stop -Verbose
            $this.Connected = $true
            $this.Status = [ConnectionStatus]::Connected
            $this.LogInfo("Started websocket receive job [$($this.ReceiveJob.Id)]")
        } catch {
            $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
        }
    }

    # Read all available data from the job
    [string]ReadReceiveJob() {
        # Read stream info from the job so we can log them
        $infoStream = $this.ReceiveJob.ChildJobs[0].Information.ReadAll()
        $warningStream = $this.ReceiveJob.ChildJobs[0].Warning.ReadAll()
        $errStream = $this.ReceiveJob.ChildJobs[0].Error.ReadAll()
        $verboseStream = $this.ReceiveJob.ChildJobs[0].Verbose.ReadAll()
        $debugStream = $this.ReceiveJob.ChildJobs[0].Debug.ReadAll()

        foreach($item in $infoStream) {
            $this.LogInfo($item.ToString())
        }
        foreach($item in $warningStream) {
            $this.LogInfo([LogSeverity]::Warning, $item.ToString())
        }
        foreach($item in $errStream) {
            $this.LogInfo([LogSeverity]::Error, $item.ToString())
        }
        foreach($item in $verboseStream) {
            $this.LogVerbose($item.ToString())
        }
        foreach($item in $debugStream) {
            $this.LogVerbose($item.ToString())
        }

        # The receive job stopped for some reason. Reestablish the connection if the job isn't running
        if($this.ReceiveJob.State -ne 'Running') {
            $this.LogInfo([LogSeverity]::Warning, "Receive job state is [$($this.ReceiveJob.State)]. Attempting to reconnect...")
            Start-Sleep -Seconds 5
            $this.Connect()
        }

        if($this.ReceiveJob.HasMoreData) {
            return $this.ReceiveJob.ChildJobs[0].Output.ReadAll()
        } else {
            return $null
        }
    }

    # Stop the receive job
    [void]Disconnect() {
        $this.LogInfo('Closing websocket')
        if($this.ReceiveJob) {
            $this.LogInfo("Stopping receive job [$($this.ReceiveJob.Id)]")
            $this.ReceiveJob | Stop-Job -Confirm:$false -PassThru | Remove-Job -Force -ErrorAction SilentlyContinue
        }
        $this.Connected = $false
        $this.Status = [ConnectionStatus]::Disconnected
    }
}
enum SparkMessageType {
    Normal
    Error
    Warning
}

class SparkMessage : Message {

    [SparkMessageType]$MessageType = [SparkMessageType]::Normal

    SparkMessage(
        [string]$To,
        [string]$From,
        [string]$Body = [string]::Empty
    ) {
        $this.To = $To
        $this.From = $From
        $this.Body = $Body
    }

}


class SparkPerson : Person {
    [string]$Email
    [string]$Type
    [string]$Status
    [string]$Avatar
    [string]$Created
    [string]$LastActivity
}

function New-PoshBotSparkBackend {
    <#
    .SYNOPSIS
        Create a new instance of a Spark backend
    .DESCRIPTION
        Create a new instance of a Spark backend
    .PARAMETER Configuration
        The hashtable containing backend-specific properties on how to create the Spark backend instance.
    .EXAMPLE
        PS C:\> $backendConfig = @{Name = 'SparkBackend'; Token = '<SPARK-API-TOKEN>'}
        PS C:\> $backend = New-PoshBotSparkBackend -Configuration $backendConfig

        Create a Spark backend using the specified API token
    .INPUTS
        Hashtable
    .OUTPUTS
        SparkBackend
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='*')]
    [cmdletbinding()]
    param(
        [parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('BackendConfiguration')]
        [hashtable[]]$Configuration
    )

    process {
        foreach($item in $Configuration) {
            if(-not $item.Token) {
                throw 'Configuration is missing [Token] parameter'
            } else {
                Write-Verbose 'Creating new Spark backend instance'
                $backend = [SparkBackend]::new($item.Token)
                if($item.Name) {
                    $backend.Name = $item.Name
                }
                $backend
            }
        }
    }
}

Export-ModuleMember -Function 'New-PoshBotSparkBackend'
