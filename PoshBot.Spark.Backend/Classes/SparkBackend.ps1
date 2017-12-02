
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