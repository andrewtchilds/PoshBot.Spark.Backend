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