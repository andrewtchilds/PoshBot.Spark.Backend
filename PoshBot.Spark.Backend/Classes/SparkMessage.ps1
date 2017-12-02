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
