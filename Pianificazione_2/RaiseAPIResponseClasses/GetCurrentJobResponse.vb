Imports Newtonsoft.Json

Public Class GetCurrentJobResponse
    <JsonProperty("data")>
    Public data As GetCurrentJobData
    <JsonProperty("error")>
    Public _error As RaiseError
    <JsonProperty("status")>
    Public status As Long
End Class

Public Class GetCurrentJobData
    <JsonProperty("file_name")>
    Public file_name As String
    <JsonProperty("print_progress")>
    Public print_progress As Double
    <JsonProperty("printed_layer")>
    Public printed_layer As Long
    <JsonProperty("printed_time")>
    Public printed_time As Long
    <JsonProperty("job_id")>
    Public job_id As String
    <JsonProperty("total_layer")>
    Public total_layer As Long
    <JsonProperty("total_time")>
    Public total_time As Long
    <JsonProperty("job_status")>
    Public job_status As String
End Class