Imports Newtonsoft.Json

Public Class LoginResponse
    <JsonProperty("data")>
    Public data As LoginResponseData
    <JsonProperty("error")>
    Public _error As RaiseError
    <JsonProperty("status")>
    Public status As Long
End Class

Public Class LoginResponseData
    <JsonProperty("token")>
    Public token As String
End Class
