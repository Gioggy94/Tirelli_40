Imports Newtonsoft.Json

Public Class RaiseError
    <JsonProperty("code")>
    Public code As Long
    <JsonProperty("msg")>
    Public msg As String
End Class
