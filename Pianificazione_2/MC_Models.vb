Public Class MC_Macchina
    Public Property ID As Integer
    Public Property Matricola As String = ""
    Public Property NomeMacchina As String = ""
    Public Property Modello As String = ""
    Public Property TipoMacchina As String = ""
    Public Property ClienteFinale As String = ""
    Public Property AnnoCostruzione As Integer?
    Public Property LinguaCodice As String = "IT"
    Public Property Attiva As Boolean = True
    Public Property Note As String = ""
    Public Property PesoKg As Double?
    Public Property ConsumoAria As Double?
    Public Property Corrente As String = ""
    Public Property Tensione As String = ""
    Public Property DataCreazione As DateTime
    Public Property DataModifica As DateTime?

    Public Overrides Function ToString() As String
        Return $"{Matricola} – {NomeMacchina}"
    End Function
End Class

Public Class MC_Fotocellula
    Public Property ID As Integer
    Public Property MacchinaID As Integer
    Public Property CatalogoID As Integer
    ' Campi derivati dal JOIN con MC_CatalogoFotocellule (solo lettura)
    Public Property Codice As String = ""
    Public Property TipoNome As String = ""
    Public Property PathImmagine As String = ""
    Public Property DataCreazione As DateTime
End Class

Public Class MC_TipoFotocellula
    Public Property ID As Integer
    Public Property Nome As String = ""
    Public Overrides Function ToString() As String
        Return Nome
    End Function
End Class

Public Class MC_CatalogoFotocellula
    Public Property ID As Integer
    Public Property Codice As String = ""
    Public Property TipoID As Integer
    Public Property TipoNome As String = ""
    Public Property PathImmagine As String = ""
    Public Property DataCreazione As DateTime
    Public Overrides Function ToString() As String
        Return $"{Codice} ({TipoNome})"
    End Function
End Class

Public Class MC_CodiceErrore
    Public Property ID As Integer
    Public Property MacchinaID As Integer
    Public Property Codice As String = ""
    Public Property Titolo As String = ""
    Public Property Descrizione As String = ""
    Public Property Causa As String = ""
    Public Property Rimedio As String = ""
    Public Property Gravita As String = "Avviso"
    Public Property NomeScreenshot As String = ""
    Public Property PathScreenshot As String = ""
    Public Property DataCreazione As DateTime
End Class

Public Class MC_Lingua
    Public Property Codice As String = ""
    Public Property Nome As String = ""
    Public Property Attiva As Boolean = True
    Public Overrides Function ToString() As String
        Return $"{Nome} ({Codice})"
    End Function
End Class

Public Class MC_Modello
    Public Property ID As Integer
    Public Property Nome As String = ""
    Public Overrides Function ToString() As String
        Return Nome
    End Function
End Class

Public Class MC_TipoMacchina
    Public Property ID As Integer
    Public Property Nome As String = ""
    Public Overrides Function ToString() As String
        Return Nome
    End Function
End Class
