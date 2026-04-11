Imports System.Data.SqlClient
Imports System.Windows.Documents

Public Class Opportunità_aggiungi_riga
    Public n_opportunità As Integer
    Private max_linea As Integer
    Private codice_OWNER(1000) As String
    Private codice_type_ARRAY(100) As String
    Private codice_addetto_ARRAY(1000) As String
    Private ultima_riga_opportunità As Integer
    Private codice_documento_array(100) As String
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Sub inizializza_opportunità_aggiungi_riga()

    End Sub

    Sub inserisci_combobox_addetto_vendite()


        ComboBox1.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "select t0.slpcode,t0.SlpName
from oslp t0
where t0.active='Y'
order by t0.SlpName
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            codice_addetto_ARRAY(Indice) = cmd_SAP_reader("slpcode")
            ComboBox1.Items.Add(cmd_SAP_reader("SlpName"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub

    Sub inserisci_combobox_type()
        ComboBox2.Items.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.stepid, t0.num, T0.Descript, T0.Canceled, T0.SalesStage, T0.PurStage FROM OOST T0 "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Dim Indice As Integer
        Indice = 0


        Do While cmd_SAP_reader_2.Read()


            If cmd_SAP_reader_2("num") = 2 Or cmd_SAP_reader_2("num") = 3 Or cmd_SAP_reader_2("num") = 5 Or cmd_SAP_reader_2("num") = 8 Or cmd_SAP_reader_2("num") = 14 Then
                ComboBox2.Items.Add(cmd_SAP_reader_2("Descript"))
                codice_type_ARRAY(Indice) = cmd_SAP_reader_2("num")

            ElseIf cmd_SAP_reader_2("num") = 7 Or cmd_SAP_reader_2("num") = 9 Or cmd_SAP_reader_2("num") = 12 Or cmd_SAP_reader_2("num") = 15 Then
                ComboBox2.Items.Add(cmd_SAP_reader_2("Descript"))
                codice_type_ARRAY(Indice) = cmd_SAP_reader_2("num")

            End If
            Indice = Indice + 1

        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Sub inserisci_combobox_owner()
        ComboBox4.Items.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.empid, T0.[lastName]+' ' +T0.[firstName] as 'Compilatore' 
FROM [TIRELLI_40].[DBO].OHEM T0 "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Dim Indice As Integer
        Indice = 0


        Do While cmd_SAP_reader_2.Read()

            ComboBox4.Items.Add(cmd_SAP_reader_2("compilatore"))
            codice_OWNER(Indice) = cmd_SAP_reader_2("empid")
            Indice = Indice + 1
        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Sub inserisci_informazioni()
        ComboBox5.Items.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT code, name
from [@informazioni] "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader_2.Read()

            ComboBox5.Items.Add(cmd_SAP_reader_2("code"))

        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Sub inserisci_priorità()
        ComboBox6.Items.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT *
from [@priorita] "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        ComboBox6.Items.Add("")


        Do While cmd_SAP_reader_2.Read()

            ComboBox6.Items.Add(cmd_SAP_reader_2("code"))

        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Sub inserisci_layout()
        ComboBox7.Items.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT *
from [@layout] "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        ComboBox7.Items.Add("")


        Do While cmd_SAP_reader_2.Read()

            ComboBox7.Items.Add(cmd_SAP_reader_2("code"))

        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Sub inserisci_combobox_tipo_documento()
        Dim Indice As Integer
        Indice = 0
        Dim codice_documento(100) As String

        ComboBox3.Items.Add("")
        codice_documento_array(Indice) = "-1"
        Indice = Indice + 1

        ComboBox3.Items.Add("Offerta")
        codice_documento_array(Indice) = "23"

        Indice = Indice + 1

        ComboBox3.Items.Add("Ordine di acquisto")
        codice_documento_array(Indice) = "22"

        Indice = Indice + 1

        ComboBox3.Items.Add("Fattura di vendita")
        codice_documento_array(Indice) = "13"

        Indice = Indice + 1

        ComboBox3.Items.Add("Ordine cliente")
        codice_documento_array(Indice) = "17"

        Indice = Indice + 1

    End Sub

    Private Sub Opportunità_aggiungi_riga_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inserisci_combobox_addetto_vendite()
        inserisci_combobox_tipo_documento()
        inserisci_layout()
        inserisci_priorità()
        inserisci_informazioni()
        inserisci_combobox_owner()
        inserisci_combobox_type()
    End Sub

    Sub aggiUNGI_RIGA(par_n_opportunità As Integer)
        Trova_riga(n_opportunità)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = cnn
        Cmd_SAP.CommandText = "
INSERT INTO OPR1 ( [OpprId]
      ,[Line]
      ,[SlpCode]
      ,[CntctCode]
      ,[OpenDate]
      ,[CloseDate]
      ,[Step_Id]
      ,[ClosePrcnt]
      ,[MaxSumLoc]
      ,[MaxSumSys]
      ,[Memo]
      ,[DocId]
      ,[ObjType]
      ,[Status]
      ,[Linked]
      ,[WtSumLoc]
      ,[WtSumSys]
      ,[UserSign]
      ,[ChnCrdCode]
      ,[ChnCrdName]
      ,[ChnCrdCon]
      ,[DocChkbox]
      ,[Owner]
      ,[DocNumber]
      ,[EncryptIV]
      ,[U_PRG_AZS_NoteLivOPP]
      ,[U_Centrodicosto]
      ,[U_Informazioni]
      ,[U_Priorita]
      ,[U_Layout]
      ,[U_Data_richiesta])

VALUES
('" & par_n_opportunità & "' 
,'" & max_linea & "' +1
,'" & codice_addetto_ARRAY(ComboBox1.SelectedIndex) & "' 
,NULL
,GETDATE()
,GETDATE()+1
,'" & codice_type_ARRAY(ComboBox2.SelectedIndex) & "' 
,5
,1
,1
,NULL
,NULL
,'" & codice_documento_array(ComboBox3.SelectedIndex) & "'
,'O'
,'N'
,0.5
,0.5
,NULL
,NULL
,NULL
,NULL
,NULL
,'" & codice_OWNER(ComboBox4.SelectedIndex) & "' 
,'" & TextBox2.Text & "'
,NULL
,'" & TextBox1.Text & "'
,NULL
,'" & ComboBox5.Text & "' 
,'" & ComboBox6.Text & "'
,'" & ComboBox7.Text & "' 
,NULL
)

"
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()

    End Sub

    Sub Trova_riga(par_n_opportunità As Integer)
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "Select max(line)+1 As 'linea' from opr1 where opprid=" & par_n_opportunità & ""

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("linea") Is System.DBNull.Value Then
                max_linea = cmd_SAP_reader_2("linea")
            Else
                max_linea = 0
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.SelectedIndex < 0 Then
            MsgBox("Selezionare un addetto alle vendite")
        Else
            If ComboBox4.SelectedIndex < 0 Then
                MsgBox("Selezionare un titolare della riga")
            Else
                If ComboBox2.SelectedIndex < 0 Then
                    MsgBox("Selezionare un type")
                Else
                    If ComboBox5.SelectedIndex < 0 Then
                        MsgBox("Selezionare se le informazioni sono complete")

                    Else
                        aggiUNGI_RIGA(n_opportunità)
                        chiudi_righe_precedenti(n_opportunità)
                        allinea_importo_potenziale(n_opportunità)
                        MsgBox("Riga inserita con successo")
                        Opportunità_N.livello_opportunità(n_opportunità)
                        Me.Close()
                    End If
                End If
            End If

        End If


    End Sub

    Sub chiudi_righe_precedenti(par_n_opportunità As Integer)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = cnn
        Cmd_SAP.CommandText = "UPDATE OPR1 SET STATUS='C', CLOSEDATE=GETDATE() WHERE LINE <=" & max_linea & " AND STATUS='O' AND OPPRID=" & par_n_opportunità & "
"
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()

    End Sub

    Sub allinea_importo_potenziale(par_n_opportunità As Integer)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = cnn
        Cmd_SAP.CommandText = "update t12 set t12.maxsumloc=t11.maxsumloc, t12.cloprcnt=t11.closeprcnt, t12.wtsumloc=t11.wtsumloc,t12.wtsumsys=t11.wtsumsys,t12.maxsumsys=t11.maxsumsys
from
(
SELECT max(T0.[Line]) as 'MAX' , t0.opprid FROM OPR1 T0 WHERE T0.[OpprId] ='" & par_n_opportunità & "' group by t0.opprid
)
as t10 inner join opr1 t11 on t10.opprid=t11.opprid and  t10.max=t11.line
inner join oopr t12 on t12.opprid=t11.opprid
"
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub
End Class