Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib


Imports System.Windows.Controls
Imports System.Reflection.Emit
Imports System.Security.Cryptography
Imports Newtonsoft.Json
Imports System.Net
Imports Microsoft.Office.Interop.Word
Imports MailMessage = System.Net.Mail.MailMessage
Imports Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.Windows.Forms
Imports System.Text


Imports System.Messaging
Imports System

Imports System.Windows.Media.Media3D


Public Class Presence

    Public Num_Mail As Integer
    Public num_ore As Integer
    Public num_registrazioni As Integer

    Public Elenco_Mail(100) As String
    Public Elenco_Notifiche(100) As Integer
    Public Elenco_Registrazione_Risorse(100) As Integer
    Public ora_lav As String



    Public Anno_Attuale As Integer
    Public Mese_Attuale As Integer
    Public Giorno_Attuale As Integer
    Public Anno As Integer
    Public Mese As String
    Public Giorno As String
    Public Timbrature_Oggi(4000) As Timbrata
    Public Dipendenti(1000) As Dipendente
    Public Num_Timbrate As Integer = 0
    Public Num_Dipendenti As Integer = 0
    Public Indice As Integer
    Public objStreamReader As StreamReader
    Public strLine As String
    Public id As Integer
    Public ListNumber As String
    Public docentry_RT As Integer
    Public lineNum_RT As Integer
    Public QUANTITY As Decimal
    Public id_list_result As Integer



    Public Testo_Mail As String

    Public destinatario1 As String
    Public destinatario2 As String = "giovanni.tirelli@tirelli.net"
    Public distinta As String
    Public code As String


    Public PROGRAMMA As String
    Public stato_programma As String
    Public stato_letto_programma As String
    Public check_inviato As String

    Public Structure Timbrata
        Public Codice_Dipendente As Integer
        Public Tipo As Integer
        Public Ore As Integer
        Public Minuti As Integer
        Public Secondi As Integer
    End Structure

    Public Structure Dipendente
        Public Codice_Dipendente As Integer
        Public Nome_Dipendente As String
    End Structure

    Public Sub dati_presence()

        Num_Mail = 0
        num_ore = 0
        num_registrazioni = 0





        Elenco_Notifiche(num_ore) = Val(23) * 60 + Val(30)
        num_ore = num_ore + 1

        Elenco_Registrazione_Risorse(num_registrazioni) = Val(22) * 60 + Val(29)
        num_registrazioni = num_registrazioni + 1

        ora_lav = Val(21) * 60 + Val(28)

    End Sub

    Sub scrivi_lista()





        'Lettura Elenco Dipendenti

        trova_dipendenti()

        'Lettura Timbratura nella Data Selezionata
        'Anno_Attuale = Year(Calendario.SelectionStart)
        Anno_Attuale = Year(Today)
        Mese_Attuale = Month(Today)
        Giorno_Attuale = DateAndTime.Day(Today)

        objStreamReader = New StreamReader("TRANSACTIONS.TXT")
        strLine = objStreamReader.ReadLine
        Do While Not strLine Is Nothing
            Anno = Val(Mid(strLine, 1, 4))
            Mese = Val(Mid(strLine, 5, 2))
            Giorno = Val(Mid(strLine, 7, 2))

            If Anno = Anno_Attuale And Mese = Mese_Attuale And Giorno = Giorno_Attuale Then


                If Len(Giorno) = 1 Then
                    Giorno = "0" & Giorno
                End If
                If Len(Mese) = 1 Then
                    Mese = "0" & Mese
                End If


                Timbrature_Oggi(Num_Timbrate).Ore = Val(Mid(strLine, 10, 2))
                Timbrature_Oggi(Num_Timbrate).Minuti = Val(Mid(strLine, 12, 2))
                Timbrature_Oggi(Num_Timbrate).Secondi = Val(Mid(strLine, 14, 2))
                Timbrature_Oggi(Num_Timbrate).Tipo = Val(Mid(strLine, 17, 1))
                Timbrature_Oggi(Num_Timbrate).Codice_Dipendente = Val(Mid(strLine, 20, 6))
                inserisci_timbrature_in_sap()
                Num_Timbrate = Num_Timbrate + 1
            End If
            strLine = objStreamReader.ReadLine
        Loop
        objStreamReader.Close()




        Dim objStreamWriter As StreamWriter
        objStreamWriter = New StreamWriter("Output.txt")
        For Indice = 0 To Num_Dipendenti - 1
            Dim Stringa As String
            Dim Num_Occorrenze As Integer = 0

            Stringa = Dipendenti(Indice).Nome_Dipendente
            For Tim As Integer = 0 To Num_Timbrate
                If Timbrature_Oggi(Tim).Codice_Dipendente = Dipendenti(Indice).Codice_Dipendente Then
                    Stringa = Stringa + " " + Timbrature_Oggi(Tim).Ore.ToString("D2") + ":" + Timbrature_Oggi(Tim).Minuti.ToString("D2") + ":" + Timbrature_Oggi(Tim).Secondi.ToString("D2")
                    If Timbrature_Oggi(Tim).Tipo = 0 Then
                        Stringa = Stringa + "(OUT)"
                    Else
                        Stringa = Stringa + "(IN)"
                    End If
                    Num_Occorrenze = Num_Occorrenze + 1
                End If
            Next Tim
            If Num_Occorrenze = 0 Then
                Stringa = Stringa + " Assente"
            End If

            Dim Esito As String
            If Num_Occorrenze < 4 Then
                Esito = "[ANOMALIA] "
            Else
                Esito = "[   OK   ] "
            End If


            objStreamWriter.WriteLine(Esito + Stringa)
        Next Indice

        objStreamWriter.Close()
        'Btn_Stampa.Enabled = True
    End Sub

    Private Sub Presence_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Timer_revisioni.Start()
        ' Form_mail_ricambi.Show()
        ' Form_mail_ricambi.WindowState = FormWindowState.Minimized
        '  Stampante_3D.risposta_stampante()
        ' Timer_stampante.Start()
        ' Timer_giornaliero.Start()
        ' Timer_Ferretto.Start()
    End Sub


    Public Function CallRaise3d(Of T)(ByVal requestUri As String) As T
        Dim retVal As T = Nothing

        Dim httpRequest As HttpWebRequest = WebRequest.CreateHttp(requestUri)
        httpRequest.Method = "GET"

        Using httpResponse As HttpWebResponse = httpRequest.GetResponse()
            Dim responseStream As Stream = httpResponse.GetResponseStream()
            Dim sr As New StreamReader(responseStream)
            Dim result As String = sr.ReadToEnd()
            retVal = JsonConvert.DeserializeObject(Of T)(result)
        End Using

        Return retVal
    End Function


    Sub trova_dipendenti()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] , t0.u_n_badge
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code where t0.active='Y'  and t0.u_n_badge<>'' order by T0.[lastName]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()

            Dipendenti(Num_Dipendenti).Codice_Dipendente = cmd_SAP_reader("u_n_badge")
            Dipendenti(Num_Dipendenti).Nome_Dipendente = cmd_SAP_reader("Nome")
            Num_Dipendenti = Num_Dipendenti + 1
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()


    End Sub

    Sub inserisci_timbrature_in_sap()
        Dim ore_import As String
        Dim minuti_import As String
        Dim secondi_import As String

        If Timbrature_Oggi(Num_Timbrate).Ore <= 9 Then
            ore_import = "0" & Timbrature_Oggi(Num_Timbrate).Ore
        Else
            ore_import = Timbrature_Oggi(Num_Timbrate).Ore

        End If

        If Timbrature_Oggi(Num_Timbrate).Minuti <= 9 Then

            minuti_import = "0" & Timbrature_Oggi(Num_Timbrate).Minuti
        Else
            minuti_import = Timbrature_Oggi(Num_Timbrate).Minuti
        End If

        If Timbrature_Oggi(Num_Timbrate).Secondi <= 9 Then
            secondi_import = "0" & Timbrature_Oggi(Num_Timbrate).Secondi
        Else
            secondi_import = Timbrature_Oggi(Num_Timbrate).Secondi
        End If


        Trova_ID()
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "insert into timbrature (id,id_cartellino,anno,mese,giorno,ore,minuti,secondi,tipo) values (" & id & "," & Timbrature_Oggi(Num_Timbrate).Codice_Dipendente & "," & Anno & ",'" & Mese & "','" & Giorno & "','" & ore_import & "','" & minuti_import & "','" & secondi_import & "','" & Timbrature_Oggi(Num_Timbrate).Tipo & "')"

        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()


    End Sub

    Sub Trova_ID()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select max(id)+1 as 'ID' from timbrature"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                id = cmd_SAP_reader_2("ID")
            Else
                id = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        check_mail_già_inviata("OA")
        If check_inviato = "NO" Then
            invio_mail_prezzi_0("OPOR", "POR1")
            aggiungi_log_prezzi_0_OA("OA")
        Else
            MsgBox("NON OA")
        End If

        check_mail_già_inviata("EM")

        If check_inviato = "EM" Then
            invio_mail_prezzi_0("OPDN", "PDN1")
            aggiungi_log_prezzi_0_OA("EM")
        Else
            MsgBox("NON EM")
        End If

        MsgBox("Fatto")
    End Sub

    Sub trova_list_results_aperte()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT TOP (1000) [id],

 SUBSTRING(
    listnumber, 
    CHARINDEX('-', listnumber) + 1, 
    CHARINDEX('-', listnumber, CHARINDEX('-', listnumber) + 1) - CHARINDEX('-', listnumber) - 1
  ) AS 'Docentry_RT'
      ,[recordStatus]
      ,[recordWritingDate]
      ,[recordImportationDate]
      ,[plantId]
      ,[response]
      ,[listType]
      ,[listNumber]
      ,[lineNumber]
      ,[item]
      ,[batch]
      ,[serialNumber]
      ,[requestedQty]
      ,[processedQty]
      ,[errorCause]
      ,[wmsGenerated]
      ,[auxText01]
      ,[auxText02]
      ,[auxText03]
      ,[auxText04]
      ,[auxText05]
      ,[auxText06]
      ,[auxText07]
      ,[auxText08]
      ,[auxText09]
      ,[auxText10]
      ,[auxInt01]
      ,[auxInt02]
      ,[auxInt03]
      ,[auxInt04]
      ,[auxInt05]
      ,[auxInt06]
      ,[auxInt07]
      ,[auxInt08]
      ,[auxInt09]
      ,[auxInt10]
      ,[auxDate01]
      ,[auxDate02]
      ,[auxDate03]
      ,[auxDate04]
      ,[auxDate05]
      ,[auxDate06]
      ,[auxDate07]
      ,[auxDate08]
      ,[auxDate09]
      ,[auxDate10]
      ,[auxBit01]
      ,[auxBit02]
      ,[auxBit03]
      ,[auxBit04]
      ,[auxBit05]
      ,[auxBit06]
      ,[auxBit07]
      ,[auxBit08]
      ,[auxBit09]
      ,[auxBit10]
      ,[auxNum01]
      ,[auxNum02]
      ,[auxNum03]
      ,[auxNum04]
      ,[auxNum05]
      ,[auxNum06]
      ,[auxNum07]
      ,[auxNum08]
      ,[auxNum09]
      ,[auxNum10]
  FROM [FGWmsErp].[dbo].[LISTS_RESULT]

  where  auxint01='7494' and recordstatus=1
order by id desc
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Dim VARIABILE As Integer = 0

        If cmd_SAP_reader.Read() Then
            id_list_result = cmd_SAP_reader("id")
            ListNumber = cmd_SAP_reader("listNumber")
            docentry_RT = cmd_SAP_reader("Docentry_RT")
            lineNum_RT = cmd_SAP_reader("lineNumber")
            QUANTITY = cmd_SAP_reader("processedQty")
            VARIABILE = 1
        Else
            VARIABILE = 0

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()

        If VARIABILE = 1 Then
            trova_RELATIVA_RT(docentry_RT, lineNum_RT)
        End If


    End Sub

    Sub trova_RELATIVA_RT(par_docentry As Integer, par_linenum As Integer)

        Dim Cnn6 As New SqlConnection
        Cnn6.ConnectionString = Homepage.sap_tirelli
        Cnn6.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn6
        CMD_SAP.CommandText = "select t0.itemcode, t0.OpenQty,t0.FromWhsCod,t0.WhsCode, t0.linestatus,CASE WHEN t1.docnum IS NULL THEN 0 ELSE T1.DOCNUM END AS 'DOCNUM_ODP',CASE WHEN  T0.U_PRG_AZS_OPLINENUM IS NULL THEN 0 ELSE T0.U_PRG_AZS_OPLINENUM END AS 'U_PRG_AZS_OPLINENUM'
, CASE WHEN T1.U_PRG_AZS_COMMESSA IS NULL THEN '' ELSE T1.U_PRG_AZS_COMMESSA END AS 'U_PRG_AZS_COMMESSA', t2.usersign
from wtq1 t0 left join owor t1 on t1.docentry=t0.u_prg_azs_opdocentry
LEFT JOIN OWTQ T2 ON T2.DOCENTRY=T0.DOCENTRY
where t0.docentry='" & par_docentry & "' and t0.linenum='" & par_linenum & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim VARIABILE_1 As Integer = 0

        If cmd_SAP_reader.Read() Then

            If cmd_SAP_reader("linestatus") = "O" Then

                If QUANTITY <= cmd_SAP_reader("OpenQty") Then
                    VARIABILE_1 = 1
                    Form_Entrate_Merci.Trasferimento_in_WIP("ODP", cmd_SAP_reader("itemcode"), cmd_SAP_reader("docnum_ODP"), 0, Replace(QUANTITY, ",", "."), cmd_SAP_reader("FromWhsCod"), cmd_SAP_reader("WHSCODE"), cmd_SAP_reader("U_PRG_AZS_OPLINENUM"), cmd_SAP_reader("U_PRG_AZS_COMMESSA"), cmd_SAP_reader("USERSIGN"), "Automatico", docentry_RT, 0, "Trasferimento")
                    CHIUDI_LIST_RESULTS(id_list_result, ListNumber)
                    CHIUDI_riga_rt(docentry_RT, lineNum_RT, Magazzino.DOCENTRY_Trasferimenti())
                Else
                    VARIABILE_1 = 0
                End If
            Else
                VARIABILE_1 = 0
            End If
        Else
            VARIABILE_1 = 0
        End If
        cmd_SAP_reader.Close()
        Cnn6.Close()


    End Sub

    Private Sub Timer_Ferretto_Tick(sender As Object, e As EventArgs) Handles Timer_Ferretto.Tick
        'trova_list_results_aperte()
    End Sub

    Sub CHIUDI_LIST_RESULTS(par_id_list_result As Integer, par_list_number As String)



        Dim Cnn2 As New SqlConnection
        Cnn2.ConnectionString = Homepage.sap_tirelli
        Cnn2.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn2

        CMD_SAP.CommandText = "UPDATE [FGWmsErp].[dbo].[LISTS_RESULT] set recordstatus ='2', recordimportationdate=getdate() WHERE ID ='" & par_id_list_result & "' AND LISTNUMBER='" & par_list_number & "' "
        CMD_SAP.ExecuteNonQuery()
        Cnn2.Close()




    End Sub

    Sub CHIUDI_riga_rt(par_docentry_rt As Integer, par_linenum_rt As Integer, par_docentry_trasferimento As Integer)

        Dim Cnn2 As New SqlConnection
        Cnn2.ConnectionString = Homepage.sap_tirelli
        Cnn2.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn2

        CMD_SAP.CommandText = "UPDATE wtq1 set linestatus='C', TrgetEntry='" & par_docentry_trasferimento & "' WHERE docentry=" & par_docentry_rt & " and linenum = " & par_linenum_rt & ""
        CMD_SAP.ExecuteNonQuery()
        Cnn2.Close()


    End Sub

    Private Sub Timer_revisioni_Tick(sender As Object, e As EventArgs) Handles Timer_revisioni.Tick
        'Try

        '    invio_mail_errore_ivsaplink()
        'Catch ex As Exception

        'End Try



        'invio_mail_pezzo_revisionato_agli_acquisti()
        'invio_mail_pezzo_revisionato_int()

        'check_mail_già_inviata("OA")


        'Timer_revisioni.Interval = 1000000
    End Sub

    Sub invio_mail_errore_ivsaplink()



        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT CODE, name , u_data, u_data_ora, u_nome_file, u_art_padre_db, u_errore, u_email ,CHARINDEX ( ',' , u_email, 1 )  as 'mino', case when case when CHARINDEX ( ',' , u_email, 1 )> 0 then substring(u_email,1,CHARINDEX ( ',' , u_email, 1 )-1) else u_email end ='' then 'deniscattabriga@tirelli.net' else case when CHARINDEX ( ',' , u_email, 1 )> 0 then substring(u_email,1,CHARINDEX ( ',' , u_email, 1 )-1) else u_email end end  as 'Destinatario1',
 case when CHARINDEX ( ',' , u_email, 1 )> 0 then substring(u_email,CHARINDEX ( ',' , u_email, 1 )+2,CHARINDEX ( ',' , u_email, 1 )+1000) else '' end as 'Destinatario2'
                                FROM ""@PRG_TIR_INVSAP_LOG""
                            WHERE name is null
                                ORDER BY Code DESC"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()
            Testo_Mail = Testo_Mail & "<table border=5>"
            Testo_Mail = Testo_Mail & "<tr>"

            Testo_Mail = Testo_Mail & "<th>Data e ora</th>"
            Testo_Mail = Testo_Mail & "<th>Codice distinta</th>"
            Testo_Mail = Testo_Mail & "<th>Errore</th>"

            Testo_Mail = Testo_Mail & "</tr>"
            Testo_Mail = Testo_Mail & "<tr>"
            destinatario1 = cmd_SAP_reader("destinatario1")
            If destinatario1 = "jobtirelli@tim.it" Then
                destinatario1 = "denis.cattabriga@tirelli.net"
            End If
            distinta = cmd_SAP_reader("u_art_padre_db")
            destinatario2 = cmd_SAP_reader("destinatario2")
            code = cmd_SAP_reader("code")


            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("u_data_ora") & "</td>"

            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("u_art_padre_db") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("u_errore") & "</td>"



            Testo_Mail = Testo_Mail & "</tr>"
            Testo_Mail = Testo_Mail & "</table>"

            Testo_Mail = Testo_Mail & "</BODY>"


            Dim Indice_Mail As Integer = 0


            Dim mySmtp As New SmtpClient
            Dim myMail As New MailMessage()



            mySmtp.UseDefaultCredentials = False
            mySmtp.Credentials = New Net.NetworkCredential(Homepage.Mittente_Mail, Homepage.password_mail)
            mySmtp.Host = "tirelli-net.mail.protection.outlook.com"
            mySmtp.Port = 25
            mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network
            mySmtp.EnableSsl = True


            myMail = New MailMessage()

            myMail.From = New MailAddress(Homepage.Mittente_Mail)

            myMail.To.Add(destinatario1)
            Try


                myMail.To.Add(destinatario2)

            Catch ex As Exception

            End Try
            myMail.Subject = "Errore distinta " & distinta & " importazione IVSAPLINK del " & Now.Day.ToString & "/" & Now.Month.ToString & "/" & Now.Year.ToString & " Delle Ore : " & Now.Hour & ":" & Now.Minute
            myMail.IsBodyHtml = True
            myMail.Body = Testo_Mail

            Try
                mySmtp.Send(myMail)
            Catch ex As Exception

            End Try

            Indice_Mail = Indice_Mail + 1
            invio_eseguito()
            Testo_Mail = Nothing
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()


    End Sub

    Sub aggiungi_log_prezzi_0_OA(par_nome As String)

        Dim Cnn2 As New SqlConnection
        Cnn2.ConnectionString = Homepage.sap_tirelli
        Cnn2.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn2

        CMD_SAP.CommandText = "INSERT INTO [Tirelli_40].[dbo].[Log_Invio_mail] (nome,data,ora) values ('" & par_nome & "',getdate(),convert(varchar, getdate(), 108))"
        CMD_SAP.ExecuteNonQuery()
        Cnn2.Close()


    End Sub

    Sub invio_eseguito()
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = " update ""@PRG_TIR_INVSAP_LOG"" set name ='fatto'
                                FROM ""@PRG_TIR_INVSAP_LOG""
                            WHERE code ='" & code & "'  "
        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()


    End Sub

    Sub invio_mail_prezzi_0(par_tabella_intestazione As String, par_tabella_riga As String)

        Testo_Mail = Testo_Mail & "<table border=5>"
        Testo_Mail = Testo_Mail & "<tr>"

        Testo_Mail = Testo_Mail & "<th>Docnum</th>"
        Testo_Mail = Testo_Mail & "<th>Data reg</th>"
        Testo_Mail = Testo_Mail & "<th>Data consegna</th>"
        Testo_Mail = Testo_Mail & "<th>Owner</th>"
        Testo_Mail = Testo_Mail & "<th>fornitore</th>"
        Testo_Mail = Testo_Mail & "<th>Codice</th>"
        Testo_Mail = Testo_Mail & "<th>Descrizione</th>"
        Testo_Mail = Testo_Mail & "<th>quantità</th>"
        Testo_Mail = Testo_Mail & "<th>prezzo</th>"



        Testo_Mail = Testo_Mail & "</tr>"
        Testo_Mail = Testo_Mail & "<tr>"

        destinatario1 = "acquisti@tirelli.net"


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[DocNum]  ,T0.[DocStatus]  ,T0.[DocDate],T1.[ShipDate], t2.lastname   , T0.[CardName],  T1.[ItemCode]  ,T1.[Dscription]     , T1.[OpenQty], T1.PRICE

 FROM " & par_tabella_intestazione & " T0  INNER JOIN " & par_tabella_riga & " T1 ON T0.[DocEntry] = T1.[DocEntry] 
LEFT JOIN [TIRELLI_40].[dbo].OHEM T2 ON T0.[OwnerCode] = T2.[empID]
 inner JOIN 
 
 (
 select *
 from
 (SELECT t0.docnum, sum(case when t2.treetype='T' then 1 else 0 end) as 'Modello'
 FROM OPOR T0 INNER JOIN POR1 T1 ON T0.DOCENTRY=t1.DOCENTRY
 inner join oitt t2 on t2.code=t1.itemcode

 WHERE T1.[OpenQty] >0 
 group by t0.docnum
 )
 as t10 where t10.modello=0
 ) A on a.docnum =t0.docnum
 
WHERE T1.[OpenQty] >0 AND T1.[Price] =0
order by t2.lastname, t0.docnum
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader



        Do While cmd_SAP_reader.Read()


            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Docnum") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("DocDate") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ShipDate") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("lastname") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("CardName") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemCode") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Dscription") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & String.Format("{0:0.00}", cmd_SAP_reader("OpenQty")) & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & String.Format("{0:0.00}", cmd_SAP_reader("PRICE")) & "</td>"

            Testo_Mail = Testo_Mail & "</tr>"


        Loop

        Testo_Mail = Testo_Mail & "</table>"
        Testo_Mail = Testo_Mail & "</tr>"
        'Testo_Mail = Testo_Mail & "<td>" & FormatNumber(Totale, 2, , , TriState.True) & "</td>"






        Testo_Mail = Testo_Mail & "</BODY>"
        ' Invio E-Mail

        Dim Indice_Mail As Integer = 0

        Dim mySmtp As New SmtpClient
        Dim myMail As New MailMessage()



        mySmtp.UseDefaultCredentials = False
        mySmtp.Credentials = New Net.NetworkCredential(Homepage.Mittente_Mail, Homepage.password_mail)
        mySmtp.Host = "tirelli-net.mail.protection.outlook.com"
        mySmtp.Port = 25
        mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network
        mySmtp.EnableSsl = True


        myMail = New MailMessage()

        myMail.From = New MailAddress(Homepage.Mittente_Mail)

        myMail.To.Add(destinatario1)
        Try


            '   myMail.To.Add(destinatario2)

        Catch ex As Exception

        End Try
        If par_tabella_intestazione = "OPOR" Then
            myMail.Subject = "Ordini di acquisto a prezzo 0"
        ElseIf par_tabella_intestazione = "OPDN" Then
            myMail.Subject = "Entrate merce a prezzo 0"
        End If


        myMail.IsBodyHtml = True
        myMail.Body = Testo_Mail

        Try
            mySmtp.Send(myMail)
        Catch ex As Exception

        End Try

        Indice_Mail = Indice_Mail + 1
        invio_acquisti_eseguito(cmd_SAP_reader("ItemCode"))

        Testo_Mail = Nothing
        cmd_SAP_reader.Close()
        Cnn.Close()


    End Sub

    Sub invio_mail_pezzo_revisionato_agli_acquisti()
        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Cnn.Open()

            Dim query As String = "SELECT t10.itemcode, t10.itemname, t10.Disegno_new, t10.disegno_acquistato, t10.docnum, t10.cardname, t10.docdate, t10.ShipDate, " &
                            "CAST(t10.openqty AS DECIMAL) AS openqty, CAST(t10.LineTotal AS DECIMAL) AS linetotal, t10.email, t10.u_prg_tir_revprod " &
                            "FROM ( " &
                            "SELECT t0.itemcode, t3.itemname, t3.u_disegno AS Disegno_new, t1.u_disegno AS Disegno_acquistato, t2.docnum, t2.cardname, t2.docdate, " &
                            "t1.ShipDate, t1.openqty, t1.LineTotal, t4.email, t3.u_prg_tir_revprod " &
                            "FROM [TIRELLISRLDB].[DBO].codici_revisionati t0 " &
                            "LEFT JOIN por1 t1 ON t0.Itemcode = t1.ItemCode AND t1.OpenQty > 0 " &
                            "LEFT JOIN opor t2 ON t1.docentry = t2.docentry " &
                            "LEFT JOIN oitm t3 ON t3.itemcode = t0.itemcode " &
                            "LEFT JOIN [TIRELLI_40].[dbo].ohem t4 ON t4.empID = t2.ownercode " &
                            "WHERE t0.mail IS NULL AND t3.u_disegno <> t1.u_disegno) AS t10"

            Using CMD_SAP As New SqlCommand(query, Cnn)
                Using cmd_SAP_reader As SqlDataReader = CMD_SAP.ExecuteReader()

                    While cmd_SAP_reader.Read()
                        Dim Testo_Mail As New StringBuilder()

                        ' Creazione tabella HTML
                        Testo_Mail.AppendLine("<table border=5>")
                        Testo_Mail.AppendLine("<tr><th>Codice</th><th>Descrizione</th><th>OA</th><th>Fornitore</th><th>Data ordine</th><th>Consegna</th><th>Q</th><th>Totale</th><th>Disegno acquistato</th><th>Disegno nuovo</th><th>Tipo revisione</th></tr>")
                        Testo_Mail.AppendFormat("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td></tr>",
                                         cmd_SAP_reader("itemcode"), cmd_SAP_reader("itemname"), cmd_SAP_reader("docnum"), cmd_SAP_reader("cardname"),
                                         cmd_SAP_reader("docdate"), cmd_SAP_reader("shipdate"), cmd_SAP_reader("openqty"), cmd_SAP_reader("linetotal"),
                                         cmd_SAP_reader("disegno_acquistato"), cmd_SAP_reader("disegno_new"), cmd_SAP_reader("u_prg_tir_revprod"))
                        Testo_Mail.AppendLine("</table>")

                        ' Aggiunta link ai disegni


                        Testo_Mail.AppendFormat("<br><a href='" & Homepage.percorso_disegni_generico & "PDF\{0}.PDF'>Disegno acquistato</a>", cmd_SAP_reader("disegno_acquistato"))
                        Testo_Mail.AppendFormat("<br><a href='" & Homepage.percorso_disegni_generico & "PDF\{0}.PDF'>Disegno nuovo</a>", cmd_SAP_reader("disegno_new"))

                        ' Definizione destinatari
                        Dim destinatario1 As String = cmd_SAP_reader("email").ToString()
                        Dim destinatario2 As String = "giovanni.tirelli@tirelli.net"

                        ' Invio Email
                        Try
                            Using myMail As New MailMessage()
                                myMail.From = New MailAddress(Homepage.Mittente_Mail)
                                myMail.To.Add(destinatario1)
                                myMail.To.Add(destinatario2)
                                myMail.Subject = $"Codice acquistato revisionato {cmd_SAP_reader("itemcode")} {cmd_SAP_reader("itemname")} ordine {cmd_SAP_reader("docnum")} A {cmd_SAP_reader("cardname")} in consegna il {cmd_SAP_reader("shipdate")}"
                                myMail.IsBodyHtml = True
                                myMail.Body = Testo_Mail.ToString()

                                Using mySmtp As New SmtpClient("tirelli-net.mail.protection.outlook.com", 25)
                                    mySmtp.UseDefaultCredentials = False
                                    mySmtp.Credentials = New Net.NetworkCredential(Homepage.Mittente_Mail, Homepage.password_mail)
                                    mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network
                                    mySmtp.EnableSsl = True
                                    mySmtp.Send(myMail)
                                End Using
                            End Using

                            ' Aggiorna stato invio
                            invio_acquisti_eseguito(cmd_SAP_reader("ItemCode"))

                        Catch ex As Exception
                            ' Log dell'errore (da implementare con un sistema di logging adeguato)
                        End Try
                    End While
                End Using
            End Using
        End Using
    End Sub

    Sub invio_mail_revisione()

    End Sub

    Sub invio_acquisti_eseguito(par_itemcode As String)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = " update t11 set t11.mail='Y', t11.destinatario=t10.email, t11.data_invio =getdate()
from
(
select t0.itemcode, t3.itemname, t3.u_disegno as'Disegno_new', t1.u_disegno as 'Disegno_acquistato', t2.docnum, t2.cardname,t2.docdate, t1.ShipDate, t1.openqty, t1.LineTotal, t4.email
from [TIRELLISRLDB].[DBO].codici_revisionati t0 left join por1 t1 on t0.Itemcode=t1.ItemCode and t1.OpenQty>0
left join opor t2 on t1.docentry=t2.docentry
left join oitm t3 on t3.itemcode=t0.itemcode
left join [TIRELLI_40].[dbo].ohem t4 on t4.empID=t2.ownercode

where t0.mail is null and t0.itemcode='" & par_itemcode & "'
)
as t10 inner join [TIRELLISRLDB].[DBO].codici_revisionati t11 on t10.itemcode=t11.itemcode
where t10.email is not null
"
        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()


    End Sub

    Sub invio_mail_pezzo_revisionato_int()
        Try
            Using Cnn As New SqlConnection(Homepage.sap_tirelli)
                Cnn.Open()

                Dim query As String = "SELECT t0.itemcode, t3.itemname, t3.u_disegno AS Disegno_new, " &
                                  "t1.u_disegno AS Disegno_acquistato, t1.docnum, " &
                                  "ISNULL(t1.u_prg_azs_commessa, '') AS u_prg_azs_commessa, " &
                                  "ISNULL(t1.U_UTILIZZ, '') AS u_utilizz, t1.postdate, t1.DueDate, t1.PlannedQty, " &
                                  "coalesce(t5.email,'vanni.ponti@tirelli.net') AS Email, t3.u_prg_tir_revprod " &
                                  "FROM [TIRELLISRLDB].[DBO].codici_revisionati t0 " &
                                  "INNER JOIN owor t1 ON t0.itemcode = t1.itemcode " &
                                  "AND (t1.status = 'P' OR t1.status = 'R') " &
                                  "AND LEFT(t1.U_PRODUZIONE, 3) = 'INT' " &
                                  "LEFT JOIN oitm t3 ON t3.itemcode = t0.itemcode " &
                                  "left join  [TIRELLI_40].[dbo].ohem t5 on t5.userid=t1.usersign " &
                                  "WHERE t0.mail IS NULL AND t3.u_disegno <> t1.u_disegno"

                Using CMD_SAP As New SqlCommand(query, Cnn)
                    Using cmd_SAP_reader As SqlDataReader = CMD_SAP.ExecuteReader()

                        While cmd_SAP_reader.Read()
                            Dim destinatario1 As String = cmd_SAP_reader("Email").ToString()
                            Dim destinatario2 As String = "stefano.bruno@tirelli.net"

                            Dim Testo_Mail As New StringBuilder()
                            Testo_Mail.AppendLine("<table border='1' style='border-collapse: collapse;'>")
                            Testo_Mail.AppendLine("<tr><th>Codice</th><th>Descrizione</th><th>ODP</th><th>Commessa</th>" &
                                              "<th>Cliente</th><th>Data ordine</th><th>Consegna</th><th>Q</th>" &
                                              "<th>Disegno lanciato</th><th>Disegno nuovo</th><th>Tipo revisione</th></tr>")

                            Testo_Mail.AppendLine("<tr>")
                            Testo_Mail.AppendLine($"<td>{cmd_SAP_reader("itemcode")}</td>")
                            Testo_Mail.AppendLine($"<td>{cmd_SAP_reader("itemname")}</td>")
                            Testo_Mail.AppendLine($"<td>{cmd_SAP_reader("docnum")}</td>")
                            Testo_Mail.AppendLine($"<td>{cmd_SAP_reader("u_prg_azs_commessa")}</td>")
                            Testo_Mail.AppendLine($"<td>{cmd_SAP_reader("u_utilizz")}</td>")
                            Testo_Mail.AppendLine($"<td>{cmd_SAP_reader("postdate")}</td>")
                            Testo_Mail.AppendLine($"<td>{cmd_SAP_reader("duedate")}</td>")
                            Testo_Mail.AppendLine($"<td>{cmd_SAP_reader("plannedqty")}</td>")
                            Testo_Mail.AppendLine($"<td>{cmd_SAP_reader("disegno_acquistato")}</td>")
                            Testo_Mail.AppendLine($"<td>{cmd_SAP_reader("disegno_new")}</td>")
                            Testo_Mail.AppendLine($"<td>{cmd_SAP_reader("u_prg_tir_revprod")}</td>")
                            Testo_Mail.AppendLine("</tr>")
                            Testo_Mail.AppendLine("</table>")

                            ' Link ai disegni

                            Testo_Mail.AppendLine("<br><a href='" & Homepage.percorso_disegni_generico & "PDF/" & cmd_SAP_reader("disegno_acquistato") & ".PDF'>Disegno lanciato</a>")
                            Testo_Mail.AppendLine("<br><a href='" & Homepage.percorso_disegni_generico & "PDF/" & cmd_SAP_reader("disegno_new") & ".PDF'>Disegno nuovo</a>")
                            ' Invio email
                            Using mySmtp As New SmtpClient("tirelli-net.mail.protection.outlook.com")
                                mySmtp.UseDefaultCredentials = False
                                mySmtp.Credentials = New Net.NetworkCredential(Homepage.Mittente_Mail, Homepage.password_mail)
                                mySmtp.Port = 25
                                mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network
                                mySmtp.EnableSsl = True

                                Using myMail As New MailMessage()
                                    myMail.From = New MailAddress(Homepage.Mittente_Mail)
                                    myMail.To.Add(destinatario1)
                                    myMail.To.Add(destinatario2)
                                    myMail.Subject = $"Codice revisionato {cmd_SAP_reader("itemcode")} {cmd_SAP_reader("itemname")} " &
                                                $"ordine {cmd_SAP_reader("docnum")} commessa {cmd_SAP_reader("u_prg_azs_commessa")} " &
                                                $"{cmd_SAP_reader("u_utilizz")} lanciato il {cmd_SAP_reader("postdate")}"
                                    myMail.IsBodyHtml = True
                                    myMail.Body = Testo_Mail.ToString()

                                    Try
                                        mySmtp.Send(myMail)
                                    Catch ex As Exception
                                        Debug.WriteLine("Errore nell'invio email: " & ex.Message)
                                    End Try
                                End Using
                            End Using

                            ' Aggiorna lo stato dell'invio
                            invio_produzione_int_eseguito(cmd_SAP_reader("itemcode"))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine("Errore nella procedura di invio mail: " & ex.Message)
        End Try
    End Sub


    Sub invio_produzione_int_eseguito(par_itemcode As String)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = " update t11 set t11.mail='Y', t11.destinatario=t10.email, t11.data_invio=getdate()
from
(
select t0.itemcode, t3.itemname, t3.u_disegno as'Disegno_new', t1.u_disegno as 'Disegno_acquistato', t1.docnum, case when t1.u_prg_azs_commessa is null then '' else t1.u_prg_azs_commessa end as 'u_prg_azs_commessa', case when t1.U_UTILIZZ is null then '' else t1.U_UTILIZZ end as 'u_utilizz', t1.postdate,t1.DueDate, t1.PlannedQty,'vanniponti@tirelli.net' as 'Email', t3.u_prg_tir_revprod
from [TIRELLISRLDB].[DBO].codici_revisionati t0 inner join owor t1 on t0.itemcode=t1.itemcode and (t1.status='P' or t1.status='R') and substring(t1.U_PRODUZIONE,1,3)='INT'
	left join oitm t3 on t3.itemcode=t0.itemcode
	where t0.mail is null and t3.u_disegno<>t1.u_disegno and t0.itemcode='" & par_itemcode & "'
)
as t10 inner join [TIRELLISRLDB].[DBO].codici_revisionati t11 on t10.itemcode=t11.itemcode
where t10.email is not null
"
        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()


    End Sub


    Sub invio_mail_pezzo_revisionato_assemblaggio()
        Dim Cnn As New SqlConnection(Homepage.sap_tirelli)

        Try
            Cnn.Open()

            Dim CMD_SAP As New SqlCommand("
            SELECT t0.itemcode, t3.itemname, t3.u_disegno AS Disegno_new, t1.u_disegno AS Disegno_acquistato, 
       t1.docnum, ISNULL(t1.u_prg_azs_commessa, '') AS u_prg_azs_commessa, 
       ISNULL(t1.U_UTILIZZ, '') AS u_utilizz, t1.postdate, t1.DueDate, 
       t1.PlannedQty, coalesce(t5.email,'marco.morandini@tirelli.net') AS Email, t3.u_prg_tir_revprod, t1.usersign
FROM [TIRELLISRLDB].[DBO].codici_revisionati t0
left JOIN owor t1 ON t0.itemcode = t1.itemcode 
     AND (t1.status = 'P' OR t1.status = 'R') 
     AND t1.U_PRODUZIONE = 'assembl'
LEFT JOIN oitm t3 ON t3.itemcode = t0.itemcode
left join ousr t4 on t4.userid=t1.usersign
left join  [TIRELLI_40].[dbo].ohem t5 on t5.userid=t1.usersign

WHERE t0.mail IS NULL 
  AND t3.u_disegno <> t1.u_disegno
 GROUP BY t0.itemcode, t3.itemname, t3.u_disegno , t1.u_disegno , 
       t1.docnum, t1.u_prg_azs_commessa,
       t1.U_UTILIZZ, t1.postdate, t1.DueDate, 
       t1.PlannedQty, t5.email, t3.u_prg_tir_revprod, t1.usersign", Cnn)

            Dim cmd_SAP_reader As SqlDataReader = CMD_SAP.ExecuteReader()

            Dim mySmtp As New SmtpClient With {
            .UseDefaultCredentials = False,
            .Credentials = New Net.NetworkCredential(Homepage.Mittente_Mail, Homepage.password_mail),
            .Host = "tirelli-net.mail.protection.outlook.com",
            .Port = 25,
            .DeliveryMethod = SmtpDeliveryMethod.Network,
            .EnableSsl = True
        }

            While cmd_SAP_reader.Read()
                Dim Testo_Mail As New System.Text.StringBuilder()
                Testo_Mail.AppendLine("<table border=5>")
                Testo_Mail.AppendLine("<tr>")
                Testo_Mail.AppendLine("<th>Codice</th><th>Descrizione</th><th>ODP</th><th>Commessa</th><th>Cliente</th>")
                Testo_Mail.AppendLine("<th>Data ordine</th><th>Consegna</th><th>Q</th><th>Disegno lanciato</th><th>Disegno nuovo</th><th>Tipo revisione</th></tr>")
                Testo_Mail.AppendLine("<tr>")
                Testo_Mail.AppendFormat("<td>{0}</td>", cmd_SAP_reader("itemcode"))
                Testo_Mail.AppendFormat("<td>{0}</td>", cmd_SAP_reader("itemname"))
                Testo_Mail.AppendFormat("<td>{0}</td>", cmd_SAP_reader("docnum"))
                Testo_Mail.AppendFormat("<td>{0}</td>", cmd_SAP_reader("u_prg_azs_commessa"))
                Testo_Mail.AppendFormat("<td>{0}</td>", cmd_SAP_reader("u_utilizz"))
                Testo_Mail.AppendFormat("<td>{0}</td>", cmd_SAP_reader("postdate"))
                Testo_Mail.AppendFormat("<td>{0}</td>", cmd_SAP_reader("duedate"))
                Testo_Mail.AppendFormat("<td>{0}</td>", cmd_SAP_reader("plannedqty"))
                Testo_Mail.AppendFormat("<td>{0}</td>", cmd_SAP_reader("disegno_acquistato"))
                Testo_Mail.AppendFormat("<td>{0}</td>", cmd_SAP_reader("disegno_new"))
                Testo_Mail.AppendFormat("<td>{0}</td>", cmd_SAP_reader("u_prg_tir_revprod"))
                Testo_Mail.AppendLine("</tr></table>")

                ' Link ai disegni
                Testo_Mail.AppendFormat("<br><a href='" & Homepage.percorso_disegni_generico & "PDF\{0}.PDF'>Disegno lanciato</a>", cmd_SAP_reader("disegno_acquistato"))
                Testo_Mail.AppendFormat("<br><a href='" & Homepage.percorso_disegni_generico & "PDF\{0}.PDF'>Disegno nuovo</a>", cmd_SAP_reader("disegno_new"))
                Testo_Mail.AppendLine("</BODY>")

                ' Invio email
                Dim myMail As New MailMessage()
                myMail.From = New MailAddress(Homepage.Mittente_Mail)
                myMail.To.Add(cmd_SAP_reader("email"))
                Try
                    myMail.To.Add(destinatario2)
                Catch ex As Exception
                    Debug.Print("Errore aggiunta destinatario2: " & ex.Message)
                End Try

                myMail.Subject = String.Format("Codice revisionato {0} {1} ordine {2} commessa {3} {4} lanciato il {5}",
                                           cmd_SAP_reader("itemcode"),
                                           cmd_SAP_reader("itemname"),
                                           cmd_SAP_reader("docnum"),
                                           cmd_SAP_reader("u_prg_azs_commessa"),
                                           cmd_SAP_reader("u_utilizz"),
                                           cmd_SAP_reader("postdate"))

                myMail.IsBodyHtml = True
                myMail.Body = Testo_Mail.ToString()

                Try
                    mySmtp.Send(myMail)
                    invio_produzione_assemblaggio_eseguito()
                Catch ex As Exception
                    Debug.Print("Errore invio email: " & ex.Message)
                End Try
            End While

        Catch ex As Exception
            Debug.Print("Errore generale: " & ex.Message)
        Finally
            If Cnn IsNot Nothing AndAlso Cnn.State = ConnectionState.Open Then
                Cnn.Close()
            End If
        End Try
    End Sub

    Sub invio_produzione_assemblaggio_eseguito()
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = " update t11 set t11.mail='Y', t11.destinatario=t10.email
from
(
select t0.itemcode, t3.itemname, t3.u_disegno as'Disegno_new', t1.u_disegno as 'Disegno_acquistato', t1.docnum, case when t1.u_prg_azs_commessa is null then '' else t1.u_prg_azs_commessa end as 'u_prg_azs_commessa', case when t1.U_UTILIZZ is null then '' else t1.U_UTILIZZ end as 'u_utilizz', t1.postdate,t1.DueDate, t1.PlannedQty,'vincenzoscaltrito@tirelli.net' as 'Email', t3.u_prg_tir_revprod
from [TIRELLISRLDB].[DBO].codici_revisionati t0 inner join owor t1 on t0.itemcode=t1.itemcode and (t1.status='P' or t1.status='R') and t1.U_PRODUZIONE='assembl'
	left join oitm t3 on t3.itemcode=t0.itemcode
	where t0.mail is null and t3.u_disegno<>t1.u_disegno
)
as t10 inner join [TIRELLISRLDB].[DBO].codici_revisionati t11 on t10.itemcode=t11.itemcode
where t10.email is not null
"
        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()


    End Sub

    Sub invio_per_revisione_eseguito()
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = " update t11 set t11.mail='Y'
from
(
select t0.itemcode, t3.itemname, t3.u_disegno as'Disegno_new', t1.u_disegno as 'Disegno_acquistato', t2.docnum, t2.cardname,t2.docdate, t1.ShipDate, t1.openqty, t1.LineTotal, t4.email
from [TIRELLISRLDB].[DBO].codici_revisionati t0 left join por1 t1 on t0.Itemcode=t1.ItemCode and t1.OpenQty>0
left join opor t2 on t1.docentry=t2.docentry
left join oitm t3 on t3.itemcode=t0.itemcode
left join [TIRELLI_40].[dbo].ohem t4 on t4.empID=t2.ownercode

where t0.mail is null
)
as t10 inner join [TIRELLISRLDB].[DBO].codici_revisionati t11 on t10.itemcode=t11.itemcode "
        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()


    End Sub

    Sub check_mail_già_inviata(par_nome As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT *
FROM [Tirelli_40].[dbo].[Log_Invio_mail]
WHERE nome = '" & par_nome & "' AND CAST(data AS DATE) = CAST(GETDATE() AS DATE)
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader



        If cmd_SAP_reader.Read() Then


            check_inviato = "SI"
        Else


            check_inviato = "NO"

        End If

        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub check_presenza_disegno()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "select *
from
(
select t1.itemcode, coalesce(t2.U_Disegno,'') as 'U_disegno'
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
left join oitm t2 on t1.itemcode=t2.itemcode

where (t0.status='R' or t0.status='P') and (substring(t1.itemcode,1,1)='0' or substring(t1.itemcode,1,1)='D')
group by t1.itemcode, t2.U_Disegno

union all 

select t1.itemcode, coalesce(t2.U_Disegno,'')
from rdr1 t1
left join oitm t2 on t1.itemcode=t2.itemcode

where (t1.linestatus='O') and (substring(t1.itemcode,1,1)='0' or substring(t1.itemcode,1,1)='D')
group by t1.itemcode, t2.U_Disegno
)
as t10
where t10.u_disegno<>''
group by t10.itemcode,t10.u_disegno
order by t10.itemcode
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader



        Do While cmd_SAP_reader.Read()


            If File.Exists(Homepage.percorso_disegni_generico & "PDF\" & cmd_SAP_reader("U_Disegno") & ".PDF") Then

            Else
                nuovo_record_disegni_mancanti(cmd_SAP_reader("ITEMCODE"), cmd_SAP_reader("U_Disegno"), "Manca PDF")
            End If

            If File.Exists(Homepage.percorso_disegni_generico & "DXF\" & cmd_SAP_reader("U_Disegno") & ".DXF") Then

            Else
                nuovo_record_disegni_mancanti(cmd_SAP_reader("ITEMCODE"), cmd_SAP_reader("U_Disegno"), "Manca DXF")
            End If

        Loop

        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub nuovo_record_disegni_mancanti(PAR_CODICE_SAP As String, PAR_CODICE_DISEGNO As String, PAR_STATO As String)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = " INSERT INTO [TIRELLI_40].[dbo].[disegni_mancanti]
           ([codice_sap]
           ,[disegno]
           ,[data]
           ,[Stato])
     VALUES
           ('" & PAR_CODICE_SAP & "'
           ,'" & PAR_CODICE_DISEGNO & "'
           ,GETDATE()
           ,'" & PAR_STATO & "')
 "
        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()


    End Sub

    Sub azzera_disegni_mancanti()
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand
        CMD_SAP_3.Connection = Cnn3

        ' Cancella tutti i record
        CMD_SAP_3.CommandText = "DELETE FROM [TIRELLI_40].[dbo].[disegni_mancanti]"
        CMD_SAP_3.ExecuteNonQuery()

        ' Resetta l'IDENTITY
        CMD_SAP_3.CommandText = "DBCC CHECKIDENT ('[TIRELLI_40].[dbo].[disegni_mancanti]', RESEED, 0)"
        CMD_SAP_3.ExecuteNonQuery()

        Cnn3.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Homepage.ID_SALVATO = 2

        Homepage.Aggiorna_INI_COMPUTER()
        End
    End Sub

    Private Sub Timer_stampante_Tick(sender As Object, e As EventArgs) Handles Timer_stampante.Tick
        'Stampante_3D.risposta_stampante()
    End Sub

    Sub invia_ordini_inutili(par_percorso_file As String)
        ' Percorso del file Excel
        Dim filePath As String = par_percorso_file


        ' SendEmail(filePath, "MAIL VERIFICA ORDINI A COMMESSA GIA' DISPONIBILI", "Controllare ordini che risultato matematicamente già disponibili.", "giovanni.tirelli@tirelli.net", "marco.morandini@tirelli.net", "acquisti@tirelli.net")
    End Sub
    Sub invia_report_ordini_in_garanzia(par_percorso_file As String)


        ' SendEmail(par_percorso_file, "REPORT ORDINI-INTERVENTI IN GARANZIA", "Report degli ultimi ordini generati in garanzia. Primo foglio gli ordine cliente, secondo gli interventi, terzo tutti gli ordine clienti aperti con causale diversa dalla vendita", "giovanni.tirelli@tirelli.net", "franco.porracin@tirelli.net", "carlo.tonini@tirelli.net")
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub SendEmail(ByVal attachmentPath As String, par_oggetto_mail As String, par_corpo_mail As String, par_destinatario_1 As String, par_destinatario_2 As String, par_destinatario_3 As String)
        Try
            Dim SmtpServer As New SmtpClient("tirelli-net.mail.protection.outlook.com") ' Cambia con il server SMTP corretto
            Dim mail As New MailMessage()
            Dim attachment As System.Net.Mail.Attachment

            attachment = New System.Net.Mail.Attachment(attachmentPath)
            mail.From = New MailAddress("report@tirelli.net") ' Cambia con il tuo indirizzo email
            If par_destinatario_1 = "" Then

            Else
                mail.To.Add(par_destinatario_1)
            End If
            If par_destinatario_2 = "" Then

            Else
                mail.To.Add(par_destinatario_2)
            End If
            If par_destinatario_3 = "" Then

            Else
                mail.To.Add(par_destinatario_3)
            End If



            mail.Subject = par_oggetto_mail
            mail.Body = par_corpo_mail
            mail.Attachments.Add(attachment)

            SmtpServer.Port = 25 ' Cambia con la porta corretta
            SmtpServer.Credentials = New Net.NetworkCredential(Homepage.Mittente_Mail, Homepage.password_mail) ' Cambia con le tue credenziali
            SmtpServer.EnableSsl = True

            SmtpServer.Send(mail)
            ' MessageBox.Show("Mail inviata con successo!")
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        invia_ordini_inutili(Homepage.percorso_acquisti & "\Ordini Inutili.xlsx")

    End Sub

    Private Sub Timer_giornaliero_Tick(sender As Object, e As EventArgs) Handles Timer_giornaliero.Tick
        ' Ottieni la data e l'ora corrente
        'Dim now As DateTime = DateTime.Now

        '' Verifica se è venerdì tra le 08:59:00 e le 09:00:10

        'Dim startTime As DateTime = New DateTime(now.Year, now.Month, now.Day, 8, 59, 0)
        'Dim endTime As DateTime = New DateTime(now.Year, now.Month, now.Day, 9, 0, 10)
        'If now >= startTime AndAlso now <= endTime Then
        '    ' Esegui la funzione invia_ordini_inutili

        '    invia_ordini_inutili(Homepage.percorso_acquisti & "\Ordini Inutili.xlsx")
        '    invia_report_ordini_in_garanzia(Homepage.percorso_server & "00-Tirelli 4.0\Report\Ordini in garanzia-recall.xlsx")
        'End If

    End Sub

    Private Sub Timer_notturno_Tick(sender As Object, e As EventArgs) Handles Timer_notturno.Tick

    End Sub
End Class