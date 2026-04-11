Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Public Class Form_mail_ricambi



    'Dati fornitore
    Public Testo_Mail As String
        Public sender_mail As String = "report@tirelli.net"
    'Public Password_Mail As String = "Ras70773"
    'Public Destinatario_Mail As String = "giovannitirelli@tirelli.net"
    'Public Testo_mail_fornitura As String
    'Public Testo_intestazione_fornitura As String = "Data" & Space(5) & "Ordine" & Space(5) & "Codice" & Space(5) & "Descrizione" & Space(5) & "Disegno" & Space(5) & "Codice_fornitore" & Space(5) & "Consegna_old" & Space(5) & "Consegna_new" & Space(5) & "RIT/ANT"
    Public destinatario1 As String
        Public destinatario2 As String
        Public distinta As String


    Public code As String
        Public ID_record As Integer
        Public ID_record_cancellazione As Integer
        'Public Cnn4 As New SqlConnection
        'Public cnn5 As New SqlConnection

        Public mail_destinatario As String
        Public codice_bp_new As String
        Public nome_bp_destinatario As String


    Public variabile_invio_giorno As String = "N"
        Public variabile_invio_ora As String = "N"
        Public mail_inviate As Integer = 0
        Public riga As Integer

        Public mostratoMessaggio As Boolean = False
        Private contatore As Integer = 0

        Private Sub Button_invio_mail_Click(sender As Object, e As EventArgs)
            'pagamenti()
            invio_mail_errore_ivsaplink()


        End Sub

        Sub invio_mail_errore_ivsaplink()



        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()


        Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
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
                    destinatario1 = "deniscattabriga@tirelli.net"
                End If
                distinta = cmd_SAP_reader("u_art_padre_db")
                destinatario2 = cmd_SAP_reader("destinatario2")
                code = cmd_SAP_reader("code")
                'MsgBox(destinatario1)
                ' MsgBox(destinatario2)


                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("u_data_ora") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("u_art_padre_db") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("u_errore") & "</td>"

                'Testo_Mail = Testo_Mail & "<td>" & FormatNumber(Totale, 2, , , TriState.True) & "</td>"

                Testo_Mail = Testo_Mail & "</tr>"
                Testo_Mail = Testo_Mail & "</table>"

                Testo_Mail = Testo_Mail & "</BODY>"
                ' Invio E-Mail

                Dim Indice_Mail As Integer = 0


                Dim mySmtp As New SmtpClient
                Dim myMail As New MailMessage()



                mySmtp.UseDefaultCredentials = False
            mySmtp.Credentials = New Net.NetworkCredential(sender_mail, Homepage.password_mail)
            mySmtp.Host = "smtp.office365.com"
                mySmtp.Port = 25
                mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network
                mySmtp.EnableSsl = True


                myMail = New MailMessage()

                myMail.From = New MailAddress(sender_mail)

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
        cnn.Close()


    End Sub

        Sub invio_mail_pezzo_revisionato_agli_acquisti()

        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()


        Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "select t10.itemcode, t10.itemname, t10.Disegno_new, t10.disegno_acquistato, t10.docnum, t10.cardname,t10.docdate, t10.ShipDate, cast(t10.openqty as decimal) as 'openqty', cast(t10.LineTotal as decimal) as 'linetotal', t10.email, t10.u_prg_tir_revprod
from
(
select t0.itemcode, t3.itemname, t3.u_disegno as'Disegno_new', t1.u_disegno as 'Disegno_acquistato', t2.docnum, t2.cardname,t2.docdate, t1.ShipDate, t1.openqty, t1.LineTotal, t4.email, t3.u_prg_tir_revprod
from [TIRELLISRLDB].[DBO].codici_revisionati t0 left join por1 t1 on t0.Itemcode=t1.ItemCode and t1.OpenQty>0
left join opor t2 on t1.docentry=t2.docentry
left join oitm t3 on t3.itemcode=t0.itemcode
left join [TIRELLI_40].[dbo].ohem t4 on t4.empID=t2.ownercode

where t0.mail is null
)
as t10
where t10.email is not null
"

            cmd_SAP_reader = CMD_SAP.ExecuteReader


            Do While cmd_SAP_reader.Read()
                Testo_Mail = Testo_Mail & "<table border=5>"
                Testo_Mail = Testo_Mail & "<tr>"

                Testo_Mail = Testo_Mail & "<th>Codice</th>"
                Testo_Mail = Testo_Mail & "<th>Descrizione</th>"
                Testo_Mail = Testo_Mail & "<th>OA</th>"
                Testo_Mail = Testo_Mail & "<th>Fornitore</th>"
                Testo_Mail = Testo_Mail & "<th>Data ordine</th>"
                Testo_Mail = Testo_Mail & "<th>Consegna</th>"
                Testo_Mail = Testo_Mail & "<th>Q</th>"
                Testo_Mail = Testo_Mail & "<th>Totale</th>"
                Testo_Mail = Testo_Mail & "<th>Disegno acquistato</th>"
                Testo_Mail = Testo_Mail & "<th>Disegno nuovo</th>"
                Testo_Mail = Testo_Mail & "<th>Tipo revisione</th>"


                Testo_Mail = Testo_Mail & "</tr>"
                Testo_Mail = Testo_Mail & "<tr>"

                destinatario1 = cmd_SAP_reader("email")

                'destinatario2 = "giovannitirelli@tirelli.net"



                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("itemcode") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("itemname") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("docnum") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("cardname") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Docdate") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("shipdate") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("openqty") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("linetotal") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("disegno_acquistato") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("disegno_new") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("u_prg_tir_revprod") & "</td>"

                'Testo_Mail = Testo_Mail & "<td>" & FormatNumber(Totale, 2, , , TriState.True) & "</td>"

                Testo_Mail = Testo_Mail & "</tr>"
                Testo_Mail = Testo_Mail & "</table>"
                Testo_Mail = Testo_Mail & "</tr>"



                Testo_Mail = Testo_Mail & "<br><a href='\\192.168.0.150\k\Tecnico\Disegni Meccanici\PDF-DXF\PDF\" & cmd_SAP_reader("disegno_acquistato") & ".PDF'>Disegno acquistato</a>"
                Testo_Mail = Testo_Mail & "</tr>"

                Testo_Mail = Testo_Mail & "<br><a href='\\192.168.0.150\k\Tecnico\Disegni Meccanici\PDF-DXF\PDF\" & cmd_SAP_reader("disegno_new") & ".PDF'>Disegno nuovo</a>"
                Testo_Mail = Testo_Mail & "</BODY>"
                ' Invio E-Mail

                Dim Indice_Mail As Integer = 0


                Dim mySmtp As New SmtpClient
                Dim myMail As New MailMessage()



                mySmtp.UseDefaultCredentials = False
            mySmtp.Credentials = New Net.NetworkCredential(sender_mail, Homepage.password_mail)
            mySmtp.Host = "tirelli-net.mail.protection.outlook.com"
                mySmtp.Port = 25
                mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network
                mySmtp.EnableSsl = True


                myMail = New MailMessage()

                myMail.From = New MailAddress(sender_mail)

                myMail.To.Add(destinatario1)
                Try


                    myMail.To.Add(destinatario2)

                Catch ex As Exception

                End Try
                myMail.Subject = "Codice revisionato " & cmd_SAP_reader("itemcode") & " " & cmd_SAP_reader("itemname") & " ordine " & cmd_SAP_reader("docnum") & " A " & cmd_SAP_reader("Cardname") & " in consegna il " & cmd_SAP_reader("shipdate") & ""
                myMail.IsBodyHtml = True
                myMail.Body = Testo_Mail

                Try
                    mySmtp.Send(myMail)
                Catch ex As Exception

                End Try

                Indice_Mail = Indice_Mail + 1
                invio_acquisti_eseguito()

                Testo_Mail = Nothing
            Loop
            cmd_SAP_reader.Close()
        cnn.Close()

        invio_per_revisione_eseguito()

        End Sub


    Sub invio_acquisti_eseguito()
        Dim CNN3 As New SqlConnection
        CNN3.ConnectionString = Homepage.sap_tirelli
        cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = cnn3


        CMD_SAP_3.CommandText = " update t11 set t11.mail='Y', t11.destinatario=t10.email
from
(
select t0.itemcode, t3.itemname, t3.u_disegno as'Disegno_new', t1.u_disegno as 'Disegno_acquistato', t2.docnum, t2.cardname,t2.docdate, t1.ShipDate, t1.openqty, t1.LineTotal, t4.email
from [TIRELLISRLDB].[DBO].codici_revisionati t0 left join por1 t1 on t0.Itemcode=t1.ItemCode and t1.OpenQty>0
left join opor t2 on t1.docentry=t2.docentry
left join oitm t3 on t3.itemcode=t0.itemcode
left join [TIRELLI_40].[dbo].ohem t4 on t4.empID=t2.ownercode

where t0.mail is null
)
as t10 inner join [TIRELLISRLDB].[DBO].codici_revisionati t11 on t10.itemcode=t11.itemcode
where t10.email is not null
"
        CMD_SAP_3.ExecuteNonQuery()
        cnn3.Close()


    End Sub

    Sub invio_per_revisione_eseguito()
        Dim CNN3 As New SqlConnection
        CNN3.ConnectionString = Homepage.sap_tirelli
        cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = cnn3


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
        cnn3.Close()


    End Sub


    Sub invio_eseguito()
        Dim CNN3 As New SqlConnection
        CNN3.ConnectionString = Homepage.sap_tirelli
        cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = cnn3


        CMD_SAP_3.CommandText = " update ""@PRG_TIR_INVSAP_LOG"" set name ='fatto'
                                FROM ""@PRG_TIR_INVSAP_LOG""
                            WHERE code ='" & code & "'  "
        CMD_SAP_3.ExecuteNonQuery()
        cnn3.Close()


    End Sub



    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
            Try

                invio_mail_errore_ivsaplink()
            Catch ex As Exception

            End Try



            invio_mail_pezzo_revisionato_agli_acquisti()

        End Sub

        Private Sub Button1_Click(sender As Object, e As EventArgs)
            Me.WindowState = FormWindowState.Minimized
        End Sub



        Private Sub Button2_Click(sender As Object, e As EventArgs)
            invio_mail_pezzo_revisionato_agli_acquisti()
        End Sub

        Public Sub Invia_Mail_interna(id As Integer)




            Dim Testo_Mail As String
            ' Testo_Mail = "<BODY><H3></h3><P>"
            '  Testo_Mail = Testo_Mail & ""



            ' Testo_Mail = Testo_Mail & "<BR><BR>"
            Testo_Mail = Testo_Mail & "This is an automatic message. Please contact spareparts@tirelli.net for additional details.<BR>"

            Testo_Mail = Testo_Mail & "<BR><H3>AROL NORTH AMERICA</h3>"

            Testo_Mail = Testo_Mail & "</P></BODY>"

            Testo_Mail = Testo_Mail & "<table border=5>"
            Testo_Mail = Testo_Mail & "<tr>"
            Testo_Mail = Testo_Mail & "<th>Order N°</th>"
            Testo_Mail = Testo_Mail & "<th>Your PO N°</th>"
            Testo_Mail = Testo_Mail & "<th>ItemCode</th>"
            Testo_Mail = Testo_Mail & "<th>ItemName</th>"

            Testo_Mail = Testo_Mail & "<th>Description</th>"

            Testo_Mail = Testo_Mail & "<th>Insert date</th>"
            Testo_Mail = Testo_Mail & "<th>Our Due date</th>"
            Testo_Mail = Testo_Mail & "<th>Causal</th>"
            'Testo_Mail = Testo_Mail & "<th>Branch</th>"
            Testo_Mail = Testo_Mail & "<th>Customer</th>"

            Testo_Mail = Testo_Mail & "<th>Open Qty</th>"
            Testo_Mail = Testo_Mail & "<th>Status</th>"
            Testo_Mail = Testo_Mail & "<th>Supplier Due date</th>"
            Testo_Mail = Testo_Mail & "<th>Toolshop due date</th>"

            Testo_Mail = Testo_Mail & "</tr>"
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select T40.DOCNUM as 'Order n°',t40.numatcard as 'Your PO N°',T40.[ItemCode], T40.[ItemName], T40.[Branch] ,t40.u_descing as 'English name',   t40.docdate as 'Insert date', t40.shipdate as 'Our Due date', t40.u_causcons as 'Causal', t40.cardname as 'Customer', t40.openqty as 'Open Qty', t40.status, min (t41.shipdate) as 'Supplier Due date', cast(min(t42.u_data_cons_mes) as date) as 'Toolshop due date'
from
(
Select T30.DOCNUM,t30.numatcard,T30.[ItemCode], T30.[ItemName],T30.[Branch], t30.u_descing, T30.[ItmsGrpNam], t30.docdate, t30.shipdate, t30.u_causcons, t30.cardname, t30.openqty, t30.u_trasferito, t30.u_datrasferire, t30.Mag, t30.Clavter, t30.Disp, case when t30.u_datrasferire = 0 then 'Picked'  when t30.mag >=t30.u_datrasferire then 'Warehouse' when t30.clavter>= t30.u_datrasferire then 'On treatment'  when t30.disp >=0 then 'On order' else 'Yet to be ordered' end as 'Status'
from
(
Select T20.DOCNUM,t20.numatcard,T20.[ItemCode], T20.[ItemName],T20.[Branch],t20.u_descing,  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter, sum(t21.onhand-t21.iscommited+t21.onorder) as 'Disp'

from
(
Select T10.DOCNUM,t10.numatcard,T10.[ItemCode], T10.[ItemName],T10.[Branch], t10.u_descing, T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, sum(t11.onhand) as 'Mag', case when t12.onhand is null then 0 else t12.onhand end as 'Clavter'


FROM
(
SELECT T0.DOCNUM,t0.numatcard,T2.[ItemCode], T2.[ItemName],t0.cardname as 'Branch', t2.u_descing,  T3.[ItmsGrpNam], t0.docdate, t1.shipdate, case when t0.u_causcons ='V' then 'Sold' when t0.u_causcons ='GAR' then 'Warranty' else t0.u_causcons end as 'u_causcons' , t4.cardname, t1.openqty, t1.u_trasferito, t1.u_datrasferire

FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] INNER JOIN OITB T3 ON T2.[ItmsGrpCod] = T3.[ItmsGrpCod] 
left JOIN OCRD T4 ON T0.[u_codicebp] = T4.[CardCode]

WHERE T1.[LineStatus] ='O' and (substring(t2.itemcode,1,1)='0' or substring(t2.itemcode,1,1)='C' or substring(t2.itemcode,1,1)='D') and (t0.cardcode='01872') and t0.u_causcons<>'COMP'
)
as t10 left join oitw t11 on t11.itemcode=t10.itemcode and t11.whscode<>'WIP' and t11.whscode<>'Clavter'
left join oitw t12 on t12.itemcode=t10.itemcode and t12.whscode='Clavter'

group by T10.DOCNUM,t10.numatcard,T10.[ItemCode], T10.[ItemName],T10.[Branch],  T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, t12.onhand,t10.u_descing
)
as t20 left join oitw t21 on t20.itemcode=t21.itemcode

group by T20.DOCNUM,t20.numatcard,T20.[ItemCode], T20.[ItemName],T20.[Branch],  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter,t20.u_descing
)
as t30
)
as t40 left join por1 t41 on t41.itemcode=t40.itemcode and t41.openqty>0 and t40.status ='On order'
left join owor t42 on t42.itemcode=t40.itemcode and (t42.status='P' or t42.status='R') and t42.u_produzione='Int' and t40.status ='On order'

group by T40.DOCNUM,t40.numatcard,T40.[ItemCode], T40.[ItemName],T40.[Branch],  T40.[ItmsGrpNam], t40.docdate, t40.shipdate, t40.u_causcons, t40.cardname, t40.openqty, t40.u_trasferito, t40.u_datrasferire, t40.Mag, t40.Clavter, t40.Disp, t40.status,t40.u_descing

order by T40.DOCNUM  "

            cmd_SAP_reader = CMD_SAP.ExecuteReader


            Do While cmd_SAP_reader.Read()

                Testo_Mail = Testo_Mail & "<tr>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Order n°") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Your PO N°") & "</td>"


                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemCode") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemName") & "</td>"



                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("English name") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Insert date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Our Due date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Causal") & "</td>"



                'Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Branch") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Customer") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & FormatNumber(cmd_SAP_reader("Open Qty"), 2, , , TriState.True) & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Status") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Supplier Due date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Toolshop due date") & "</td>"
                Testo_Mail = Testo_Mail & "</tr>"
            Loop
            cmd_SAP_reader.Close()
        cnn.Close()

        Testo_Mail = Testo_Mail & "</table>"




            Testo_Mail = Testo_Mail & "<BR><BR>"
            Testo_Mail = Testo_Mail & "<BR><h3>AROL LATIN AMERICA</h3>"
            Testo_Mail = Testo_Mail & "<BR><BR>"
            Testo_Mail = Testo_Mail & "</P></BODY>"

            Testo_Mail = Testo_Mail & "<table border=5>"
            Testo_Mail = Testo_Mail & "<tr>"
            Testo_Mail = Testo_Mail & "<th>Order N°</th>"
            Testo_Mail = Testo_Mail & "<th>Your PO N°</th>"
            Testo_Mail = Testo_Mail & "<th>ItemCode</th>"
            Testo_Mail = Testo_Mail & "<th>ItemName</th>"

            Testo_Mail = Testo_Mail & "<th>Description</th>"
            Testo_Mail = Testo_Mail & "<th>Insert date</th>"
            Testo_Mail = Testo_Mail & "<th>Our Due date</th>"
            Testo_Mail = Testo_Mail & "<th>Causal</th>"
            'Testo_Mail = Testo_Mail & "<th>Branch</th>"
            Testo_Mail = Testo_Mail & "<th>Customer</th>"

            Testo_Mail = Testo_Mail & "<th>Open Qty</th>"
            Testo_Mail = Testo_Mail & "<th>Status</th>"
            Testo_Mail = Testo_Mail & "<th>Supplier Due date</th>"
            Testo_Mail = Testo_Mail & "<th>Toolshop due date</th>"

            Testo_Mail = Testo_Mail & "</tr>"

        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select T40.DOCNUM as 'Order n°',t40.numatcard as 'Your PO N°',T40.[ItemCode], T40.[ItemName], T40.[Branch] ,t40.u_descing as 'English name',   t40.docdate as 'Insert date', t40.shipdate as 'Our Due date', t40.u_causcons as 'Causal', t40.cardname as 'Customer', t40.openqty as 'Open Qty', t40.status, min (t41.shipdate) as 'Supplier Due date', cast(min(t42.u_data_cons_mes) as date) as 'Toolshop due date'
from
(
Select T30.DOCNUM,t30.numatcard,T30.[ItemCode], T30.[ItemName],T30.[Branch], t30.u_descing, T30.[ItmsGrpNam], t30.docdate, t30.shipdate, t30.u_causcons, t30.cardname, t30.openqty, t30.u_trasferito, t30.u_datrasferire, t30.Mag, t30.Clavter, t30.Disp, case when t30.u_datrasferire = 0 then 'Picked'  when t30.mag >=t30.u_datrasferire then 'Warehouse' when t30.clavter>= t30.u_datrasferire then 'On treatment'  when t30.disp >=0 then 'On order' else 'Yet to be ordered' end as 'Status'
from
(
Select T20.DOCNUM,t20.numatcard,T20.[ItemCode], T20.[ItemName],T20.[Branch],t20.u_descing,  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter, sum(t21.onhand-t21.iscommited+t21.onorder) as 'Disp'

from
(
Select T10.DOCNUM,t10.numatcard,T10.[ItemCode], T10.[ItemName],T10.[Branch], t10.u_descing, T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, sum(t11.onhand) as 'Mag', case when t12.onhand is null then 0 else t12.onhand end as 'Clavter'


FROM
(
SELECT T0.DOCNUM,t0.numatcard,T2.[ItemCode], T2.[ItemName],t0.cardname as 'Branch', t2.u_descing,  T3.[ItmsGrpNam], t0.docdate, t1.shipdate, case when t0.u_causcons ='V' then 'Sold' when t0.u_causcons ='GAR' then 'Warranty' else t0.u_causcons end as 'u_causcons', t4.cardname, t1.openqty, t1.u_trasferito, t1.u_datrasferire

FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] INNER JOIN OITB T3 ON T2.[ItmsGrpCod] = T3.[ItmsGrpCod] 
left JOIN OCRD T4 ON T0.[u_codicebp] = T4.[CardCode]

WHERE T1.[LineStatus] ='O' and (substring(t2.itemcode,1,1)='0' or substring(t2.itemcode,1,1)='C' or substring(t2.itemcode,1,1)='D') and (t0.cardcode='01924') and t0.u_causcons<>'COMP'
)
as t10 left join oitw t11 on t11.itemcode=t10.itemcode and t11.whscode<>'WIP' and t11.whscode<>'Clavter'
left join oitw t12 on t12.itemcode=t10.itemcode and t12.whscode='Clavter'

group by T10.DOCNUM,t10.numatcard,T10.[ItemCode], T10.[ItemName],T10.[Branch],  T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, t12.onhand,t10.u_descing
)
as t20 left join oitw t21 on t20.itemcode=t21.itemcode

group by T20.DOCNUM,t20.numatcard,T20.[ItemCode], T20.[ItemName],T20.[Branch],  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter,t20.u_descing
)
as t30
)
as t40 left join por1 t41 on t41.itemcode=t40.itemcode and t41.openqty>0 and t40.status ='On order'
left join owor t42 on t42.itemcode=t40.itemcode and (t42.status='P' or t42.status='R') and t42.u_produzione='Int' and t40.status ='On order'

group by T40.DOCNUM,t40.numatcard,T40.[ItemCode], T40.[ItemName],T40.[Branch],  T40.[ItmsGrpNam], t40.docdate, t40.shipdate, t40.u_causcons, t40.cardname, t40.openqty, t40.u_trasferito, t40.u_datrasferire, t40.Mag, t40.Clavter, t40.Disp, t40.status,t40.u_descing

order by T40.DOCNUM  "

            cmd_SAP_reader = CMD_SAP.ExecuteReader


            Do While cmd_SAP_reader.Read()

                Testo_Mail = Testo_Mail & "<tr>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Order n°") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Your PO N°") & "</td>"


                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemCode") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemName") & "</td>"



                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("English name") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Insert date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Our Due date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Causal") & "</td>"



                'Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Branch") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Customer") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & FormatNumber(cmd_SAP_reader("Open Qty"), 2, , , TriState.True) & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Status") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Supplier Due date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Toolshop due date") & "</td>"
                Testo_Mail = Testo_Mail & "</tr>"
            Loop
            cmd_SAP_reader.Close()
        cnn.Close()

        Testo_Mail = Testo_Mail & "</table>"

            Testo_Mail = Testo_Mail & "<BR><H3>Status legend:</h3>"
            Testo_Mail = Testo_Mail & "<BR>Picked = Ready for shipment"
            Testo_Mail = Testo_Mail & "<BR>Warehouse = Available for picking"
            Testo_Mail = Testo_Mail & "<BR>On order = Parts not received or produced yet"
            Testo_Mail = Testo_Mail & "<BR>Yet to be ordered = Looking for supplier"



            Testo_Mail = Testo_Mail & "</BODY>"


            Dim mySmtp As New SmtpClient
            Dim myMail As New MailMessage()



            mySmtp.UseDefaultCredentials = False
        mySmtp.Credentials = New Net.NetworkCredential(sender_mail, Homepage.password_mail)
        mySmtp.Host = "tirelli-net.mail.protection.outlook.com"
            mySmtp.Port = 25
            mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network
            mySmtp.EnableSsl = True


            myMail = New MailMessage()
            myMail.From = New MailAddress("Report@tirelli.net")

            myMail.To.Add("giovannitirelli@tirelli.net")
            myMail.To.Add("spareparts@tirelli.net")
            myMail.To.Add("denismarcon@tirelli.net")
            myMail.To.Add("giacomotirelli@tirelli.net")
            myMail.To.Add("chiaraandreussi@tirelli.net")

            myMail.Subject = "Tirelli's automatic spare parts report"
            myMail.IsBodyHtml = True
            myMail.Body = Testo_Mail


            mySmtp.Send(myMail)


        End Sub


        Public Sub nuovo_Invia_Mail_esterna(id As Integer)

            Dim i = 0

            Do While i < DataGridView1.RowCount - 1

                mail_destinatario = DataGridView1.Rows(i).Cells(columnName:="mail").Value
                codice_bp_new = DataGridView1.Rows(i).Cells(columnName:="Codice_BP").Value
                nome_bp_destinatario = DataGridView1.Rows(i).Cells(columnName:="Nome_bP").Value


                Dim Testo_Mail As String
                ' Testo_Mail = "<BODY><H3></h3><P>"
                '  Testo_Mail = Testo_Mail & ""



                ' Testo_Mail = Testo_Mail & "<BR><BR>"
                Testo_Mail = Testo_Mail & "This is an automatic message. Please contact spareparts@tirelli.net for additional details.<BR>"

                Testo_Mail = Testo_Mail & "<BR><H3>" & nome_bp_destinatario & "</h3>"

                Testo_Mail = Testo_Mail & "</P></BODY>"

                Testo_Mail = Testo_Mail & "<table border=5>"
                Testo_Mail = Testo_Mail & "<tr>"
                Testo_Mail = Testo_Mail & "<th>Order N°</th>"
                Testo_Mail = Testo_Mail & "<th>Your PO N°</th>"
                Testo_Mail = Testo_Mail & "<th>ItemCode</th>"
                Testo_Mail = Testo_Mail & "<th>ItemName</th>"

                Testo_Mail = Testo_Mail & "<th>Description</th>"

                Testo_Mail = Testo_Mail & "<th>Insert date</th>"
                Testo_Mail = Testo_Mail & "<th>Our Due date</th>"
            '  Testo_Mail = Testo_Mail & "<th>Causal</th>"
            'Testo_Mail = Testo_Mail & "<th>Branch</th>"
            Testo_Mail = Testo_Mail & "<th>Customer</th>"

            'Testo_Mail = Testo_Mail & "<th>Open Qty</th>"
            'Testo_Mail = Testo_Mail & "<th>Status</th>"
            '  Testo_Mail = Testo_Mail & "<th>Supplier Due date</th>"
            'Testo_Mail = Testo_Mail & "<th>Toolshop due date</th>"

            Testo_Mail = Testo_Mail & "</tr>"



            Dim CNN As New SqlConnection
            CNN.ConnectionString = homepage.sap_tirelli
            cnn.Open()

            Dim CMD_SAP As New SqlCommand
                Dim cmd_SAP_reader As SqlDataReader

            CMD_SAP.Connection = cnn
            CMD_SAP.CommandText = "declare @codice_bp as varchar(10)
set @codice_bp = '" & codice_bp_new & "'

Select T40.DOCNUM as 'Order n°',t40.numatcard as 'Your reference N°',T40.[ItemCode],t40.Codice_brb,t40.substitute, T40.[ItemName], T40.[cliente] ,t40.u_descing as 'English name',   t40.docdate as 'Insert date', t40.shipdate as 'Our Due date', t40.u_causcons as 'Causal', t40.cardname as 'Customer', t40.openqty as 'Open Qty', t40.status, min (t41.shipdate) as 'Supplier Due date', cast(min(t42.u_data_cons_mes) as date) as 'Toolshop due date', case when t40.cardname is null then t40.cliente else t40.cardname end as 'Cliente_report' 
from
(
Select T30.DOCNUM,t30.numatcard,T30.[ItemCode],t30.Codice_brb,t30.substitute, T30.[ItemName],T30.[cliente], t30.u_descing, T30.[ItmsGrpNam], t30.docdate, t30.shipdate, t30.u_causcons, t30.cardname, t30.openqty, t30.u_trasferito, t30.u_datrasferire, t30.Mag, t30.Clavter, t30.Disp, case when t30.u_datrasferire = 0 then 'Picked'  when t30.mag >=t30.u_datrasferire then 'Warehouse' when t30.clavter>= t30.u_datrasferire then 'On treatment'  when t30.disp >=0 then 'On order' else 'Yet to be ordered' end as 'Status'
from
(
Select T20.DOCNUM,t20.numatcard,T20.[ItemCode],t20.Codice_brb,t20.substitute, T20.[ItemName],T20.[cliente],t20.u_descing,  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter, sum(t21.onhand-t21.iscommited+t21.onorder) as 'Disp'

from
(
Select T10.DOCNUM,t10.numatcard,T10.[ItemCode],t10.Codice_brb,t10.substitute, T10.[ItemName],T10.[cliente], t10.u_descing, T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, sum(t11.onhand) as 'Mag', case when t12.onhand is null then 0 else t12.onhand end as 'Clavter'


FROM
(
SELECT T0.DOCNUM,t0.numatcard,T2.[ItemCode],coalesce(t2.u_codice_brb,'') as 'Codice_brb', coalesce(t5.Substitute,'') as 'substitute', T1.[dscription] as 'Itemname',t0.cardname as 'cliente', t2.u_descing,  T3.[ItmsGrpNam], t0.docdate, t1.shipdate, case when t0.u_causcons ='V' then 'Sold' when t0.u_causcons ='GAR' then 'Warranty' else t0.u_causcons end as 'u_causcons' , t4.cardname, t1.openqty, t1.u_trasferito, t1.u_datrasferire

FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] 
INNER JOIN OITB T3 ON T2.[ItmsGrpCod] = T3.[ItmsGrpCod] 
left JOIN OCRD T4 ON T0.[u_codicebp] = T4.[CardCode]
left join oscn t5 on t5.itemcode=t1.itemcode and t0.cardcode=t5.cardcode

WHERE T1.[LineStatus] ='O' and (substring(t2.itemcode,1,1)='0' or substring(t2.itemcode,1,1)='C' or substring(t2.itemcode,1,1)='D') and (t0.cardcode=@codice_bp or t0.u_codicebp=@codice_bp ) and t0.u_causcons<>'COMP'
)
as t10 left join oitw t11 on t11.itemcode=t10.itemcode and t11.whscode<>'WIP' and t11.whscode<>'Clavter'
left join oitw t12 on t12.itemcode=t10.itemcode and t12.whscode='Clavter'

group by T10.DOCNUM,t10.numatcard,T10.[ItemCode], T10.[ItemName],T10.[cliente],  T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, t12.onhand,t10.u_descing,t10.Codice_brb,t10.substitute
)
as t20 left join oitw t21 on t20.itemcode=t21.itemcode

group by T20.DOCNUM,t20.numatcard,T20.[ItemCode],t20.Codice_brb,t20.substitute, T20.[ItemName],T20.[cliente],  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter,t20.u_descing
)
as t30
)
as t40 left join por1 t41 on t41.itemcode=t40.itemcode and t41.openqty>0 and t40.status ='On order'
left join owor t42 on t42.itemcode=t40.itemcode and (t42.status='P' or t42.status='R') and substring(t42.u_produzione,1,3)='Int' and t40.status ='On order'

group by T40.DOCNUM,t40.numatcard,T40.[ItemCode],t40.Codice_brb,t40.substitute, T40.[ItemName],T40.[cliente],  T40.[ItmsGrpNam], t40.docdate, t40.shipdate, t40.u_causcons, t40.cardname, t40.openqty, t40.u_trasferito, t40.u_datrasferire, t40.Mag, t40.Clavter, t40.Disp, t40.status,t40.u_descing

order by T40.DOCNUM"

                cmd_SAP_reader = CMD_SAP.ExecuteReader


                Do While cmd_SAP_reader.Read()

                    Testo_Mail = Testo_Mail & "<tr>"

                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Order n°") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Your reference N°") & "</td>"

                If cmd_SAP_reader("Substitute") <> "" Then
                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Substitute") & " <br> " & cmd_SAP_reader("itemcode") & "</td>"
                Else
                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemCode") & "</td>"
                End If

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemName") & "</td>"



                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("English name") & "</td>"

                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Insert date") & "</td>"

                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Our Due date") & "</td>"

                ' Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Causal") & "</td>"



                'Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Branch") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Cliente_report") & "</td>"

                'Testo_Mail = Testo_Mail & "<td>" & FormatNumber(cmd_SAP_reader("Open Qty"), 2, , , TriState.True) & "</td>"

                ' Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Status") & "</td>"

                '   Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Supplier Due date") & "</td>"

                'Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Toolshop due date") & "</td>"
                Testo_Mail = Testo_Mail & "</tr>"
                Loop
                cmd_SAP_reader.Close()
            cnn.Close()

            Testo_Mail = Testo_Mail & "</table>"


            'Testo_Mail = Testo_Mail & "<BR><H3>Status legend:</h3>"
            'Testo_Mail = Testo_Mail & "<BR>Picked = Ready for shipment"
            'Testo_Mail = Testo_Mail & "<BR>Warehouse = Available for picking"
            'Testo_Mail = Testo_Mail & "<BR>On order = Parts not received or produced yet"
            'Testo_Mail = Testo_Mail & "<BR>Yet to be ordered = Looking for supplier"



            Testo_Mail = Testo_Mail & "</BODY>"


                Dim mySmtp As New SmtpClient
                Dim myMail As New MailMessage()
                mySmtp.UseDefaultCredentials = False
            mySmtp.Credentials = New Net.NetworkCredential("report@tirelli.net", Homepage.password_mail)
            mySmtp.Host = "tirelli-net.mail.protection.outlook.com"
                mySmtp.Port = 25
            mySmtp.EnableSsl = False
            mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network



            myMail = New MailMessage()
                myMail.From = New MailAddress("Report@tirelli.net")

                myMail.To.Add(mail_destinatario)



                myMail.Subject = "Tirelli's automatic spare parts report customer: " & nome_bp_destinatario & ""
                myMail.IsBodyHtml = True
                myMail.Body = Testo_Mail


                mySmtp.Send(myMail)
                i = i + 1
                Testo_Mail = ""
            Loop


        End Sub

        Public Sub nuovo_Invia_Mail_interna(id As Integer)

            Dim i = 0

            Do While i < DataGridView1.RowCount - 1

                mail_destinatario = DataGridView1.Rows(i).Cells(columnName:="mail").Value
                codice_bp_new = DataGridView1.Rows(i).Cells(columnName:="Codice_BP").Value
                nome_bp_destinatario = DataGridView1.Rows(i).Cells(columnName:="Nome_bP").Value

                Dim Testo_Mail As String
                ' Testo_Mail = "<BODY><H3></h3><P>"
                '  Testo_Mail = Testo_Mail & ""



                ' Testo_Mail = Testo_Mail & "<BR><BR>"
                Testo_Mail = Testo_Mail & "This is an automatic message. Please contact spareparts@tirelli.net for additional details.<BR>"

                Testo_Mail = Testo_Mail & "<BR><H3>" & nome_bp_destinatario & "</h3>"

                Testo_Mail = Testo_Mail & "</P></BODY>"

                Testo_Mail = Testo_Mail & "<table border=5>"
                Testo_Mail = Testo_Mail & "<tr>"
                Testo_Mail = Testo_Mail & "<th>Order N°</th>"
                Testo_Mail = Testo_Mail & "<th>Your PO N°</th>"
                Testo_Mail = Testo_Mail & "<th>ItemCode</th>"
                Testo_Mail = Testo_Mail & "<th>ItemName</th>"

                Testo_Mail = Testo_Mail & "<th>Description</th>"

                Testo_Mail = Testo_Mail & "<th>Insert date</th>"
                Testo_Mail = Testo_Mail & "<th>Our Due date</th>"
                Testo_Mail = Testo_Mail & "<th>Causal</th>"
                'Testo_Mail = Testo_Mail & "<th>Branch</th>"
                Testo_Mail = Testo_Mail & "<th>Customer</th>"

                Testo_Mail = Testo_Mail & "<th>Open Qty</th>"
                Testo_Mail = Testo_Mail & "<th>Status</th>"
                Testo_Mail = Testo_Mail & "<th>Supplier Due date</th>"
                Testo_Mail = Testo_Mail & "<th>Toolshop due date</th>"

                Testo_Mail = Testo_Mail & "</tr>"



            Dim CNN As New SqlConnection
            CNN.ConnectionString = homepage.sap_tirelli
            cnn.Open()

            Dim CMD_SAP As New SqlCommand
                Dim cmd_SAP_reader As SqlDataReader

            CMD_SAP.Connection = cnn
            CMD_SAP.CommandText = "declare @codice_bp as varchar(10)
set @codice_bp = '" & codice_bp_new & "'

Select T40.DOCNUM as 'Order n°',t40.numatcard as 'Your PO N°',T40.[ItemCode],t40.Codice_brb,t40.substitute, T40.[ItemName], T40.[cliente] ,t40.u_descing as 'English name',   t40.docdate as 'Insert date', t40.shipdate as 'Our Due date', t40.u_causcons as 'Causal', t40.cardname as 'Customer', t40.openqty as 'Open Qty', t40.status, min (t41.shipdate) as 'Supplier Due date', cast(min(t42.u_data_cons_mes) as date) as 'Toolshop due date', case when t40.cardname is null then t40.cliente else t40.cardname end as 'Cliente_report' 
from
(
Select T30.DOCNUM,t30.numatcard,T30.[ItemCode],t30.Codice_brb,t30.substitute, T30.[ItemName],T30.[cliente], t30.u_descing, T30.[ItmsGrpNam], t30.docdate, t30.shipdate, t30.u_causcons, t30.cardname, t30.openqty, t30.u_trasferito, t30.u_datrasferire, t30.Mag, t30.Clavter, t30.Disp, case when t30.u_datrasferire = 0 then 'Picked'  when t30.mag >=t30.u_datrasferire then 'Warehouse' when t30.clavter>= t30.u_datrasferire then 'On treatment'  when t30.disp >=0 then 'On order' else 'Yet to be ordered' end as 'Status'
from
(
Select T20.DOCNUM,t20.numatcard,T20.[ItemCode],t20.Codice_brb,t20.substitute, T20.[ItemName],T20.[cliente],t20.u_descing,  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter, sum(t21.onhand-t21.iscommited+t21.onorder) as 'Disp'

from
(
Select T10.DOCNUM,t10.numatcard,T10.[ItemCode],t10.Codice_brb,t10.substitute, T10.[ItemName],T10.[cliente], t10.u_descing, T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, sum(t11.onhand) as 'Mag', case when t12.onhand is null then 0 else t12.onhand end as 'Clavter'


FROM
(
SELECT T0.DOCNUM,t0.numatcard,T2.[ItemCode],coalesce(t2.u_codice_brb,'') as 'Codice_brb', coalesce(t5.Substitute,'') as 'substitute', T1.[dscription] as 'Itemname',t0.cardname as 'cliente', t2.u_descing,  T3.[ItmsGrpNam], t0.docdate, t1.shipdate, case when t0.u_causcons ='V' then 'Sold' when t0.u_causcons ='GAR' then 'Warranty' else t0.u_causcons end as 'u_causcons' , t4.cardname, t1.openqty, t1.u_trasferito, t1.u_datrasferire

FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] 
INNER JOIN OITB T3 ON T2.[ItmsGrpCod] = T3.[ItmsGrpCod] 
left JOIN OCRD T4 ON T0.[u_codicebp] = T4.[CardCode]
left join oscn t5 on t5.itemcode=t1.itemcode and t0.cardcode=t5.cardcode

WHERE T1.[LineStatus] ='O' and (substring(t2.itemcode,1,1)='0' or substring(t2.itemcode,1,1)='C' or substring(t2.itemcode,1,1)='D') and (t0.cardcode=@codice_bp or t0.u_codicebp=@codice_bp ) and t0.u_causcons<>'COMP'
)
as t10 left join oitw t11 on t11.itemcode=t10.itemcode and t11.whscode<>'WIP' and t11.whscode<>'Clavter'
left join oitw t12 on t12.itemcode=t10.itemcode and t12.whscode='Clavter'

group by T10.DOCNUM,t10.numatcard,T10.[ItemCode], T10.[ItemName],T10.[cliente],  T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, t12.onhand,t10.u_descing,t10.Codice_brb,t10.substitute
)
as t20 left join oitw t21 on t20.itemcode=t21.itemcode

group by T20.DOCNUM,t20.numatcard,T20.[ItemCode],t20.Codice_brb,t20.substitute, T20.[ItemName],T20.[cliente],  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter,t20.u_descing
)
as t30
)
as t40 left join por1 t41 on t41.itemcode=t40.itemcode and t41.openqty>0 and t40.status ='On order'
left join owor t42 on t42.itemcode=t40.itemcode and (t42.status='P' or t42.status='R') and substring(t42.u_produzione,1,3)='Int' and t40.status ='On order'

group by T40.DOCNUM,t40.numatcard,T40.[ItemCode],t40.Codice_brb,t40.substitute, T40.[ItemName],T40.[cliente],  T40.[ItmsGrpNam], t40.docdate, t40.shipdate, t40.u_causcons, t40.cardname, t40.openqty, t40.u_trasferito, t40.u_datrasferire, t40.Mag, t40.Clavter, t40.Disp, t40.status,t40.u_descing

order by T40.DOCNUM"

                cmd_SAP_reader = CMD_SAP.ExecuteReader


                Do While cmd_SAP_reader.Read()

                    Testo_Mail = Testo_Mail & "<tr>"

                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Order n°") & "</td>"
                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Your PO N°") & "</td>"


                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemCode") & "</td>"

                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemName") & "</td>"



                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("English name") & "</td>"

                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Insert date") & "</td>"

                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Our Due date") & "</td>"

                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Causal") & "</td>"



                    'Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Branch") & "</td>"
                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Cliente_report") & "</td>"

                    Testo_Mail = Testo_Mail & "<td>" & FormatNumber(cmd_SAP_reader("Open Qty"), 2, , , TriState.True) & "</td>"

                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Status") & "</td>"

                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Supplier Due date") & "</td>"

                    Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Toolshop due date") & "</td>"
                    Testo_Mail = Testo_Mail & "</tr>"
                Loop
                cmd_SAP_reader.Close()
            cnn.Close()

            Testo_Mail = Testo_Mail & "</table>"


                Testo_Mail = Testo_Mail & "<BR><H3>Status legend:</h3>"
                Testo_Mail = Testo_Mail & "<BR>Picked = Ready for shipment"
                Testo_Mail = Testo_Mail & "<BR>Warehouse = Available for picking"
                Testo_Mail = Testo_Mail & "<BR>On order = Parts not received or produced yet"
                Testo_Mail = Testo_Mail & "<BR>Yet to be ordered = Looking for supplier"



                Testo_Mail = Testo_Mail & "</BODY>"


                Dim mySmtp As New SmtpClient
                Dim myMail As New MailMessage()
                mySmtp.UseDefaultCredentials = False
                mySmtp.Credentials = New Net.NetworkCredential("report@tirelli.net", "Ras70773")
                mySmtp.Host = "smtp.office365.com"
                mySmtp.Port = 25
                mySmtp.EnableSsl = True
                mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network


                myMail = New MailMessage()
                myMail.From = New MailAddress("Report@tirelli.net")

            myMail.To.Add("spareparts@tirelli.net")







            myMail.Subject = "Tirelli's automatic spare parts report customer: " & nome_bp_destinatario & ""
                myMail.IsBodyHtml = True
                myMail.Body = Testo_Mail


                mySmtp.Send(myMail)
                i = i + 1
                Testo_Mail = ""
            Loop


        End Sub

        Public Sub Invia_Mail_riga_selezionata(id As Integer)



            mail_destinatario = DataGridView1.Rows(riga).Cells(columnName:="mail").Value
            codice_bp_new = DataGridView1.Rows(riga).Cells(columnName:="Codice_BP").Value
            nome_bp_destinatario = DataGridView1.Rows(riga).Cells(columnName:="Nome_bP").Value

            Dim Testo_Mail As String
            ' Testo_Mail = "<BODY><H3></h3><P>"
            '  Testo_Mail = Testo_Mail & ""



            ' Testo_Mail = Testo_Mail & "<BR><BR>"
            Testo_Mail = Testo_Mail & "This is an automatic message. Please contact spareparts@tirelli.net for additional details.<BR>"

            Testo_Mail = Testo_Mail & "<BR><H3>" & nome_bp_destinatario & "</h3>"

            Testo_Mail = Testo_Mail & "</P></BODY>"

            Testo_Mail = Testo_Mail & "<table border=5>"
            Testo_Mail = Testo_Mail & "<tr>"
            Testo_Mail = Testo_Mail & "<th>Order N°</th>"
            Testo_Mail = Testo_Mail & "<th>Your PO N°</th>"
            Testo_Mail = Testo_Mail & "<th>ItemCode</th>"
            Testo_Mail = Testo_Mail & "<th>ItemName</th>"

            Testo_Mail = Testo_Mail & "<th>Description</th>"

            Testo_Mail = Testo_Mail & "<th>Insert date</th>"
            Testo_Mail = Testo_Mail & "<th>Our Due date</th>"
            Testo_Mail = Testo_Mail & "<th>Causal</th>"
            'Testo_Mail = Testo_Mail & "<th>Branch</th>"
            Testo_Mail = Testo_Mail & "<th>Customer</th>"

            Testo_Mail = Testo_Mail & "<th>Open Qty</th>"
            Testo_Mail = Testo_Mail & "<th>Status</th>"
            Testo_Mail = Testo_Mail & "<th>Supplier Due date</th>"
            Testo_Mail = Testo_Mail & "<th>Toolshop due date</th>"

            Testo_Mail = Testo_Mail & "</tr>"



        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "declare @codice_bp as varchar(10)
set @codice_bp = '" & codice_bp_new & "'

Select T40.DOCNUM as 'Order n°',t40.numatcard as 'Your PO N°',T40.[ItemCode],t40.Codice_brb,t40.substitute, T40.[ItemName], T40.[cliente] ,t40.u_descing as 'English name',   t40.docdate as 'Insert date', t40.shipdate as 'Our Due date', t40.u_causcons as 'Causal', t40.cardname as 'Customer', t40.openqty as 'Open Qty', t40.status, min (t41.shipdate) as 'Supplier Due date', cast(min(t42.u_data_cons_mes) as date) as 'Toolshop due date', case when t40.cardname is null then t40.cliente else t40.cardname end as 'Cliente_report' 
from
(
Select T30.DOCNUM,t30.numatcard,T30.[ItemCode],t30.Codice_brb,t30.substitute, T30.[ItemName],T30.[cliente], t30.u_descing, T30.[ItmsGrpNam], t30.docdate, t30.shipdate, t30.u_causcons, t30.cardname, t30.openqty, t30.u_trasferito, t30.u_datrasferire, t30.Mag, t30.Clavter, t30.Disp, case when t30.u_datrasferire = 0 then 'Picked'  when t30.mag >=t30.u_datrasferire then 'Warehouse' when t30.clavter>= t30.u_datrasferire then 'On treatment'  when t30.disp >=0 then 'On order' else 'Yet to be ordered' end as 'Status'
from
(
Select T20.DOCNUM,t20.numatcard,T20.[ItemCode],t20.Codice_brb,t20.substitute, T20.[ItemName],T20.[cliente],t20.u_descing,  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter, sum(t21.onhand-t21.iscommited+t21.onorder) as 'Disp'

from
(
Select T10.DOCNUM,t10.numatcard,T10.[ItemCode],t10.Codice_brb,t10.substitute, T10.[ItemName],T10.[cliente], t10.u_descing, T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, sum(t11.onhand) as 'Mag', case when t12.onhand is null then 0 else t12.onhand end as 'Clavter'


FROM
(
SELECT T0.DOCNUM,t0.numatcard,T2.[ItemCode],coalesce(t2.u_codice_brb,'') as 'Codice_brb', coalesce(t5.Substitute,'') as 'substitute', T1.[dscription] as 'Itemname',t0.cardname as 'cliente', t2.u_descing,  T3.[ItmsGrpNam], t0.docdate, t1.shipdate, case when t0.u_causcons ='V' then 'Sold' when t0.u_causcons ='GAR' then 'Warranty' else t0.u_causcons end as 'u_causcons' , t4.cardname, t1.openqty, t1.u_trasferito, t1.u_datrasferire

FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] 
INNER JOIN OITB T3 ON T2.[ItmsGrpCod] = T3.[ItmsGrpCod] 
left JOIN OCRD T4 ON T0.[u_codicebp] = T4.[CardCode]
left join oscn t5 on t5.itemcode=t1.itemcode and t0.cardcode=t5.cardcode

WHERE T1.[LineStatus] ='O' and (substring(t2.itemcode,1,1)='0' or substring(t2.itemcode,1,1)='C' or substring(t2.itemcode,1,1)='D') and (t0.cardcode=@codice_bp or t0.u_codicebp=@codice_bp ) and t0.u_causcons<>'COMP'
)
as t10 left join oitw t11 on t11.itemcode=t10.itemcode and t11.whscode<>'WIP' and t11.whscode<>'Clavter'
left join oitw t12 on t12.itemcode=t10.itemcode and t12.whscode='Clavter'

group by T10.DOCNUM,t10.numatcard,T10.[ItemCode], T10.[ItemName],T10.[cliente],  T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, t12.onhand,t10.u_descing,t10.Codice_brb,t10.substitute
)
as t20 left join oitw t21 on t20.itemcode=t21.itemcode

group by T20.DOCNUM,t20.numatcard,T20.[ItemCode],t20.Codice_brb,t20.substitute, T20.[ItemName],T20.[cliente],  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter,t20.u_descing
)
as t30
)
as t40 left join por1 t41 on t41.itemcode=t40.itemcode and t41.openqty>0 and t40.status ='On order'
left join owor t42 on t42.itemcode=t40.itemcode and (t42.status='P' or t42.status='R') and substring(t42.u_produzione,1,3)='Int' and t40.status ='On order'

group by T40.DOCNUM,t40.numatcard,T40.[ItemCode],t40.Codice_brb,t40.substitute, T40.[ItemName],T40.[cliente],  T40.[ItmsGrpNam], t40.docdate, t40.shipdate, t40.u_causcons, t40.cardname, t40.openqty, t40.u_trasferito, t40.u_datrasferire, t40.Mag, t40.Clavter, t40.Disp, t40.status,t40.u_descing

order by T40.DOCNUM
"

            cmd_SAP_reader = CMD_SAP.ExecuteReader


            Do While cmd_SAP_reader.Read()

                Testo_Mail = Testo_Mail & "<tr>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Order n°") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Your PO N°") & "</td>"

            If cmd_SAP_reader("Substitute") <> "" Then
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Substitute") & " <br> " & cmd_SAP_reader("itemcode") & "</td>"
            Else
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemCode") & "</td>"
            End If


            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemName") & "</td>"



                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("English name") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Insert date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Our Due date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Causal") & "</td>"



                'Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Branch") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Cliente_report") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & FormatNumber(cmd_SAP_reader("Open Qty"), 2, , , TriState.True) & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Status") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Supplier Due date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Toolshop due date") & "</td>"
                Testo_Mail = Testo_Mail & "</tr>"
            Loop
            cmd_SAP_reader.Close()
        cnn.Close()

        Testo_Mail = Testo_Mail & "</table>"


            Testo_Mail = Testo_Mail & "<BR><H3>Status legend:</h3>"
            Testo_Mail = Testo_Mail & "<BR>Picked = Ready for shipment"
            Testo_Mail = Testo_Mail & "<BR>Warehouse = Available for picking"
            Testo_Mail = Testo_Mail & "<BR>On order = Parts not received or produced yet"
            Testo_Mail = Testo_Mail & "<BR>Yet to be ordered = Looking for supplier"



            Testo_Mail = Testo_Mail & "</BODY>"


            Dim mySmtp As New SmtpClient
            Dim myMail As New MailMessage()
            mySmtp.UseDefaultCredentials = False
            mySmtp.Credentials = New Net.NetworkCredential("report@tirelli.net", "Ras70773")
        mySmtp.Host = "tirelli-net.mail.protection.outlook.com"
        mySmtp.Port = 25
            mySmtp.EnableSsl = True
            mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network


            myMail = New MailMessage()
            myMail.From = New MailAddress("Report@tirelli.net")

            myMail.To.Add(mail_destinatario)







            myMail.Subject = "Tirelli's automatic spare parts report customer: " & nome_bp_destinatario & ""
            myMail.IsBodyHtml = True
            myMail.Body = Testo_Mail


            mySmtp.Send(myMail)

            Testo_Mail = ""



        End Sub


        Public Sub Invia_Mail_esterna(id As Integer)




            Dim Testo_Mail As String
            ' Testo_Mail = "<BODY><H3></h3><P>"
            '  Testo_Mail = Testo_Mail & ""



            ' Testo_Mail = Testo_Mail & "<BR><BR>"
            Testo_Mail = Testo_Mail & "This is an automatic message. Please contact spareparts@tirelli.net for additional details.<BR>"

            Testo_Mail = Testo_Mail & "<BR><H3>AROL NORTH AMERICA</h3>"

            Testo_Mail = Testo_Mail & "</P></BODY>"

            Testo_Mail = Testo_Mail & "<table border=5>"
            Testo_Mail = Testo_Mail & "<tr>"
            Testo_Mail = Testo_Mail & "<th>Order N°</th>"
            Testo_Mail = Testo_Mail & "<th>Your PO N°</th>"
            Testo_Mail = Testo_Mail & "<th>ItemCode</th>"
            Testo_Mail = Testo_Mail & "<th>ItemName</th>"

            Testo_Mail = Testo_Mail & "<th>Description</th>"

            Testo_Mail = Testo_Mail & "<th>Insert date</th>"
            Testo_Mail = Testo_Mail & "<th>Our Due date</th>"
            Testo_Mail = Testo_Mail & "<th>Causal</th>"
            'Testo_Mail = Testo_Mail & "<th>Branch</th>"
            Testo_Mail = Testo_Mail & "<th>Customer</th>"

            Testo_Mail = Testo_Mail & "<th>Open Qty</th>"
            Testo_Mail = Testo_Mail & "<th>Status</th>"
            Testo_Mail = Testo_Mail & "<th>Supplier Due date</th>"
            Testo_Mail = Testo_Mail & "<th>Toolshop due date</th>"

            Testo_Mail = Testo_Mail & "</tr>"
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select T40.DOCNUM as 'Order n°',t40.numatcard as 'Your PO N°',T40.[ItemCode], T40.[ItemName], T40.[Branch] ,t40.u_descing as 'English name',   t40.docdate as 'Insert date', t40.shipdate as 'Our Due date', t40.u_causcons as 'Causal', t40.cardname as 'Customer', t40.openqty as 'Open Qty', t40.status, min (t41.shipdate) as 'Supplier Due date', cast(min(t42.u_data_cons_mes) as date) as 'Toolshop due date'
from
(
Select T30.DOCNUM,t30.numatcard,T30.[ItemCode], T30.[ItemName],T30.[Branch], t30.u_descing, T30.[ItmsGrpNam], t30.docdate, t30.shipdate, t30.u_causcons, t30.cardname, t30.openqty, t30.u_trasferito, t30.u_datrasferire, t30.Mag, t30.Clavter, t30.Disp, case when t30.u_datrasferire = 0 then 'Picked'  when t30.mag >=t30.u_datrasferire then 'Warehouse' when t30.clavter>= t30.u_datrasferire then 'On treatment'  when t30.disp >=0 then 'On order' else 'Yet to be ordered' end as 'Status'
from
(
Select T20.DOCNUM,t20.numatcard,T20.[ItemCode], T20.[ItemName],T20.[Branch],t20.u_descing,  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter, sum(t21.onhand-t21.iscommited+t21.onorder) as 'Disp'

from
(
Select T10.DOCNUM,t10.numatcard,T10.[ItemCode], T10.[ItemName],T10.[Branch], t10.u_descing, T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, sum(t11.onhand) as 'Mag', case when t12.onhand is null then 0 else t12.onhand end as 'Clavter'


FROM
(
SELECT T0.DOCNUM,t0.numatcard,T2.[ItemCode], T2.[ItemName],t0.cardname as 'Branch', t2.u_descing,  T3.[ItmsGrpNam], t0.docdate, t1.shipdate, case when t0.u_causcons ='V' then 'Sold' when t0.u_causcons ='GAR' then 'Warranty' else t0.u_causcons end as 'u_causcons' , t4.cardname, t1.openqty, t1.u_trasferito, t1.u_datrasferire

FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] INNER JOIN OITB T3 ON T2.[ItmsGrpCod] = T3.[ItmsGrpCod] 
left JOIN OCRD T4 ON T0.[u_codicebp] = T4.[CardCode]

WHERE T1.[LineStatus] ='O' and (substring(t2.itemcode,1,1)='0' or substring(t2.itemcode,1,1)='C' or substring(t2.itemcode,1,1)='D') and (t0.cardcode='01872') and t0.u_causcons<>'COMP'
)
as t10 left join oitw t11 on t11.itemcode=t10.itemcode and t11.whscode<>'WIP' and t11.whscode<>'Clavter'
left join oitw t12 on t12.itemcode=t10.itemcode and t12.whscode='Clavter'

group by T10.DOCNUM,t10.numatcard,T10.[ItemCode], T10.[ItemName],T10.[Branch],  T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, t12.onhand,t10.u_descing
)
as t20 left join oitw t21 on t20.itemcode=t21.itemcode

group by T20.DOCNUM,t20.numatcard,T20.[ItemCode], T20.[ItemName],T20.[Branch],  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter,t20.u_descing
)
as t30
)
as t40 left join por1 t41 on t41.itemcode=t40.itemcode and t41.openqty>0 and t40.status ='On order'
left join owor t42 on t42.itemcode=t40.itemcode and (t42.status='P' or t42.status='R') and t42.u_produzione='Int' and t40.status ='On order'

group by T40.DOCNUM,t40.numatcard,T40.[ItemCode], T40.[ItemName],T40.[Branch],  T40.[ItmsGrpNam], t40.docdate, t40.shipdate, t40.u_causcons, t40.cardname, t40.openqty, t40.u_trasferito, t40.u_datrasferire, t40.Mag, t40.Clavter, t40.Disp, t40.status,t40.u_descing

order by T40.DOCNUM  "

            cmd_SAP_reader = CMD_SAP.ExecuteReader


            Do While cmd_SAP_reader.Read()

                Testo_Mail = Testo_Mail & "<tr>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Order n°") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Your PO N°") & "</td>"


                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemCode") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemName") & "</td>"



                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("English name") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Insert date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Our Due date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Causal") & "</td>"



                'Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Branch") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Customer") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & FormatNumber(cmd_SAP_reader("Open Qty"), 2, , , TriState.True) & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Status") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Supplier Due date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Toolshop due date") & "</td>"
                Testo_Mail = Testo_Mail & "</tr>"
            Loop
            cmd_SAP_reader.Close()
        cnn.Close()

        Testo_Mail = Testo_Mail & "</table>"




            Testo_Mail = Testo_Mail & "<BR><BR>"
            Testo_Mail = Testo_Mail & "<BR><h3>AROL LATIN AMERICA</h3>"
            Testo_Mail = Testo_Mail & "<BR><BR>"
            Testo_Mail = Testo_Mail & "</P></BODY>"

            Testo_Mail = Testo_Mail & "<table border=5>"
            Testo_Mail = Testo_Mail & "<tr>"
            Testo_Mail = Testo_Mail & "<th>Order N°</th>"
            Testo_Mail = Testo_Mail & "<th>Your PO N°</th>"
            Testo_Mail = Testo_Mail & "<th>ItemCode</th>"
            Testo_Mail = Testo_Mail & "<th>ItemName</th>"

            Testo_Mail = Testo_Mail & "<th>Description</th>"
            Testo_Mail = Testo_Mail & "<th>Insert date</th>"
            Testo_Mail = Testo_Mail & "<th>Our Due date</th>"
            Testo_Mail = Testo_Mail & "<th>Causal</th>"
            'Testo_Mail = Testo_Mail & "<th>Branch</th>"
            Testo_Mail = Testo_Mail & "<th>Customer</th>"

            Testo_Mail = Testo_Mail & "<th>Open Qty</th>"
            Testo_Mail = Testo_Mail & "<th>Status</th>"
            Testo_Mail = Testo_Mail & "<th>Supplier Due date</th>"
            Testo_Mail = Testo_Mail & "<th>Toolshop due date</th>"

            Testo_Mail = Testo_Mail & "</tr>"

        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select T40.DOCNUM as 'Order n°',t40.numatcard as 'Your PO N°',T40.[ItemCode], T40.[ItemName], T40.[Branch] ,t40.u_descing as 'English name',   t40.docdate as 'Insert date', t40.shipdate as 'Our Due date', t40.u_causcons as 'Causal', t40.cardname as 'Customer', t40.openqty as 'Open Qty', t40.status, min (t41.shipdate) as 'Supplier Due date', cast(min(t42.u_data_cons_mes) as date) as 'Toolshop due date'
from
(
Select T30.DOCNUM,t30.numatcard,T30.[ItemCode], T30.[ItemName],T30.[Branch], t30.u_descing, T30.[ItmsGrpNam], t30.docdate, t30.shipdate, t30.u_causcons, t30.cardname, t30.openqty, t30.u_trasferito, t30.u_datrasferire, t30.Mag, t30.Clavter, t30.Disp, case when t30.u_datrasferire = 0 then 'Picked'  when t30.mag >=t30.u_datrasferire then 'Warehouse' when t30.clavter>= t30.u_datrasferire then 'On treatment'  when t30.disp >=0 then 'On order' else 'Yet to be ordered' end as 'Status'
from
(
Select T20.DOCNUM,t20.numatcard,T20.[ItemCode], T20.[ItemName],T20.[Branch],t20.u_descing,  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter, sum(t21.onhand-t21.iscommited+t21.onorder) as 'Disp'

from
(
Select T10.DOCNUM,t10.numatcard,T10.[ItemCode], T10.[ItemName],T10.[Branch], t10.u_descing, T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, sum(t11.onhand) as 'Mag', case when t12.onhand is null then 0 else t12.onhand end as 'Clavter'


FROM
(
SELECT T0.DOCNUM,t0.numatcard,T2.[ItemCode], T2.[ItemName],t0.cardname as 'Branch', t2.u_descing,  T3.[ItmsGrpNam], t0.docdate, t1.shipdate, case when t0.u_causcons ='V' then 'Sold' when t0.u_causcons ='GAR' then 'Warranty' else t0.u_causcons end as 'u_causcons', t4.cardname, t1.openqty, t1.u_trasferito, t1.u_datrasferire

FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] INNER JOIN OITB T3 ON T2.[ItmsGrpCod] = T3.[ItmsGrpCod] 
left JOIN OCRD T4 ON T0.[u_codicebp] = T4.[CardCode]

WHERE T1.[LineStatus] ='O' and (substring(t2.itemcode,1,1)='0' or substring(t2.itemcode,1,1)='C' or substring(t2.itemcode,1,1)='D') and (t0.cardcode='01924') and t0.u_causcons<>'COMP'
)
as t10 left join oitw t11 on t11.itemcode=t10.itemcode and t11.whscode<>'WIP' and t11.whscode<>'Clavter'
left join oitw t12 on t12.itemcode=t10.itemcode and t12.whscode='Clavter'

group by T10.DOCNUM,t10.numatcard,T10.[ItemCode], T10.[ItemName],T10.[Branch],  T10.[ItmsGrpNam], t10.docdate, t10.shipdate, t10.u_causcons, t10.cardname, t10.openqty, t10.u_trasferito, t10.u_datrasferire, t12.onhand,t10.u_descing
)
as t20 left join oitw t21 on t20.itemcode=t21.itemcode

group by T20.DOCNUM,t20.numatcard,T20.[ItemCode], T20.[ItemName],T20.[Branch],  T20.[ItmsGrpNam], t20.docdate, t20.shipdate, t20.u_causcons, t20.cardname, t20.openqty, t20.u_trasferito, t20.u_datrasferire, t20.Mag, t20.Clavter,t20.u_descing
)
as t30
)
as t40 left join por1 t41 on t41.itemcode=t40.itemcode and t41.openqty>0 and t40.status ='On order'
left join owor t42 on t42.itemcode=t40.itemcode and (t42.status='P' or t42.status='R') and t42.u_produzione='Int' and t40.status ='On order'

group by T40.DOCNUM,t40.numatcard,T40.[ItemCode], T40.[ItemName],T40.[Branch],  T40.[ItmsGrpNam], t40.docdate, t40.shipdate, t40.u_causcons, t40.cardname, t40.openqty, t40.u_trasferito, t40.u_datrasferire, t40.Mag, t40.Clavter, t40.Disp, t40.status,t40.u_descing

order by T40.DOCNUM  "

            cmd_SAP_reader = CMD_SAP.ExecuteReader


            Do While cmd_SAP_reader.Read()

                Testo_Mail = Testo_Mail & "<tr>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Order n°") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Your PO N°") & "</td>"


                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemCode") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemName") & "</td>"



                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("English name") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Insert date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Our Due date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Causal") & "</td>"



                'Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Branch") & "</td>"
                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Customer") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & FormatNumber(cmd_SAP_reader("Open Qty"), 2, , , TriState.True) & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Status") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Supplier Due date") & "</td>"

                Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Toolshop due date") & "</td>"
                Testo_Mail = Testo_Mail & "</tr>"
            Loop
            cmd_SAP_reader.Close()
        cnn.Close()

        Testo_Mail = Testo_Mail & "</table>"

            Testo_Mail = Testo_Mail & "<BR><H3>Status legend:</h3>"
            Testo_Mail = Testo_Mail & "<BR>Picked = Ready for shipment"
            Testo_Mail = Testo_Mail & "<BR>Warehouse = Available for picking"
            Testo_Mail = Testo_Mail & "<BR>On order = Parts not received or produced yet"
            Testo_Mail = Testo_Mail & "<BR>Yet to be ordered = Looking for supplier"



            Testo_Mail = Testo_Mail & "</BODY>"


            Dim mySmtp As New SmtpClient
            Dim myMail As New MailMessage()
            mySmtp.UseDefaultCredentials = False
            mySmtp.Credentials = New Net.NetworkCredential("report@tirelli.net", "Ras70773")
            mySmtp.Host = "smtp.office365.com"
            mySmtp.Port = 25
            mySmtp.EnableSsl = True
            mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network


            myMail = New MailMessage()
            myMail.From = New MailAddress("Report@tirelli.net")

            myMail.To.Add("benedito.pinto@arol.com")
            myMail.To.Add("Navaneeth.Kumar@arol.com")
            myMail.To.Add("nemias.florencio@arol.com")
            myMail.To.Add("nemias.florencio@arol.com")

            myMail.To.Add("giovannitirelli@tirelli.net")
            myMail.To.Add("spareparts@tirelli.net")
            myMail.To.Add("denismarcon@tirelli.net")
            myMail.To.Add("giacomotirelli@tirelli.net")
            myMail.To.Add("chiaraandreussi@tirelli.net")

            myMail.Subject = "Tirelli's automatic spare parts report"
            myMail.IsBodyHtml = True
            myMail.Body = Testo_Mail


            mySmtp.Send(myMail)


        End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Invia_Mail_interna(1)
        nuovo_Invia_Mail_interna(1)
        MsgBox("Mail inviate con successo")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Invia_Mail_esterna(1)
        nuovo_Invia_Mail_esterna(1)
        MsgBox("Mail inviate con successo")
    End Sub

    Sub riempi_datagridview()

        DataGridView1.Rows.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli

        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "select t0.id, t0.cardcode, t0.mail, T1.CARDNAME 
from [Tirelli_40].[dbo].[Mail_ricambi] t0 LEFT JOIN OCRD T1 ON T0.CARDCODE=T1.CARDCODE where t0.cardcode   Like '%%" & TextBox3.Text & "%%' and t1.cardname Like '%%" & TextBox4.Text & "%%' and t0.mail   Like '%%" & TextBox5.Text & "%%' ORDER BY T1.CARDNAME "
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()
            DataGridView1.Rows.Add(cmd_SAP_reader("id"), cmd_SAP_reader("cardcode"), cmd_SAP_reader("CARDNAME"), cmd_SAP_reader("mail"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()
        DataGridView1.ClearSelection()

    End Sub

    Private Sub Pianificazione_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer_ricambi.Start()


        riempi_datagridview()
    End Sub

    Sub inserisci_record()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli


        cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "INSERT INTO [Tirelli_40].[dbo].[Mail_ricambi](ID,CARDCODE,MAIL) VALUES (" & ID_record & ",'" & TextBox1.Text & "', '" & TextBox2.Text & "')"
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()

    End Sub

    Sub elimina_record()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli


        cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "delete [Tirelli_40].[dbo].[Mail_ricambi] where id =" & ID_record_cancellazione & ""
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()

    End Sub

    Sub trova_ID()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "SELECT 'o', max(case when t0.id is null then 0 else t0.id end )+1 as 'ID' from [Tirelli_40].[dbo].[Mail_ricambi] t0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then

            If Not cmd_SAP_reader("ID") Is System.DBNull.Value Then
                ID_record = cmd_SAP_reader("ID")
            Else
                ID_record = 1
            End If
        Else
            ID_record = 1
        End If

        cnn.Close()
        cmd_SAP_reader.Close()


    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        trova_ID()
        inserisci_record()
        riempi_datagridview()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            ID_record_cancellazione = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ID").Value
            riga = e.RowIndex
        End If
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        elimina_record()
        riempi_datagridview()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        riempi_datagridview()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        riempi_datagridview()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        riempi_datagridview()
    End Sub

    Private Sub Timer1_Tick_1(sender As Object, e As EventArgs) Handles Timer_ricambi.Tick

        If Weekday(Now) = 6 Then

            If Mid(Now, 12, 2) = 10 And mail_inviate = 0 Then

                nuovo_Invia_Mail_esterna(1)

                mail_inviate = 1



            End If


        Else

            mail_inviate = 0

        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Invia_Mail_riga_selezionata(1)
        MsgBox("Mail inviata con successo")
    End Sub

    Private Sub Timer_tickets_Tick(sender As Object, e As EventArgs)
        Timer1_Timer()
    End Sub

    Sub Timer1_Timer()



        ' Ottiene l'ora corrente
        Dim ora As Date = TimeOfDay
        ' Imposta l'ora desiderata a 10:00
        Dim oraDesiderata As Date = CDate("09:00 AM")
        'Confronta l'ora corrente con l'ora desiderata
        Console.WriteLine(contatore & " ora " & ora)
        Console.WriteLine(contatore & " ora.hour " & ora.Hour)
        Console.WriteLine(contatore & " oraDesiderata.Hour " & oraDesiderata.Hour)

        Console.WriteLine(contatore & " ora.Minute " & ora.Minute)
        Console.WriteLine(contatore & " oraDesiderata.Minute " & oraDesiderata.Minute)

        If ora.Hour = oraDesiderata.Hour And ora.Minute = oraDesiderata.Minute And mostratoMessaggio = False Then
            ' Mostra una finestra di messaggio

            mail_destinatario = "giovannitirelli@tirelli.net"
            invia_mail_tickets(1)
            invia_mail_OC_garanzia(1)
            mail_destinatario = "francoporracin@tirelli.net"
            invia_mail_tickets(1)
            invia_mail_OC_garanzia(1)
            mail_destinatario = "carlotonini@tirelli.net"
            invia_mail_tickets(1)
            invia_mail_OC_garanzia(1)
            ' Imposta la variabile mostratoMessaggio a True
            mostratoMessaggio = True
        ElseIf ora.Hour <> oraDesiderata.Hour Then
            ' Se l'ora corrente non corrisponde all'ora desiderata, reimposta la variabile mostratoMessaggio a False
            mostratoMessaggio = False
        End If
        contatore = contatore + 1
    End Sub


    Sub invia_mail_tickets(id As Integer)

        Dim i = 0



        Dim Testo_Mail As String
        ' Testo_Mail = "<BODY><H3></h3><P>"
        '  Testo_Mail = Testo_Mail & ""



        ' Testo_Mail = Testo_Mail & "<BR><BR>"
        Testo_Mail = Testo_Mail & "Riepilogo Nuovi Tickets generati Ieri.<BR>"



        Testo_Mail = Testo_Mail & "</P></BODY>"

        Testo_Mail = Testo_Mail & "<table border=5>"
        Testo_Mail = Testo_Mail & "<tr>"
        Testo_Mail = Testo_Mail & "<th>ID Ticket</th>"
        Testo_Mail = Testo_Mail & "<th>Descrizione motivo</th>"
        Testo_Mail = Testo_Mail & "<th>Commessa</th>"
        Testo_Mail = Testo_Mail & "<th>Nome commessa</th>"

        Testo_Mail = Testo_Mail & "<th>Cliente</th>"

        Testo_Mail = Testo_Mail & "<th>Descrizione</th>"
        Testo_Mail = Testo_Mail & "<th>Reparto Mittente</th>"
        Testo_Mail = Testo_Mail & "<th>Reparto Destinatario</th>"
        'Testo_Mail = Testo_Mail & "<th>Branch</th>"
        Testo_Mail = Testo_Mail & "<th>Nome mittente</th>"



        Testo_Mail = Testo_Mail & "</tr>"

        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT *
from
(
select
t0.[Id_Ticket]
,t5.[Descrizione_Motivo]
      ,t0.[Commessa],  case WHEN t12.itemname is not null then t12.itemname when t6.itemname is null then '' else t6.itemname end as 'Itemname',
	  case when substring(t0.COMMESSA,1,3)='CDS' THEN case when t12.U_Final_customer_name is null then t11.custmrName else t12.U_Final_customer_name end when t7.cardname is null then t6.u_final_customer_name else t7.cardname end as 'Cliente'
      ,t0.[Data_Creazione]
	  , DATEDIFF(day,t0.[Data_Creazione], getdate()) as 'giorni'
      ,t0.[Data_Chiusura]
      ,t0.[Data_Prevista_Chiusura]
      ,t0.[Aperto]
      ,t0.[Descrizione] 
	  ,t4.Descrizione as 'Mittente_padre'
      ,t1.[descrizione] as 'Mittente'
      ,t2.[descrizione] as 'Destinatario'
      ,t0.[Immagine]
      ,t0.[Id_Padre]
      ,t0.[Business]
, t0.oggetto
      ,t0.[Utente], concat(t9.firstname,' ', t9.lastname) as 'Nome_utente'
, concat(t10.firstname,' ', t10.lastname) as 'Utente_padre'
, case when t0.assegnato is null then '' else concat(t8.firstname,' ', t8.lastname) end as 'Assegnato'
      ,t0.[Data_chiusura_totale], case when t0.aperto =1 then 'Y' else 'N' end as 'stato'

from  [TIRELLI_40].[DBO].coll_tickets t0 
  left join [TIRELLI_40].[DBO].COLL_Reparti t1 on t1.Id_Reparto=t0.Mittente
  left join [TIRELLI_40].[DBO].COLL_Reparti t2 on t2.Id_Reparto=t0.destinatario
  left join [TIRELLI_40].[DBO].coll_tickets t3 on t3.Id_Ticket= t0.id_padre
  left join [TIRELLI_40].[DBO].COLL_Reparti t4 on t4.Id_Reparto=t3.Mittente
  left join [TIRELLI_40].[DBO].COLL_motivazione t5 on t5.Id_Motivo = t0.Motivazione
LEFT JOIN oitm t6 on t6.itemcode=t0.[Commessa]
left join ocrd t7 on t7.cardcode=t6.u_final_customer_code
left join [TIRELLI_40].[dbo].ohem t8 on t8.empid=t0.assegnato
left join [TIRELLI_40].[dbo].ohem t9 on t9.empid=t0.utente
left join [TIRELLI_40].[dbo].ohem t10 on t10.empid=t3.utente
left join oscl t11 on cast(t11.callid as varchar) = CAST(substring(t0.COMMESSA,4,999) AS VARCHAR) and substring(t0.COMMESSA,1,3)='CDS'
left join oitm t12 on t12.itemcode=t11.itemcode

inner join
(select t0.[Id_Padre], max(t0.[Id_Ticket]) as 'Ticket_max' from [TIRELLI_40].[DBO].coll_tickets t0 group by t0.[Id_Padre] ) a on t0.[Id_Ticket]=a.[Ticket_max]

 where t0.Data_Creazione>=GETDATE()-2    
)
as t10

where t10.Data_Creazione>=GETDATE()-2 and t10.Id_Ticket=t10.id_padre      

  order by t10.Data_Creazione "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()





            Testo_Mail = Testo_Mail & "<tr>"

            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ID_Ticket") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Descrizione_motivo") & "</td>"


            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Commessa") & "</td>"

            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("ItemName") & "</td>"



            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Cliente") & "</td>"

            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Descrizione") & "</td>"

            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Mittente") & "</td>"

            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Destinatario") & "</td>"



            'Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Branch") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Nome_utente") & "</td>"


            Testo_Mail = Testo_Mail & "</tr>"
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

        Testo_Mail = Testo_Mail & "</table>"

        Testo_Mail = Testo_Mail & "</BODY>"


        Dim mySmtp As New SmtpClient
        Dim myMail As New MailMessage()
        mySmtp.UseDefaultCredentials = False
        mySmtp.Credentials = New Net.NetworkCredential("report@tirelli.net", "Ras70773")
        mySmtp.Host = "tirelli-net.mail.protection.outlook.com"
        mySmtp.Port = 25
        mySmtp.EnableSsl = True
        mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network







        myMail = New MailMessage()
        myMail.From = New MailAddress("Report@tirelli.net")

        myMail.To.Add(mail_destinatario)



        myMail.Subject = "Report automatico Ticket generati ieri"
        myMail.IsBodyHtml = True
        myMail.Body = Testo_Mail


        mySmtp.Send(myMail)
        i = i + 1
        Testo_Mail = ""



    End Sub

    Sub invia_mail_OC_garanzia(id As Integer)
        Dim i = 0
        Dim Testo_Mail As String
        ' Testo_Mail = "<BODY><H3></h3><P>"
        '  Testo_Mail = Testo_Mail & ""

        ' Testo_Mail = Testo_Mail & "<BR><BR>"
        Testo_Mail = Testo_Mail & "Riepilogo Ordini garanzia/recall/comp generati ieri.<BR>"



        Testo_Mail = Testo_Mail & "</P></BODY>"

        Testo_Mail = Testo_Mail & "<table border=5>"
        Testo_Mail = Testo_Mail & "<tr>"
        Testo_Mail = Testo_Mail & "<th>N° doc</th>"
        Testo_Mail = Testo_Mail & "<th>Data</th>"
        Testo_Mail = Testo_Mail & "<th>Compilatore</th>"
        Testo_Mail = Testo_Mail & "<th>Cliente</th>"
        Testo_Mail = Testo_Mail & "<th>Cliente finale</th>"

        Testo_Mail = Testo_Mail & "<th>Causale</th>"

        Testo_Mail = Testo_Mail & "<th>Osservazioni</th>"
        Testo_Mail = Testo_Mail & "<th>Codice</th>"
        Testo_Mail = Testo_Mail & "<th>Descrizione</th>"
        Testo_Mail = Testo_Mail & "<th>Quantità</th>"
        Testo_Mail = Testo_Mail & "<th>importo</th>"



        Testo_Mail = Testo_Mail & "</tr>"

        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.[DocNum], T0.[DocDate],concat(t3.lastname,' ', t3.firstname) as 'Compilatore',  T0.[CardName], case when T1.[CardName] is null then '' else t1.cardname end as 'Cliente_finale',t0.u_causcons, t0.comments, T2.[ItemCode], T2.[Dscription], T2.[Quantity], T2.[LineTotal] FROM ORDR T0  left JOIN OCRD T1 ON T0.[u_codicebp] = T1.[CardCode] 
INNER JOIN RDR1 T2 ON T0.[DocEntry] = T2.[DocEntry] 
left join [TIRELLI_40].[dbo].ohem t3 on t3.code=t0.ownercode

WHERE T0.[DocDate] >getdate()-2 and t0.u_causcons <>'V'
order by t0.docnum "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()





            Testo_Mail = Testo_Mail & "<tr>"

            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("docnum") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Docdate") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Compilatore") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Cardname") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Cliente_finale") & "</td>"

            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("u_causcons") & "</td>"



            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Comments") & "</td>"

            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Itemcode") & "</td>"
            Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Dscription") & "</td>"

            Testo_Mail = Testo_Mail & "<td>" & FormatNumber(cmd_SAP_reader("Quantity"), 2, , , TriState.True) & "</td>"

            Testo_Mail = Testo_Mail & "<td>" & FormatNumber(cmd_SAP_reader("Linetotal"), 2, , , TriState.True) & "</td>"



            'Testo_Mail = Testo_Mail & "<td>" & cmd_SAP_reader("Branch") & "</td>"



            Testo_Mail = Testo_Mail & "</tr>"
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

        Testo_Mail = Testo_Mail & "</table>"

        Testo_Mail = Testo_Mail & "</BODY>"


        Dim mySmtp As New SmtpClient
        Dim myMail As New MailMessage()
        mySmtp.UseDefaultCredentials = False
        mySmtp.Credentials = New Net.NetworkCredential("report@tirelli.net", "Ras70773")
        mySmtp.Host = "tirelli-net.mail.protection.outlook.com"
        mySmtp.Port = 25
        mySmtp.EnableSsl = True
        mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network


        myMail = New MailMessage()
        myMail.From = New MailAddress("Report@tirelli.net")

        myMail.To.Add(mail_destinatario)



        myMail.Subject = "Riepilogo Ordini garanzia/recall/comp generati ieri"
        myMail.IsBodyHtml = True
        myMail.Body = Testo_Mail


        mySmtp.Send(myMail)
        i = i + 1
        Testo_Mail = ""



    End Sub




    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        mail_destinatario = "giovannitirelli@tirelli.net"
        invia_mail_tickets(1)
        invia_mail_OC_garanzia(1)
    End Sub

End Class