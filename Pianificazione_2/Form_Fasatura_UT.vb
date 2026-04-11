Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports ADGV
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Runtime.InteropServices.ComTypes

Public Class Form_Fasatura_UT
    Private filtro_docnum As String
    Private filtro_padre As String
    Private filtro_nome_padre As String
    Private filtro_disegno As String
    Private filtro_commessa As String
    Private filtro_figlio_odp As String
    Private filtro_figlio_db As String
    Private filtro_disegno_figlio As String
    Private filtro_delta As String
    Private dataTable As New DataTable()
    Private isShiftKeyDown As Boolean = False
    Private startIndex As Integer = -1

    Public Elenco_odp(1000) As Integer
    Public Elenco_odp_eseguiti(1000) As Integer
    Public contatore_odp_eseguiti As Integer = 0

    Sub inizializza_form()
        datagridview_fasatura()
    End Sub


    Sub datagridview_fasatura()

        DataGridView1.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "
select top " & TextBox1.Text & " *
from
(
select t20.docnum,t20.plannedqty,t20.postdate, t20.Padre,t21.itemname,t21.u_disegno,t20.status,t20.u_prg_azs_commessa, t20.u_utilizz, t20.Figlio,t22.itemname as 'Nome_figlio', t22.u_disegno as 'Disegno',coalesce(t22.U_PRG_TIR_Explosion,'Y') as 'Fantasma', t20.Q,t20.D, t20.Q-t20.D as 'Delta',t20.T, t20.dt
from
(
select t10.docnum,t10.plannedqty,t10.postdate, t10.Padre,t10.status,t10.u_prg_azs_commessa, t10.u_utilizz, t10.Figlio, sum(case when t10.Q is null then 0 else t10.q end) as 'Q',sum(case when t10.D is null then 0 else t10.D end) as 'D', sum(case when t10.T is null then 0 else t10.t end) as 'T', sum(t10.dt) as 'DT'
from
(
select t0.docnum,t0.plannedqty, t0.postdate, t0.itemcode as 'Padre',t0.status,t0.u_prg_azs_commessa, t0.u_utilizz, t1.itemcode as 'Figlio', sum(t1.baseqty) as'Q',0 as 'D', sum(t1.u_prg_wip_qtaspedita) as 'T', sum(coalesce(t1.U_PRG_WIP_QtaDaTrasf,0)) as 'DT'
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
inner join oitm t2 on t2.itemcode=t0.itemcode
inner join oitm t3 on t3.itemcode=t1.itemcode

where (t0.status='P' or t0.status='R') 
and (substring(t1.itemcode,1,1)='0' or substring(t1.itemcode,1,1)='C' or substring(t1.itemcode,1,1)='D' or substring(t1.itemcode,1,1)='F' or substring(t1.itemcode,1,1)='M')

 
 " & filtro_docnum & filtro_padre & filtro_nome_padre & filtro_commessa & filtro_disegno & filtro_disegno_figlio & filtro_figlio_odp & "

group by t0.docnum,t0.plannedqty,t0.postdate, t0.itemcode,t2.itemname,t2.u_disegno,t0.status,t0.u_prg_azs_commessa, t0.u_utilizz, t1.itemcode,t3.itemname, t3.u_disegno

union all

select t0.docnum,t0.plannedqty,t0.postdate, t0.itemcode as 'Padre',t0.status,t0.u_prg_azs_commessa, t0.u_utilizz, t1.code as 'Figlio', 0 as'Q',sum(t1.[Quantity]/t2.qauntity)  as 'D', 0 as 'T', 0 as 'DT'
from owor t0 inner join itt1 t1 on t0.itemcode=t1.father
inner join oitt t2 on t2.code=t0.itemcode
where (t0.status='P' or t0.status='R') and (substring(t1.code,1,1)='0' or substring(t1.code,1,1)='C' or substring(t1.code,1,1)='D' or substring(t1.code,1,1)='F' or substring(t1.code,1,1)='M')
" & filtro_docnum & filtro_padre & filtro_nome_padre & filtro_commessa & filtro_disegno & filtro_disegno_figlio & filtro_figlio_db & "
group by t0.itemcode,t0.plannedqty, t1.code, t0.docnum,t0.postdate,t0.status,t0.u_prg_azs_commessa,t0.u_utilizz
)
as t10 
group by t10.docnum,t10.plannedqty,t10.postdate, t10.Padre,t10.status,t10.u_prg_azs_commessa, t10.u_utilizz, t10.Figlio
)
as t20 inner join oitm t21 on t21.itemcode=t20.padre
inner join oitm t22 on t22.itemcode=t20.figlio
)
as t30 where 0 = 0 " & filtro_delta & "



"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()



            DataGridView1.Rows.Add(False, cmd_SAP_reader("Docnum"), cmd_SAP_reader("postdate"), cmd_SAP_reader("padre"), cmd_SAP_reader("itemname"), cmd_SAP_reader("u_disegno"), cmd_SAP_reader("plannedqty"), cmd_SAP_reader("status"), cmd_SAP_reader("u_prg_azs_commessa"), cmd_SAP_reader("u_utilizz"), cmd_SAP_reader("figlio"), cmd_SAP_reader("nome_figlio"), cmd_SAP_reader("disegno"), cmd_SAP_reader("Fantasma"), cmd_SAP_reader("Q"), cmd_SAP_reader("D"), cmd_SAP_reader("Delta"), cmd_SAP_reader("T"), cmd_SAP_reader("DT"))


        Loop
        cmd_SAP_reader.Close()
        cnn.Close()


    End Sub

    Sub advanceddatagridview_fasatura()

        DataGridView1.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "
select top " & TextBox1.Text & " *
from
(
select t20.docnum,t20.postdate, t20.Padre,t21.itemname,t21.u_disegno,t20.status,t20.u_prg_azs_commessa, t20.u_utilizz, t20.Figlio,t22.itemname as 'Nome_figlio', t22.u_disegno as 'Disegno', t20.Q,t20.D, t20.Q-t20.D as 'Delta',t20.T
from
(
select t10.docnum,t10.postdate, t10.Padre,t10.status,t10.u_prg_azs_commessa, t10.u_utilizz, t10.Figlio, sum(case when t10.Q is null then 0 else t10.q end) as 'Q',sum(case when t10.D is null then 0 else t10.D end) as 'D', sum(case when t10.T is null then 0 else t10.t end) as 'T'
from
(
select t0.docnum,t0.postdate, t0.itemcode as 'Padre',t0.status,t0.u_prg_azs_commessa, t0.u_utilizz, t1.itemcode as 'Figlio', sum(t1.baseqty) as'Q',0 as 'D', sum(t1.u_prg_wip_qtaspedita) as 'T'
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
inner join oitm t2 on t2.itemcode=t0.itemcode
inner join oitm t3 on t3.itemcode=t1.itemcode

where (t0.status='P' or t0.status='R') and (substring(t1.itemcode,1,1)='O' or substring(t1.itemcode,1,1)='C' or substring(t1.itemcode,1,1)='D' or substring(t1.itemcode,1,1)='F')

 " & filtro_docnum & filtro_padre & filtro_nome_padre & filtro_commessa & filtro_disegno & filtro_disegno_figlio & filtro_figlio_odp & "

group by t0.docnum,t0.postdate, t0.itemcode,t2.itemname,t2.u_disegno,t0.status,t0.u_prg_azs_commessa, t0.u_utilizz, t1.itemcode,t3.itemname, t3.u_disegno

union all

select t0.docnum,t0.postdate, t0.itemcode as 'Padre',t0.status,t0.u_prg_azs_commessa, t0.u_utilizz, t1.code as 'Figlio', 0 as'Q',sum(t1.[Quantity]/t2.qauntity)  as 'D', 0 as 'T'
from owor t0 inner join itt1 t1 on t0.itemcode=t1.father
inner join oitt t2 on t2.code=t0.itemcode
where (t0.status='P' or t0.status='R') and (substring(t1.code,1,1)='O' or substring(t1.code,1,1)='C' or substring(t1.code,1,1)='D' or substring(t1.code,1,1)='F')
" & filtro_docnum & filtro_padre & filtro_nome_padre & filtro_commessa & filtro_disegno & filtro_disegno_figlio & filtro_figlio_db & "
group by t0.itemcode, t1.code, t0.docnum,t0.postdate,t0.status,t0.u_prg_azs_commessa,t0.u_utilizz
)
as t10 
group by t10.docnum,t10.postdate, t10.Padre,t10.status,t10.u_prg_azs_commessa, t10.u_utilizz, t10.Figlio
)
as t20 inner join oitm t21 on t21.itemcode=t20.padre
inner join oitm t22 on t22.itemcode=t20.figlio
)
as t30 where 0 = 0 " & filtro_delta & "



"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()



            dataTable.Rows.Add(cmd_SAP_reader("Docnum"), cmd_SAP_reader("postdate"), cmd_SAP_reader("padre"), cmd_SAP_reader("itemname"), cmd_SAP_reader("u_disegno"), cmd_SAP_reader("status"), cmd_SAP_reader("u_prg_azs_commessa"), cmd_SAP_reader("u_utilizz"), False, cmd_SAP_reader("figlio"), cmd_SAP_reader("nome_figlio"), cmd_SAP_reader("disegno"), cmd_SAP_reader("Q"), cmd_SAP_reader("D"), cmd_SAP_reader("Delta"), cmd_SAP_reader("T"))


        Loop
        cmd_SAP_reader.Close()
        cnn.Close()


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Form_Fasatura_UT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        ' crea_datagridview()
        datagridview_fasatura()
    End Sub

    Sub crea_datagridview()



        dataTable.Columns.Add("column")



        dataTable.Columns.Add("Postdate")
        dataTable.Columns.Add("Padre")
        dataTable.Columns.Add("Itemname")
        dataTable.Columns.Add("U_disegno")
        dataTable.Columns.Add("Status")
        dataTable.Columns.Add("U_prg_azs_commessa")
        dataTable.Columns.Add("U_utilizz")
        dataTable.Columns.Add("Seleziona")
        dataTable.Columns.Add("Figlio")
        dataTable.Columns.Add("Nome_figlio")
        dataTable.Columns.Add("Disegno")
        dataTable.Columns.Add("Q")
        dataTable.Columns.Add("D")
        dataTable.Columns.Add("Delta")
        dataTable.Columns.Add("T")

        datagridview_fasatura()



        '' Aggiunta delle righe con i dati al DataTable
        'dataTable.Rows.Add("DRA", "DRA1", "DRA2")
        'dataTable.Rows.Add("MILL", "MILL1", "MILL2")
        'dataTable.Rows.Add("TRO", "TROI", "MOMOSISSOKIs")

        ' Inizializzazione di un BindingSource e associazione al DataTable

        'BindingSource1.DataSource = dataTable


    End Sub





    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = "" Then
            filtro_docnum = ""
        Else
            filtro_docnum = " and t0.docnum   Like '%%" & TextBox2.Text & "%%' "
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Then
            filtro_padre = ""
        Else
            filtro_padre = " and t0.itemcode  Like '%%" & TextBox3.Text & "%%' "
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = "" Then
            filtro_nome_padre = ""
        Else
            filtro_nome_padre = " and t0.prodname   Like '%%" & TextBox4.Text & "%%' "
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text = "" Then
            filtro_disegno = ""
        Else
            filtro_disegno = " and t2.u_disegno  Like '%%" & TextBox5.Text & "%%' "
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If TextBox6.Text = "" Then
            filtro_commessa = ""
        Else
            filtro_commessa = " and t0.u_prg_azs_commessa   Like '%%" & TextBox6.Text & "%%' "
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.Text = "" Then
            filtro_figlio_odp = ""
            filtro_figlio_db = ""
        Else
            filtro_figlio_odp = " and t1.itemcode  Like '%%" & TextBox7.Text & "%%' "
            filtro_figlio_db = " and t1.code  Like '%%" & TextBox7.Text & "%%'"
        End If
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        If TextBox8.Text = "" Then
            filtro_disegno_figlio = ""
        Else
            filtro_disegno_figlio = " and t3.u_disegno   Like '%%" & TextBox8.Text & "%%' "
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        If RadioButton1.Checked = True Then
            filtro_delta = ""

        End If



    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged



        If RadioButton2.Checked = True Then
            filtro_delta = "and t30.delta >0"

        End If


    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged



        If RadioButton3.Checked = True Then
            filtro_delta = "and t30.delta <0"

        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        datagridview_fasatura()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For i As Integer = 0 To contatore_odp_eseguiti
            Elenco_odp_eseguiti(i) = 0
        Next
        contatore_odp_eseguiti = 0



        If Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = Nothing Or Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = 0 Then
            MsgBox("Selezionare un utente a cui sia associata licenza SAP")
            Return
        End If
        Dim contatore_odp As Integer = 0
        Dim check_selezione As Integer = 0
        For Each row As DataGridViewRow In DataGridView1.Rows ' Sostituisci DataGridView1 con il nome effettivo del tuo controllo DataGridView
            ' Verifica se la cella nella colonna "Seleziona" è True

            If Convert.ToBoolean(row.Cells("Seleziona").Value) = True Then
                check_selezione += 1

                If row.Cells("Delta").Value < 0 Then

                    If row.Cells("ODP").Value > 0 Then

                        aggiorna_riga(row.Cells("FIGLIO").Value, row.Cells("Docnum").Value, row.Cells("DB").Value * row.Cells("quantita").Value)

                    Else

                        If MessageBox.Show($"Aggiungere Il codice " & vbCrLf & row.Cells("FIGLIO").Value & " Quantità " & -row.Cells("Delta").Value & vbCrLf & " Nell'ODP " & row.Cells("Docnum").Value & " ?. Proseguire?", "Prosegui", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                            ODP_Form.inserisci_record_modifica_odp(Homepage.ID_SALVATO, row.Cells("Docnum").Value)


                            inserisci_riga_odp(row.Cells("figlio").Value, row.Cells("Docnum").Value, row.Cells("DB").Value * row.Cells("quantita").Value)



                        End If

                        Elenco_odp(contatore_odp) = row.Cells("Docnum").Value

                        contatore_odp += 1

                    End If


                    ' Chiamata alla tua sub passando i parametri necessari
                ElseIf row.Cells("Delta").Value > 0 Then

                    If row.Cells("Da_trasf").Value > 0 Then

                        If row.Cells("Da_trasf").Value = row.Cells("odp").Value And row.Cells("Delta").Value > 0 Then

                            cancella_riga_odp(row.Cells("FIGLIO").Value, row.Cells("Docnum").Value)

                        ElseIf row.Cells("Da_trasf").Value >= row.Cells("Delta").Value Then

                            aggiorna_riga(row.Cells("FIGLIO").Value, row.Cells("Docnum").Value, row.Cells("DB").Value * row.Cells("quantita").Value)

                        Else

                            aggiorna_riga(row.Cells("FIGLIO").Value, row.Cells("Docnum").Value, row.Cells("DB").Value * row.Cells("quantita").Value)
                        End If
                    Else
                        MsgBox("ODP " & row.Cells("docnum").Value & " Codice " & row.Cells("FIGLIO").Value & " Ha da trasferire = 0 , non è possibile diminuire la quantità")
                    End If

                    row.Cells("Seleziona").Value = False
                End If
            End If
            row.Cells("Seleziona").Value = False
        Next
        MsgBox("Documento aggiornato con successo")

        Dim contatore As Integer = 0

        Do While contatore <= contatore_odp

            If controllo_odp_eseguiti(Elenco_odp(contatore)) = 0 Then
                ODP_Form.AWOR(Elenco_odp(contatore), Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
                ODP_Form.AWO1(Elenco_odp(contatore))
                Elenco_odp_eseguiti(contatore_odp_eseguiti) = Elenco_odp(contatore)
                contatore_odp_eseguiti += 1

            End If

            contatore += 1

        Loop



        If check_selezione = 0 Then
            MsgBox("Non è stata selezionata nessuna riga")
        Else
            MsgBox("Fasatura terminata")
        End If


    End Sub
    Public Function controllo_odp_eseguiti(par_docnum As Integer)
        Dim eseguito As Integer = 0
        Dim contatore As Integer = 0

        Do While contatore <= contatore_odp_eseguiti
            If par_docnum = Elenco_odp_eseguiti(contatore) Then
                eseguito = eseguito + 1

            End If
            contatore += 1
        Loop
        Return eseguito

    End Function


    Sub inserisci_riga_odp(par_itemcode As String, par_docnum As Integer, par_quantità As String)

        ' itemcode_riga = DataGridView_ODP.Rows(contatore).Cells(columnName:="Codice").Value
        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader
        Cmd_Ticket.Connection = Cnn_Ticket

        Cmd_Ticket.CommandText = "SELECT T1.VALIDFOR AS 'Valido'
, t1.itemcode as 'Codice' 
, coalesce(t0.u_prg_azs_commessa,'') as 'Commessa'
, t0.warehouse as 'Mag_dest_odp'
,max(coalesce(t2.linenum,0)) as 'Linenum'
FROM OITM T1 ,
owor t0 
left join wor1 t2 on t2.docentry=t0.docentry
WHERE T1.[itemcode]= '" & par_itemcode & "' and t0.docnum=" & par_docnum & "
group by T1.VALIDFOR,t1.itemcode,t0.u_prg_azs_commessa,t0.warehouse
"
        Reader_Ticket = Cmd_Ticket.ExecuteReader()

        If Reader_Ticket.Read() = True Then
            If Reader_Ticket("Valido") = "N" Then
                Reader_Ticket.Close()
                Cnn_Ticket.Close()
                MsgBox("Il codice " & par_itemcode & " è inattivo ")

                Return

            ElseIf Acquisti.check_non_duplicazione_impegni_con_anticipi(par_itemcode, Reader_Ticket("Commessa")) <> "OK" Then

                MsgBox(Acquisti.check_non_duplicazione_impegni_con_anticipi(par_itemcode, Reader_Ticket("Commessa")))
                Reader_Ticket.Close()
                Cnn_Ticket.Close()
                Return
            Else

                Dim magazzino_riga As String = "01"
                If Reader_Ticket("Mag_dest_odp") = "B02" Or Reader_Ticket("Mag_dest_odp") = "BCAP2" Then
                    magazzino_riga = "B01"
                End If

                ODP_Form.inserisci_riga(par_docnum, par_itemcode, par_quantità, magazzino_riga, "4", 0, 0, par_quantità, Reader_Ticket("Linenum") + 1)



                ODP_Form.ripara_confermati(par_itemcode)
            End If


        End If


        Reader_Ticket.Close()


        Cnn_Ticket.Close()



    End Sub

    Sub aggiorna_riga(par_itemcode As String, par_docnum As Integer, par_quantità As String)
        par_quantità = Replace(par_quantità, ",", ".")
        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand

        Cmd_Ticket.Connection = Cnn_Ticket

        Cmd_Ticket.CommandText = "update t1 set t1.plannedqty=" & par_quantità & ", T1.[BaseQty]=" & par_quantità & "/t0.plannedqty
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
where t0.docnum=" & par_docnum & " and t1.itemcode='" & par_itemcode & "'"



        Cmd_Ticket.ExecuteNonQuery()

        Cmd_Ticket.CommandText = "update t1 set T1.U_PRG_WIP_QTADATRASF=T1.PLANNEDQTY-COALESCE(T1.U_PRG_WIP_QTASPEDITA,0)
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
where t0.docnum=" & par_docnum & " and t1.itemcode='" & par_itemcode & "'"


        Cmd_Ticket.ExecuteNonQuery()


        ODP_Form.ripara_confermati(par_itemcode)


        Cnn_Ticket.Close()





    End Sub

    Sub cancella_riga_odp(par_itemcode As String, par_docnum As Integer)

        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = Homepage.sap_tirelli
        Cnn_Ticket.Open()

        Dim Cmd_Ticket As New SqlCommand
        Cmd_Ticket.Connection = Cnn_Ticket

        ' Query per verificare se esistono altre righe oltre a quella che si vuole cancellare
        Cmd_Ticket.CommandText = "SELECT COUNT(*) 
                              FROM wor1 
                              WHERE docentry = (SELECT docentry FROM owor WHERE docnum = @docnum)"
        Cmd_Ticket.Parameters.AddWithValue("@docnum", par_docnum)

        Dim rowCount As Integer = Convert.ToInt32(Cmd_Ticket.ExecuteScalar())

        ' Se ci sono più di una riga, procedi con la cancellazione
        If rowCount > 1 Then
            Cmd_Ticket.CommandText = "DELETE t1
                                  FROM owor t0 
                                  INNER JOIN wor1 t1 ON t0.docentry = t1.docentry
                                  WHERE t0.docnum = @docnum AND t1.itemcode = @itemcode"
            Cmd_Ticket.Parameters.Clear()
            Cmd_Ticket.Parameters.AddWithValue("@docnum", par_docnum)
            Cmd_Ticket.Parameters.AddWithValue("@itemcode", par_itemcode)

            Cmd_Ticket.ExecuteNonQuery()
        Else
            MsgBox("La riga non può essere cancellata perché è l'unica riga. Proseguire su SAP")
        End If

        ODP_Form.ripara_confermati(par_itemcode)








    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then


            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Docnum) Then




                ODP_Form.docnum_odp = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Docnum").Value
                ODP_Form.Show()
                ODP_Form.inizializza_form(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Docnum").Value)




            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(Figlio) Then

                Magazzino.Codice_SAP = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Figlio").Value

                ' Ripristina la finestra se è minimizzata
                If Magazzino.WindowState = FormWindowState.Minimized Then
                    Magazzino.WindowState = FormWindowState.Normal
                End If

                ' Porta la finestra in primo piano
                Magazzino.BringToFront()
                Magazzino.Activate()
                Magazzino.Show()

                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)


            End If







        End If


    End Sub

    Private Sub DataGridView1_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDown
        If e.Button = MouseButtons.Left AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If ModifierKeys = Keys.Shift AndAlso startIndex >= 0 Then
                'Se è premuto Shift, cambia il flag per le righe comprese tra startIndex ed e.RowIndex
                Dim endIndex As Integer = e.RowIndex
                Dim minIndex As Integer = Math.Min(startIndex, endIndex) + 1
                Dim maxIndex As Integer = Math.Max(startIndex, endIndex) - 1

                For i As Integer = minIndex To maxIndex
                    DataGridView1.Rows(i).SetValues(True)
                Next i
            Else
                '  Altrimenti, imposta startIndex alla riga corrente
                startIndex = e.RowIndex
            End If
        End If
    End Sub

    Private Sub DataGridView_ODP_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        ' Controlla se il tasto Shift è stato premuto
        isShiftKeyDown = (e.KeyCode = Keys.ShiftKey)
    End Sub

    Private Sub DataGridView_ODP_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyUp
        ' Controlla se il tasto Shift è stato rilasciato
        isShiftKeyDown = (e.KeyCode <> Keys.ShiftKey)
    End Sub
End Class