Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop


Public Class Giacenze

    Public codice_magazzino As String
    Public filtro_codice As String
    Public filtro_descrizione As String
    Public filtro_disegno As String
    Public filtro_ubicazione As String

    Public codice_SAP As String


    Sub riempi_datagridview(par_datagridview As DataGridView)
        par_datagridview.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn




        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "declare @magazzino as varchar(18)
set @magazzino ='" & codice_magazzino & "'


select top 100 t20.itemcode, t20.itemname, t20.u_disegno,t20.u_prg_tir_trattamento, t20.u_ubicazione, t20.onhand,sum(case when t21.onhand is null then 0 else t21.onhand end - case when t21.iscommited is null then 0 else t21.iscommited end + case when t21.onorder is null then 0 else t21.onorder end) as 'Disponibile' , t20.price, t20.Valore, t20.Entrata , t20.Uscita
from
(
select t10.itemcode, t10.itemname, t10.u_disegno,t10.u_prg_tir_trattamento, t10.u_ubicazione, t10.onhand, t10.price, t10.Valore, t10.Entrata , max(t12.docdate) as 'Uscita'
from
(
select t1.itemcode, t1.itemname, t1.u_disegno, case when t1.u_prg_tir_trattamento is null then '' else t1.u_prg_tir_trattamento end as 'u_prg_tir_trattamento', t1.u_ubicazione, t0.onhand, t3.price, t0.onhand* t3.price as 'Valore', max(t5.docdate) as 'Entrata'
from oitw t0 left join oitm t1 on t0.itemcode=t1.itemcode
left join oitb t2 on t2.ItmsGrpCod=t1.ItmsGrpCod
inner join itm1 t3 on t3.itemcode=t1.itemcode
left join OIVL t4 on t4.itemcode=t0.itemcode and t4.LocCode=t0.whscode and t4.InQty>0
LEFT JOIN [OILM] T5 ON T4.[MessageID] = T5.[MessageID]




where t0.onhand>0 and t0.whscode=@magazzino and t3.pricelist=2
group by t1.itemcode, t1.itemname, t1.u_disegno,t1.u_prg_tir_trattamento, t1.u_ubicazione,t0.onhand, t3.price
)
as t10  
left join OIVL t11 on t11.itemcode=t10.itemcode and t11.LocCode=@magazzino and t11.OutQty>0
LEFT JOIN [OILM] T12 ON T12.[MessageID] = T11.[MessageID]


group by t10.itemcode, t10.itemname, t10.u_disegno,t10.u_prg_tir_trattamento, t10.u_ubicazione, t10.onhand, t10.price, t10.Valore, t10.Entrata
)
as t20 left join oitw t21 on t20.itemcode=t21.itemcode
where 0=0 " & filtro_codice & " " & filtro_descrizione & " " & filtro_disegno & " " & filtro_ubicazione & "
group by t20.itemcode, t20.itemname, t20.u_disegno,t20.u_prg_tir_trattamento, t20.u_ubicazione, t20.onhand, t20.price, t20.Valore, t20.Entrata , t20.Uscita
order by t20.valore DESC
"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            par_datagridview.Rows.Add(cmd_SAP_reader("itemcode"), cmd_SAP_reader("itemname"), cmd_SAP_reader("u_disegno"), cmd_SAP_reader("u_prg_tir_trattamento"), cmd_SAP_reader("u_ubicazione"), cmd_SAP_reader("onhand"), cmd_SAP_reader("Disponibile"), cmd_SAP_reader("price"), cmd_SAP_reader("valore"), cmd_SAP_reader("Entrata"), cmd_SAP_reader("Uscita"))

        Loop


        cmd_SAP_reader.Close()
        Cnn.Close()
        par_datagridview.ClearSelection()
        ' Seleziona la prima riga, se presente
        If par_datagridview.Rows.Count > 0 Then
            par_datagridview.Rows(0).Selected = True
        End If
    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs)
        Magazzino.ordinato(codice_SAP, DataGridView_ordinato)
    End Sub

    Sub N_articoli()
        DataGridView1.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn




        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "declare @magazzino as varchar(18)
set @magazzino ='" & codice_magazzino & "'

select count(t30.itemcode) as 'N'
from
(
select t20.itemcode, t20.itemname, t20.u_disegno, t20.u_ubicazione, t20.onhand,sum(case when t21.onhand is null then 0 else t21.onhand end - case when t21.iscommited is null then 0 else t21.iscommited end + case when t21.onorder is null then 0 else t21.onorder end) as 'Disponibile' , t20.price, t20.Valore, t20.Entrata , t20.Uscita
from
(
select t10.itemcode, t10.itemname, t10.u_disegno, t10.u_ubicazione, t10.onhand, t10.price, t10.Valore, t10.Entrata , max(t12.taxdate) as 'Uscita'
from
(
select t1.itemcode, t1.itemname, t1.u_disegno, t1.u_ubicazione, t0.onhand, t3.price, t0.onhand* t3.price as 'Valore', max(t5.taxdate) as 'Entrata'
from oitw t0 left join oitm t1 on t0.itemcode=t1.itemcode
left join oitb t2 on t2.ItmsGrpCod=t1.ItmsGrpCod
inner join itm1 t3 on t3.itemcode=t1.itemcode
left join OIVL t4 on t4.itemcode=t0.itemcode and t4.LocCode=t0.whscode and t4.InQty>0
LEFT JOIN [OILM] T5 ON T4.[MessageID] = T5.[MessageID]

where t0.onhand>0 and t0.whscode=@magazzino and t3.pricelist=2
group by t1.itemcode, t1.itemname, t1.u_disegno, t1.u_ubicazione,t0.onhand, t3.price
)
as t10  
left join OIVL t11 on t11.itemcode=t10.itemcode and t11.LocCode=@magazzino and t11.OutQty>0
LEFT JOIN [OILM] T12 ON T12.[MessageID] = T11.[MessageID]

group by t10.itemcode, t10.itemname, t10.u_disegno, t10.u_ubicazione, t10.onhand, t10.price, t10.Valore, t10.Entrata
)
as t20 left join oitw t21 on t20.itemcode=t21.itemcode

where 0=0 " & filtro_codice & " " & filtro_descrizione & " " & filtro_disegno & " " & filtro_ubicazione & "
group by t20.itemcode, t20.itemname, t20.u_disegno, t20.u_ubicazione, t20.onhand, t20.price, t20.Valore, t20.Entrata , t20.Uscita
)
as t30
"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            Label1.Text = cmd_SAP_reader("N")
        Loop


        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Private Sub TableLayoutPanel2_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub












    Private Sub Button1_Click(sender As Object, e As EventArgs)
        riempi_datagridview(DataGridView1)
    End Sub















    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            codice_SAP = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice").Value
            Button7.Text = codice_SAP
            Txt_trasferibile.Text = Form_Entrate_Merci.giacenze_IN_magazzino(codice_SAP, codice_magazzino)
            Button5.Text = Magazzino.OttieniDettagliAnagrafica(codice_SAP).Disegno
            Form_Entrate_Merci.trasferito(codice_SAP, DataGridView2)
            Label3.Text = Magazzino.OttieniDettagliAnagrafica(codice_SAP).Ubicazione

            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Codice) Then

                Magazzino.Codice_SAP = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice").Value


                Magazzino.Show()

                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)


            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(Disegno) Then


                Magazzino.visualizza_disegno(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disegno").Value)


                'Else
                '    Magazzino.trasferito(codice_SAP, DataGridView_trasferito)
                '    Magazzino.Lista_registrazioni(codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, "", "", "", "", "")
                '    Magazzino.giacenze_magazzino(DataGridView_magazzino, codice_SAP)
            End If
        End If



    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)


            Try
                codice_SAP = selectedRow.Cells("Codice").Value.ToString()
            Catch ex As Exception
                Return
            End Try


            Button7.Text = codice_SAP
                Label3.Text = Magazzino.OttieniDettagliAnagrafica(codice_SAP).Ubicazione
                Txt_trasferibile.Text = Form_Entrate_Merci.giacenze_IN_magazzino(codice_SAP, codice_magazzino)
                Button5.Text = Magazzino.OttieniDettagliAnagrafica(codice_SAP).Disegno
                Form_Entrate_Merci.trasferito(codice_SAP, DataGridView2)

                Magazzino.trasferito(codice_SAP, DataGridView_trasferito)
                Magazzino.Lista_registrazioni(codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, "", "", "", "", "")
                Magazzino.giacenze_magazzino(DataGridView_magazzino, codice_SAP)

        End If
    End Sub



    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub TextBox3_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        codice_magazzino = TextBox3.Text

    End Sub

    Private Sub DataGridView_magazzino_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_magazzino.CellContentClick

    End Sub

    Private Sub DataGridView_magazzino_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_magazzino.CellFormatting
        Dim par_datagridview As DataGridView
        par_datagridview = DataGridView_magazzino

        If e.RowIndex >= 0 Then

            If par_datagridview.Rows(e.RowIndex).Cells(columnName:="A_MAGA").Value = 0 Then
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="A_MAGA").Style.ForeColor = Color.White
            End If
            If par_datagridview.Rows(e.RowIndex).Cells(columnName:="CONF_").Value = 0 Then
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="CONF_").Style.ForeColor = Color.White
            End If
            If par_datagridview.Rows(e.RowIndex).Cells(columnName:="ORD_").Value = 0 Then
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="ORD_").Style.ForeColor = Color.White
            End If

        End If


        If par_datagridview.Rows(e.RowIndex).Cells(columnName:="MAG").Value = "TOTALE" Then
            par_datagridview.Rows(e.RowIndex).DefaultCellStyle.Font = New Font(par_datagridview.Font, FontStyle.Bold)

            If par_datagridview.Rows(e.RowIndex).Cells(columnName:="DISP_").Value < 0 Then

                par_datagridview.Rows(e.RowIndex).Cells(columnName:="DISP_").Style.ForeColor = Color.OrangeRed

            End If
        End If
    End Sub

    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        Magazzino.Lista_registrazioni(codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, "", "", "", "", "")
    End Sub

    Private Sub DateTimePicker4_ValueChanged_1(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        Magazzino.Lista_registrazioni(codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, "", "", "", "", "")
    End Sub

    Private Sub TextBox4_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = Nothing Then
            filtro_ubicazione = ""
        Else
            filtro_ubicazione = " and t20.u_ubicazione    Like '%%" & TextBox4.Text & "%%'  "
        End If


    End Sub

    Private Sub TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        If TextBox1.Text = Nothing Then
            filtro_codice = ""
        Else
            filtro_codice = " and t20.itemcode    Like '%%" & TextBox1.Text & "%%'  "
        End If

    End Sub

    Private Sub TextBox10_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        If TextBox10.Text = Nothing Then
            filtro_disegno = ""
        Else
            filtro_disegno = " and t20.u_disegno    Like '%%" & TextBox10.Text & "%%'  "
        End If


    End Sub

    Private Sub TextBox2_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = Nothing Then
            filtro_descrizione = ""
        Else
            filtro_descrizione = " and t20.itemname    Like '%%" & TextBox2.Text & "%%'  "
        End If


    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        N_articoli()
        riempi_datagridview(DataGridView1)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Giacenze_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        If Homepage.Centro_di_costo = "BRB01" Then
            Button2.Text = "BSCA"
            Button9.Text = "B03"
        Else
            Button2.Text = "SCA"
            Button9.Text = "03"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim par_magazzino_destinazione As String
        If Homepage.Centro_di_costo = "BRB01" Then
            par_magazzino_destinazione = "BSCA"

        Else
            par_magazzino_destinazione = "SCA"
        End If

        Form_Entrate_Merci.trasferimento_altro_magazzino_DEF(codice_SAP, codice_magazzino, "Giacenze", 0, CheckBox1.Checked, par_magazzino_destinazione)
        N_articoli()
        riempi_datagridview(DataGridView1)


    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Magazzino.nuovo_valore_string = InputBox("Inserire nuova ubicazione")
        Magazzino.cambiare_gestione_ubicazione(codice_SAP, Magazzino.nuovo_valore_string, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
        Magazzino.OttieniDettagliAnagrafica(codice_SAP)
        Label3.Text = Magazzino.OttieniDettagliAnagrafica(codice_SAP).Ubicazione
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim par_magazzino_destinazione As String
        If Homepage.Centro_di_costo = "BRB01" Then
            par_magazzino_destinazione = "B16"

        Else
            par_magazzino_destinazione = "16"
        End If

        Form_Entrate_Merci.trasferimento_altro_magazzino_DEF(codice_SAP, codice_magazzino, "Giacenze", 0, CheckBox1.Checked, par_magazzino_destinazione)
        N_articoli()
        riempi_datagridview(DataGridView1)


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim par_magazzino_destinazione As String
        If Homepage.Centro_di_costo = "BRB01" Then
            MsgBox("Labelling divisione non dotata di magazzino CQ")

        Else
            par_magazzino_destinazione = "CQ"
            Form_Entrate_Merci.trasferimento_altro_magazzino_DEF(codice_SAP, codice_magazzino, "Giacenze", 0, CheckBox1.Checked, par_magazzino_destinazione)
            N_articoli()
            riempi_datagridview(DataGridView1)
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim par_magazzino_destinazione As String
        If Homepage.Centro_di_costo = "BRB01" Then
            par_magazzino_destinazione = "B03"

        Else
            par_magazzino_destinazione = "03"
        End If

        Form_Entrate_Merci.trasferimento_altro_magazzino_DEF(codice_SAP, codice_magazzino, "Giacenze", 0, CheckBox1.Checked, par_magazzino_destinazione)
        N_articoli()
        riempi_datagridview(DataGridView1)


    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim par_magazzino_destinazione As String
        If Homepage.Centro_di_costo = "BRB01" Then
            MsgBox("Labelling division non dotata di magazzino refilling")

        Else
            par_magazzino_destinazione = "15"

            Form_Entrate_Merci.trasferimento_altro_magazzino_DEF(codice_SAP, codice_magazzino, "Giacenze", 0, CheckBox1.Checked, par_magazzino_destinazione)
            N_articoli()
            riempi_datagridview(DataGridView1)
        End If
    End Sub





    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        Dim divValue As String = DataGridView2.Rows(e.RowIndex).Cells("DIV").Value.ToString()
        Select Case divValue
            Case "BRB01"
                DataGridView2.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.Yellow
            Case "TIR01"
                DataGridView2.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.LightBlue
            Case "KTF01"
                DataGridView2.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.Green
        End Select
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        N_articoli()
        riempi_datagridview(DataGridView1)
    End Sub



    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick

        Dim linenum As Integer = 0
        Dim commessa_destinazione As String = ""
        Dim par_datagridview As DataGridView = DataGridView2
        Dim par_documento As String
        Dim riga As Integer


        If e.RowIndex >= 0 Then

            riga = e.RowIndex
            linenum = par_datagridview.Rows(e.RowIndex).Cells(columnName:="Linenum_").Value
            par_documento = par_datagridview.Rows(e.RowIndex).Cells(columnName:="DOC_").Value

            If e.ColumnIndex = par_datagridview.Columns.IndexOf(Trasferisci) Then



                Form_Entrate_Merci.trasferimento_wip(par_documento, linenum, codice_SAP, par_datagridview.Rows(e.RowIndex).Cells(columnName:="REP_").Value, Homepage.Centro_di_costo, codice_magazzino, par_datagridview.Rows(e.RowIndex).Cells(columnName:="ODP_").Value, par_datagridview.Rows(e.RowIndex).Cells(columnName:="OC_").Value, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Commessa_").Value, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Cliente_").Value, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Trasferisci").Value, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Stato_").Value, 0, "Giacenze", par_datagridview, riga, CheckBox1.Checked)
                N_articoli()
                riempi_datagridview(DataGridView1)





            End If
        End If

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub
End Class