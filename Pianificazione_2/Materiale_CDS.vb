Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Materiale_CDS
    Public riga_1 As Integer
    Public STAMPA_ETICHETTA As String

    Public testata_oc_docnum as string
    Public testata_oc_cardname as String
    Public testata_oc_u_categoria As String
    Public testata_oc_u_matrcds As String
    Public testata_oc_clientefinale As String
    Public testata_oc_data As String
    Public testata_oc_max_righe As Integer

    Public percorso_documento As String

    Public oWord As Word.Application
    Public oDoc As Word.Document
    Public oTable As Word.Table




    Sub righe_ordine()

        DataGridView_riga_OC.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = cnn1

        CMD_SAP_1.CommandText = " declare @oc as varchar(10)
declare @CDS as varchar(10)
set @oc=" & Commesse_MES.OC & "
set @CDS=" & Commesse_MES.CDS_ & "

SELECT *
FROM
(
select case when t60.Tipo =1 then 'Materiale OC' else concat('Materiale per ODP',t61.prodname) end as 'Fonte', t60.tipo,t60.DOC, t60.oc,t60.ODP_FONTE, t60.CDS,t60.itemcode,t60.itemname, t60.whscode, t60.u_disegno,t60.ItmsGrpNam , t60.linenum, t60.openqty, t60.U_Datrasferire,t60.Giacenza,t60.CQ, t60.Clavter, t60.Ordinato,  t60.Disp,
t60.Azione,
t60.ODP,t60.[Cons ODP],t60.U_PRODUZIONE, t60.OA,t60.Fornitore,t60.ShipDate
,
case when  t60.Azione='OK' then 1 
when  t60.Azione='Trasferibile' then 2
when  t60.Azione='CQ' then 3
when  t60.Azione='Clavter' then 4
when  t60.Azione='06' then 3.5
when  t60.Azione='16' then 4.5
when  t60.Azione='09' then 4.6
when  t60.Azione='CAP2' then 4.7
when  t60.Azione='In_approv_scaduto' then 5
when  t60.Azione='In_approv_futuro' then 6
when  t60.Azione='In_approv_scaduto/Da_Ordinare' then 7
when  t60.Azione='In_approv_futuro/Da_Ordinare' then 8
when  t60.Azione='Da_ordinare' then 9
when  t60.Azione='Assembl' then 10
when  t60.Azione='?' then 11
ELSE 999
end as 'Azione_N'
, t61.itemcode as 'Codice_ODP', t61.prodname

from
(

select t50.Tipo,t50.DOC, t50.oc,t50.OC_dell_ODP,t50.odp as 'ODP_fonte', t50.CDS,t50.itemcode,t51.itemname, t50.whscode, t51.u_disegno,t52.ItmsGrpNam , t50.linenum, t50.openqty, t50.U_Datrasferire,t50.Giacenza,t50.CQ,T50.[06],T50.[16],T50.[CAP2],T50.[09], t50.Clavter, t50.Ordinato,  t50.Disp,

case when t50.azione='In_approv' and a.u_produzione='ASSEMBL' then 'Assembl'

when t50.azione='In_approv' and case when a.DueDate is null then b.ShipDate when b.ShipDate is null then a.DueDate when  a.DueDate <=b.ShipDate then a.DueDate else b.ShipDate end <= getdate() then 'In_approv_scaduto'
when t50.azione='In_approv' and case when a.DueDate is null then b.ShipDate when b.ShipDate is null then a.DueDate when  a.DueDate <=b.ShipDate then a.DueDate else b.ShipDate end > getdate() then 'In_approv_futuro'
when t50.azione='In_approv/Da_Ordinare' and case when a.DueDate is null then b.ShipDate when b.ShipDate is null then a.DueDate when  a.DueDate <=b.ShipDate then a.DueDate else b.ShipDate end <= getdate() then 'In_approv_scaduto/Da_Ordinare'
when t50.azione='In_approv/Da_Ordinare' and case when a.DueDate is null then b.ShipDate when b.ShipDate is null then a.DueDate when  a.DueDate <=b.ShipDate then a.DueDate else b.ShipDate end > getdate() then 'In_approv_futuro/Da_Ordinare'
else t50.azione end as 'Azione',

A.Docnum as 'ODP',a.DueDate as 'Cons ODP',a.U_PRODUZIONE, b.Docnum as 'OA',b.CardName as 'Fornitore',b.ShipDate


from
(
select t40.Tipo,t40.DOC, t40.oc,t40.OC_dell_ODP,t40.odp, t40.CDS,t40.itemcode,t40.linenum, t40.openqty, t40.U_Datrasferire,t40.Giacenza,t40.CQ, t40.Clavter,T40.[06],T40.[16],T40.[CAP2],T40.[09], t40.Ordinato,  t40.Disp,
case when t40.U_Datrasferire =0 or t40.U_Datrasferire is null then 'OK'
when t40.Giacenza>=t40.U_Datrasferire then 'Trasferibile'
when t40.Giacenza+t40.cq>=t40.U_Datrasferire then 'CQ'
when t40.Giacenza+t40.clavter>=t40.U_Datrasferire then 'Clavter'
when t40.Giacenza+t40.[06]>=t40.U_Datrasferire then '06'
when t40.Giacenza+t40.[16]>=t40.U_Datrasferire then '16'
when t40.Giacenza+t40.[09]>=t40.U_Datrasferire then '09'
when t40.Giacenza+t40.[CAP2]>=t40.U_Datrasferire then 'CAP2'
when t40.Giacenza+t40.ordinato>= t40.U_Datrasferire and t40.Disp>= 0 then 'In_approv'
when t40.Giacenza+t40.ordinato>= t40.U_Datrasferire and t40.Disp< 0 then 'In_approv/Da_Ordinare'
when t40.Giacenza+t40.ordinato<=t40.U_Datrasferire and t40.Disp< 0 then 'Da_ordinare'
else '?'
end as 'Azione', t40.whscode
from
(
select t30.Tipo,t30.DOC, t30.oc,t30.OC_dell_ODP,t30.odp, t30.CDS,t30.itemcode,t30.linenum, t30.openqty, t30.U_Datrasferire,t30.Giacenza,case when t32.onhand is null then 0 else t32.onhand end as 'CQ', case when t33.onhand is null then 0 else t33.onhand end as'Clavter', case when t34.onhand is null then 0 else t34.onhand end as'06',
case when t36.onhand is null then 0 else t36.onhand end as'09',
case when t37.onhand is null then 0 else t37.onhand end as 'CAP2',
case when t35.onhand is null then 0 else t35.onhand end as'16', sum(case when t31.onorder is null then 0 else t31.onorder end) as 'Ordinato',  t30.Disp, t30.whscode

from
(
select 
t20.Tipo,t20.DOC, t20.oc,t20.OC_dell_ODP,t20.odp, t20.CDS,t20.itemcode,t20.linenum, t20.openqty, t20.U_Datrasferire,t20.Giacenza, sum(case when t21.onhand is null then 0 else t21.onhand end -case when t21.iscommited is null then 0 else t21.iscommited end + case when t21.onorder is null then 0 else t21.onorder end) as 'Disp', t20.whscode

from
(
select 
t10.Tipo,t10.DOC, t10.OC,t10.OC_dell_ODP,t10.odp, t10.CDS,t10.itemcode,t10.linenum, t10.openqty, t10.U_Datrasferire,sum(t11.onhand) as 'Giacenza', t10.whscode

from
(

select 1 as 'Tipo','OC' as 'DOC', t0.docnum as 'OC','' as 'ODP','' as 'OC_dell_ODP', cast(t3.callid as varchar) as 'CDS',t1.itemcode,t1.linenum, t1.openqty, t1.U_Datrasferire, t1.whscode

from ordr t0 inner join rdr1 t1 on t0.docentry=t1.docentry
left join oclg t2 on t2.DocNum=t0.docnum and t2.DocType =17 AND T2.[parentId] >1
left join oscl t3 on t3.callid=T2.[parentId]

WHERE T0.DocStatus = N'o'  and (substring(t1.itemcode,1,1)='0' or substring(t1.itemcode,1,1)='C' or substring(t1.itemcode,1,1)='D') AND T0.DOCNUM=@OC

union all

select 2,'ODP',T0.ORIGINNUM as 'OC', t0.docnum as 'ODP','' as 'OC_dell_ODP',cast(substring(t0.u_prg_azs_commessa,4,99) as varchar), t1.itemcode, t1.linenum,t1.PlannedQty, t1.U_PRG_WIP_QtaDaTrasf, t1.wareHouse
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
where (t0.Status= 'p' or t0.Status= 'r') and  T0.ORIGINNUM=@OC


)
as t10 left join oitw t11 on t10.itemcode=t11.itemcode and t11.whscode<>'WIP' and t11.whscode<>'CQ' and t11.whscode<>'Clavter' and t11.whscode<>'06' and t11.whscode<>'16'and t11.whscode<>'09'and t11.whscode<>'CAP2'
group by t10.Tipo,t10.DOC, t10.oc,t10.OC_dell_ODP, t10.CDS,t10.itemcode,t10.linenum, t10.openqty, t10.U_Datrasferire,t10.odp, t10.whscode
)
as t20 left join oitw t21 on t20.itemcode=t21.itemcode

group by t20.Tipo,t20.DOC, t20.oc,t20.OC_dell_ODP, t20.CDS,t20.itemcode,t20.linenum, t20.openqty, t20.U_Datrasferire,t20.Giacenza,t20.odp, t20.whscode
)
as t30 left join oitw t31 on t31.itemcode=t30.itemcode
left join oitw t32 on t32.itemcode=t30.itemcode and t32.whscode='CQ'
left join oitw t33 on t33.itemcode=t30.itemcode and t33.whscode='Clavter'
left join oitw t34 on t34.itemcode=t30.itemcode and t34.whscode='06'
left join oitw t35 on t35.itemcode=t30.itemcode and t35.whscode='16'
left join oitw t36 on t36.itemcode=t30.itemcode and t36.whscode='09'
left join oitw t37 on t37.itemcode=t30.itemcode and t37.whscode='CAP2'
group by t30.Tipo,t30.DOC, t30.oc,t30.OC_dell_ODP, t30.CDS,t30.itemcode,t30.linenum, t30.openqty, t30.U_Datrasferire,t30.Giacenza,  t30.Disp, t32.onhand, t33.onhand,t30.odp, t30.whscode, T34.ONHAND, T35.ONHAND,T36.ONHAND,T37.ONHAND
)
as t40
)
as t50

left join 
(
select t10.itemcode, min(t11.docnum) as 'Docnum', t11.duedate, t11.U_PRODUZIONE
from 
(
select t0.itemcode,  min(t0.DueDate) as 'Min_data_odp'
from owor t0
where (t0.status='P' or t0.status='R')
group by t0.itemcode
)
as t10 left join owor t11 on t11.itemcode=t10.itemcode and (t11.status='P' or t11.status='R') and t10.Min_data_odp=t11.duedate
group by t11.duedate,t10.itemcode, t11.U_PRODUZIONE
)
A on a.itemcode=t50.itemcode and t50.azione   Like '%%IN_APPROV%%' 

left join 
(
select t10.itemcode, min(t12.docnum) as 'Docnum', t11.shipdate, t12.cardname
from 
(
select t0.itemcode,  min(t0.shipDate) as 'Min_data_Oa'
from por1 t0
where (t0.OpenQty>0)
group by t0.itemcode
)
as t10 left join por1 t11 on t11.itemcode=t10.itemcode and t11.OpenQty>0 and t10.Min_data_oa=t11.shipdate
left join opor t12 on t12.docentry=t11.docentry
group by t11.shipdate,t10.itemcode,t12.cardname
)
B on b.itemcode=t50.itemcode and t50.azione   Like '%%IN_APPROV%%'
LEFT JOIN OITM t51 on t51.itemcode=t50.itemcode
left join oitb t52 on t52.ItmsGrpCod=t51.ItmsGrpCod


)
as t60 LEFT JOIN OWOR T61 ON T61.DOCNUM=T60.ODP_FONTE AND T60.ODP_FONTE<>0
)
AS T70
order by t70.tipo, t70.azione_n DESC,t70.odp_fonte   
  "


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()

            DataGridView_riga_OC.Rows.Add(cmd_SAP_reader_1("Tipo"), cmd_SAP_reader_1("Azione_n"), cmd_SAP_reader_1("Fonte"), cmd_SAP_reader_1("itemcode"), cmd_SAP_reader_1("itemname"), cmd_SAP_reader_1("u_Disegno"), cmd_SAP_reader_1("ItmsGrpNam"), cmd_SAP_reader_1("whscode"), cmd_SAP_reader_1("Openqty"), cmd_SAP_reader_1("u_datrasferire"), cmd_SAP_reader_1("Azione"), cmd_SAP_reader_1("ODP"), cmd_SAP_reader_1("Cons ODP"), cmd_SAP_reader_1("U_produzione"), cmd_SAP_reader_1("OA"), cmd_SAP_reader_1("Fornitore"), cmd_SAP_reader_1("shipdate"), cmd_SAP_reader_1("codice_odp"), cmd_SAP_reader_1("Prodname"))

        Loop
        cnn1.Close()
        DataGridView_riga_OC.ClearSelection()
    End Sub




    Sub inserimento_reparti()

        ComboBox3.Items.Clear()
        DataGridView_riga_OC.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = cnn1

        CMD_SAP_1.CommandText = " SELECT T0.[Code] FROM [dbo].[@COMMENTO_ORDINE]  T0"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()
            ComboBox3.Items.Add(cmd_SAP_reader_1("Code"))


        Loop
        cnn1.Close()
        DataGridView_riga_OC.ClearSelection()
    End Sub


    Sub Testata_ordine()

        DataGridView_riga_OC.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = cnn1

        CMD_SAP_1.CommandText = "SELECT cast(case when t0.u_peson is null then 0 else t0.u_peson end as decimal) as 'u_peson', cast(case when t0.u_pesol is null then 0 else t0.u_pesol end as decimal) as 'u_pesol', case when t0.u_commento is null then '' else t0.u_commento end as 'U_commento' , case when t0.u_prg_azs_dimimb is null then '' else t0.u_prg_azs_dimimb end as 'u_prg_azs_dimimb'  FROM ORDR T0 WHERE T0.[DocNum] ='" & Commesse_MES.OC & "'"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        If cmd_SAP_reader_1.Read() Then

            TextBox2.Text = cmd_SAP_reader_1("u_peson")
            TextBox6.Text = cmd_SAP_reader_1("u_pesol")
            ComboBox3.Text = cmd_SAP_reader_1("u_commento")
            TextBox7.Text = cmd_SAP_reader_1("u_prg_azs_dimimb")



        End If
        cnn1.Close()
        DataGridView_riga_OC.ClearSelection()
    End Sub


    Sub aggiorna_peson()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = cnn
        Cmd_SAP.CommandText = "UPDATE Ordr SET ordr.U_peson='" & TextBox2.Text & "' where ordr.docnum = " & Commesse_MES.OC & ""
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()
    End Sub

    Sub aggiorna_pesol()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = cnn
        Cmd_SAP.CommandText = "UPDATE Ordr SET ordr.U_pesol='" & TextBox6.Text & "' where ordr.docnum = " & Commesse_MES.OC & ""
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()
    End Sub

    Sub aggiorna_reparto()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = cnn
        Cmd_SAP.CommandText = "UPDATE Ordr SET ordr.u_commento='" & ComboBox3.Text & "' where ordr.docnum = " & Commesse_MES.OC & ""
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()
    End Sub

    Sub aggiorna_DIM_IMB()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = cnn
        Cmd_SAP.CommandText = "UPDATE Ordr SET ordr.u_prg_azs_dimimb='" & TextBox7.Text & "' where ordr.docnum = " & Commesse_MES.OC & ""
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Close()
    End Sub

    Private Sub DataGridView_riga_OC_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_riga_OC.CellClick
        If e.RowIndex >= 0 Then
            riga_1 = e.RowIndex
            If e.ColumnIndex = DataGridView_riga_OC.Columns.IndexOf(Codice) Then

                Magazzino.Codice_SAP = DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Codice").Value

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
            If e.ColumnIndex = DataGridView_riga_OC.Columns.IndexOf(Disegno) Then

                Try
                    Process.Start(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Disegno").Value & ".PDF")
                Catch ex As Exception
                    MsgBox("Il disegno " & DataGridView_riga_OC.Rows(e.RowIndex).Cells(3).Value & " non è ancora stato processato")
                End Try

            End If

            If e.ColumnIndex = DataGridView_riga_OC.Columns.IndexOf(ODP) Then




                ODP_Form.docnum_odp = DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="ODP").Value
                ODP_Form.Show()
                ODP_Form.inizializza_form(DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="ODP").Value)


            End If




        End If
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        STAMPA_ETICHETTA = "no"

        testata_oc()


        If CheckBox1.Checked = True And CheckBox2.Checked = False Then
            percorso_documento = Homepage.PERCORSO_DOCUMENTO_OC
            Genera_ordine()
        ElseIf CheckBox2.Checked = True Then
            percorso_documento = Homepage.PERCORSO_DOCUMENTO_OC
            STAMPA_ETICHETTA = "YES"
            Genera_ordine()
        End If
    End Sub

    Sub testata_oc()


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "SELECT count(t1.itemcode) as 'MAx_righe',[DocNum], T0.[CardName], T0.[U_Categoria], case when T0.[U_MATRcds] is null then '' else t0.u_matrcds end as 'u_matrcds', case when T0.[U_Clientefinale] is null then '' else t0.u_clientefinale end as 'u_clientefinale', getdate() as 'Data' FROM ORDR T0 inner join rdr1 t1 on t0.docentry=t1.docentry WHERE T0.[DocNum] ='" & Commesse_MES.OC & "' group by [DocNum], T0.[CardName], T0.[U_Categoria], T0.[U_MATRcds], T0.[U_Clientefinale]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() = True Then
            testata_oc_docnum = cmd_SAP_reader_2("DOCNUM")
            testata_oc_cardname = cmd_SAP_reader_2("cardname")
            testata_oc_u_categoria = cmd_SAP_reader_2("u_categoria")

            testata_oc_u_matrcds = cmd_SAP_reader_2("u_matrcds")
            testata_oc_clientefinale = cmd_SAP_reader_2("u_clientefinale")
            testata_oc_data = cmd_SAP_reader_2("data")

            testata_oc_max_righe = cmd_SAP_reader_2("Max_righe")

        End If
        cmd_SAP_reader_2.Close()
        cnn1.Close()

    End Sub

    Sub Genera_ordine()

        testata_oc()


        oWord = CreateObject("Word.Application")

        oDoc = oWord.Documents.Add("" & percorso_documento & "")

        segnalibri_testata_oc()
        If STAMPA_ETICHETTA = "YES" Then
            segnalibri_etichetta_cassetta()
        End If


        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("Tabella").Range, testata_oc_max_righe + 1, 7)

        oTable.Cell(1, 1).Range.Text = "Codice"
        oTable.Cell(1, 2).Range.Text = "Descrizione"
        oTable.Cell(1, 3).Range.Text = "U.M."
        oTable.Cell(1, 4).Range.Text = "Q.TA"
        oTable.Cell(1, 5).Range.Text = "T"
        oTable.Cell(1, 6).Range.Text = "MAG"
        oTable.Cell(1, 7).Range.Text = "UBIC"


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        Dim i As Integer = 2

        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "select t1.itemcode,t2.itemname, case when T2.[InvntryUom] is null then '' else t2.invntryuom end as 'invntryuom' , t1.quantity, case when t1.u_trasferito is null then 0 else t1.u_trasferito end as 'u_trasferito', t1.whscode, case when t2.u_ubicazione is null then '' else t2.u_ubicazione end as 'u_ubicazione'
from ordr t0 inner join rdr1 t1 on t0.docentry=t1.docentry inner join oitm t2 on t2.itemcode=t1.itemcode where t0.docnum='" & Commesse_MES.OC & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            oTable.Cell(i, 1).Range.Text = cmd_SAP_reader_2("ItemCode")
            oTable.Cell(i, 2).Range.Text = cmd_SAP_reader_2("ItemName")
            oTable.Cell(i, 3).Range.Text = cmd_SAP_reader_2("InvntryUom")

            oTable.Cell(i, 4).Range.Text = FormatNumber(cmd_SAP_reader_2("quantity"), 1, , , TriState.True)
            If cmd_SAP_reader_2("u_trasferito") = 0 Then
                oTable.Cell(i, 5).Range.Text = ""
            Else
                oTable.Cell(i, 5).Range.Text = FormatNumber(cmd_SAP_reader_2("u_trasferito"), 1, , , TriState.True)
            End If

            oTable.Cell(i, 6).Range.Text = cmd_SAP_reader_2("whscode")
            oTable.Cell(i, 7).Range.Text = cmd_SAP_reader_2("u_ubicazione")


            oTable.Cell(i, 2).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
            oTable.Rows(i).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter



            i = i + 1
        Loop



        cnn1.Close()

        'oTable.Range.ParagraphFormat.SpaceAfter = 6

        'oTable.Cell(r, 1).Range.Text = r - 1
        'oTable.Cell(r, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
        'oTable.Cell(r, 2).Range.Text = Codice
        'ble.Cell(r, 3).Range.Text = Descrizione
        'With oTable.Cell(r, 3).Range.Font
        '.Name = "Arial"
        '.Bold = 1
        ' If Len(Descrizione) < 35 Then
        '.Size = 9
        'Else
        '.Size = 7
        'End If

        'End With '
        'oTable.Cell(r, 4).Range.Text = Note
        'oTable.Cell(r, 5).Range.Text = Quantità
        'oTable.Cell(r, 5).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
        'oTable.Cell(r, 6).Range.Text = Valuta & " " & FormatNumber(Prezzo, 2, , , TriState.True)
        'oTable.Cell(r, 6).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
        'oTable.Cell(r, 7).Range.Text = Valuta & " " & FormatNumber(Totale, 2, , , TriState.True)
        ' oTable.Cell(r, 7).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight

        'Next

        oTable.AutoFormat(ApplyColor:=False, ApplyBorders:=False)
        oTable.Rows.Item(1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Rows.Item(1).Range.Font.Bold = True
        'oTable.Columns.Item(1).Width = oWord.InchesToPoints(1)   'Change width of columns 1 & 2
        oTable.Rows(1).Cells.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        oTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        oTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        oTable.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)


        'oWord.Visible = True
        'oWord.ShowMe()

        oWord.PrintOut()

    End Sub

    Sub segnalibri_testata_oc()


        oDoc.Bookmarks.Item("ordine").Range.Text = testata_oc_docnum
        oDoc.Bookmarks.Item("cliente").Range.Text = testata_oc_cardname
        oDoc.Bookmarks.Item("causale").Range.Text = testata_oc_u_categoria
        oDoc.Bookmarks.Item("commessa").Range.Text = testata_oc_u_matrcds
        oDoc.Bookmarks.Item("utilizzatore").Range.Text = testata_oc_clientefinale
        oDoc.Bookmarks.Item("data").Range.Text = testata_oc_data


    End Sub

    Sub segnalibri_etichetta_cassetta()

        oDoc.Bookmarks.Item("ordine_ETI").Range.Text = testata_oc_docnum
        oDoc.Bookmarks.Item("cliente_ETI").Range.Text = testata_oc_cardname
        oDoc.Bookmarks.Item("causale_ETI").Range.Text = testata_oc_u_categoria
        oDoc.Bookmarks.Item("commessa_ETI").Range.Text = testata_oc_u_matrcds
        oDoc.Bookmarks.Item("utilizzatore_ETI").Range.Text = testata_oc_clientefinale
        oDoc.Bookmarks.Item("data_ETI").Range.Text = testata_oc_data


    End Sub

    Private Sub DataGridView_CDS_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        aggiorna_peson()
        MsgBox("Campo aggiornato con successo")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        aggiorna_pesol()
        MsgBox("Campo aggiornato con successo")
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        aggiorna_reparto()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        aggiorna_DIM_IMB()
        MsgBox("Campo aggiornato con successo")
    End Sub

    Private Sub DataGridView_riga_oc_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_riga_OC.CellFormatting



        If DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Azione_n").Value = "2" Then
            DataGridView_riga_OC.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime

        ElseIf DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Azione_n").Value = "3" Then
            DataGridView_riga_OC.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Crimson
        ElseIf DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Azione_n").Value = "4" Then

            DataGridView_riga_OC.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Aquamarine
        ElseIf DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Azione_n").Value = "4,6" Then

            DataGridView_riga_OC.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightYellow

        ElseIf DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Azione_n").Value = "4,7" Then

            DataGridView_riga_OC.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightYellow
        ElseIf DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Azione_n").Value = "5" Then
            DataGridView_riga_OC.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightYellow
        ElseIf DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Azione_n").Value = "6" Then
            DataGridView_riga_OC.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Yellow
        ElseIf DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Azione_n").Value = "7" Then
            DataGridView_riga_OC.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Orange
        ElseIf DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Azione_n").Value = "8" Then
            DataGridView_riga_OC.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.OrangeRed
        ElseIf DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Azione_n").Value = "9" Then
            DataGridView_riga_OC.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Red
        ElseIf DataGridView_riga_OC.Rows(e.RowIndex).Cells(columnName:="Azione_n").Value = "10" Then
            DataGridView_riga_OC.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Chocolate

        End If



    End Sub

    Private Sub DataGridView_riga_OC_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_riga_OC.CellContentClick

    End Sub
End Class