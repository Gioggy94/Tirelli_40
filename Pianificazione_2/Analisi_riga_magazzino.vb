Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib

Public Class Analisi_riga_magazzino

    Public riga_1 As Integer
    Public CODICE_confermato As String
    Public condizione_commessa As String
    Public Condizione_trasferibile As String

    Sub Materiale_mancante()
        Dim Cnn1 As New SqlConnection
        DataGridView_analisi_riga.Rows.Clear()
        cnn1.ConnectionString = Homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "
    
    Select T30.DOCNUM,T30.[Status], T30.RESNAME, T30.ITEMCODE, T30.PRODNAME,T30.[U_PRG_AZS_Commessa], T30.LINENUM, T30.WAREHOUSE,t30.articolo as 'Articolo', t30.[Desc articolo] as 'Desc articolo', t30.Disegno as 'Disegno', T30.[ItmsGrpNam], t30.Quantita as 'Quantita', t30.Trasferito as 'Trasferito', t30.[Da trasferire] as 'Da trasferire' ,t30.azione as 'Azione' , t31.docnum as 'ODP', t31.[U_PRG_AZS_Commessa] as 'Commessa', t31.U_produzione as 'Reparto', t31.duedate as 'Cons ODP', t30.OA as'OA',t30.Fornitore as 'Fornitore', t30.[Cons OA] as 'Cons OA', T30.POSTDATE
from
(
Select T20.DOCNUM,T20.[Status],T20.[U_PRG_AZS_Commessa], T23.RESNAME, T20.ITEMCODE, T20.PRODNAME, t20.linenum, T20.[wareHouse], t20.articolo, t20.[Desc articolo], t20.Disegno, T20.[ItmsGrpNam], t20.Quantita, t20.Trasferito, t20.[Da trasferire],t20.azione, min(t20.[Cons ODP]) as 'Cons ODP', t22.docnum as 'OA' , t22.cardname as 'Fornitore', t21.shipdate as 'Cons OA', T20.POSTDATE
from
(
Select T10.DOCNUM,T10.[Status], T10.U_FASE,T10.[U_PRG_AZS_Commessa], T10.ITEMCODE, T10.PRODNAME, t10.linenum, T10.[wareHouse],t10.articolo, t10.[Desc articolo], t10.Disegno, T10.[ItmsGrpNam], t10.Quantita, t10.Trasferito, t10.[Da trasferire],t10.azione, t10.ODP, t10.[Cons ODP], t10.Commessa, t10.Reparto, min(t10.[Cons OA]) as 'Cons OA', T10.POSTDATE
from
(

Select T100.DOCNUM,T100.[Status], T100.U_FASE,T100.[U_PRG_AZS_Commessa],T100.ITEMCODE, T100.PRODNAME, t100.linenum, T100.[wareHouse], T100.Articolo, t100.[Desc articolo] , t100.Disegno, T100.[ItmsGrpNam], t100.Quantita,t100.Trasferito, t100.[Da trasferire], 

case when t100.[Da trasferire]=0 then 'OK' when (t100.giacenza>=t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 )  then 'Trasferibile/Da ordinare' when (t100.giacenza>=t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)>=0 )  then 'Trasferibile' when t100.[Da trasferire]>0 and t100.giacenza_TOT-t100.giacenza>=t100.[Da trasferire] then 'Mag_esterno'  when (t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)>=0 and t100.giacenza<t100.[Da trasferire]) then 'IN APPROV'   when sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 then 'Da ordinare' end as 'Azione', case when t100.[Da trasferire]=0 then '' else t102.docnum end as 'ODP', case when t100.[Da trasferire]>0 then T102.[DueDate]  end As 'Cons ODP' , case when t100.[Da trasferire]=0 then '' else t102.U_PRG_AZS_commessa end as 'Commessa' ,case when t100.[Da trasferire]=0 then '' else  t102.U_produzione end as 'Reparto', case when t100.[Da trasferire]=0 then '' else t107.docnum  end  as 'OA', 
case when t100.[Da trasferire]=0 then '' else t107.cardname end as 'Fornitore', case when t100.[Da trasferire]>0 then  t103.[ShipDate]   end  as 'Cons OA', T100.POSTDATE

from
(
SELECT  T0.DOCNUM,T0.[Status], T0.U_FASE, T0.ITEMCODE, T0.PRODNAME, T0.[U_PRG_AZS_Commessa],t1.linenum, T1.[wareHouse], T9.[ITEMCODE] as 'Articolo', t9.itemname as 'Desc articolo' , case when t9.u_disegno is null then '' else t9.u_disegno end as 'Disegno', T11.[ItmsGrpNam], t1.plannedqty as 'Quantita',case when t1.U_prg_wip_qtaspedita is null then 0 else t1.U_prg_wip_qtaspedita end as 'Trasferito', t1.u_prg_wip_qtadatrasf as 'Da trasferire', sum(t20.onhand) as 'giacenza', sum(t21.onhand) as 'giacenza_TOT', t1.docentry, T0.POSTDATE

from wor1 t1 inner join owor t0 on t0.docentry=t1.docentry
inner join oitm t9 on t9.itemcode=t1.itemcode
LEFT JOIN OITB T11 ON T9.[ItmsGrpCod] = T11.[ItmsGrpCod]
inner join oitw t20 on t20.itemcode=t1.itemcode
inner join oitw t21 on t21.itemcode=t1.itemcode
LEFT JOIN OWOR T10 ON T10.ITEMCODE=T1.ITEMCODE AND (T10.STATUS='P' OR T10.STATUS='R') and T10.[U_PRODUZIONE]='ASSEMBL'

   
     
 WHERE  (T0.[Status]='P' or T0.[Status]='R')
     
     
    
    and t1.itemtype=4 and (substring(T9.[ITEMCODE],1,1)='0' or substring(T9.[ITEMCODE],1,1)='C' or substring(T9.[ITEMCODE],1,1)='D') and (t20.whscode='01' or t20.whscode='03' or t20.whscode='SCA' or t20.whscode='FERRETTO' OR t20.whscode='MUT' ) AND T21.WHSCODE<>'WIP'  AND T10.DOCNUM IS NULL

group by 
T0.DOCNUM,T0.[Status],  T0.U_FASE,T0.[U_PRG_AZS_Commessa], T0.ITEMCODE, T0.PRODNAME, t1.linenum, T1.[wareHouse], T9.[ITEMCODE] , t9.itemname  , t9.u_disegno , T11.[ItmsGrpNam], t1.plannedqty, t1.U_prg_wip_qtaspedita , t1.u_prg_wip_qtadatrasf, t1.docentry, T0.POSTDATE
)
as t100 left join wor1 t101 on t101.itemcode=t100.articolo and t101.docentry=t100.docentry and t100.linenum=t101.linenum
left join owor t102 on t101.itemcode=t102.itemcode and (T102.Status ='P' or T102.Status ='R' )
left join por1 t103 on t103.itemcode=t101.itemcode and t103.opencreqty >0
LEFT OUTER JOIN ITT1 T104 on T101.itemCode = T104.Father
left join oitw t105 on t105.itemcode=t104.code and t105.[WhsCode]='01'
left join oitw t106 on t106.itemcode=t101.itemcode
left join opor t107 on t107.docentry=t103.docentry

group by
 T100.[articolo],T100.[Status], T100.U_FASE,T100.[U_PRG_AZS_Commessa], t100.trasferito, T100.[DESC articolo], t100.linenum, T100.[wareHouse], t100.quantita,  t100.disegno, T100.[ItmsGrpNam], t100.giacenza,t100.[da trasferire], t102.docnum, T102.[DueDate],t102.U_PRG_AZS_commessa,t102.U_produzione,t107.docnum,t107.cardname,t103.[ShipDate],T100.DOCNUM, T100.ITEMCODE, T100.PRODNAME, t100.giacenza_tot, T100.POSTDATE
)
as t10
group by T10.DOCNUM,T10.[Status], T10.U_FASE, T10.[U_PRG_AZS_Commessa],T10.ITEMCODE, T10.PRODNAME, t10.linenum, T10.[wareHouse] ,t10.articolo, t10.[Desc articolo], t10.[Desc articolo], t10.Disegno, T10.[ItmsGrpNam], t10.Quantita, t10.Trasferito, t10.[Da trasferire],t10.azione, t10.ODP, t10.[Cons ODP], t10.Commessa, t10.Reparto, T10.POSTDATE
) 
as t20
left join por1 t21 on t21.itemcode=t20.articolo and t21.shipdate=t20.[Cons OA] and t21.opencreqty >0
left join opor t22 on t22.docentry=t21.docentry
LEFT JOIN ORSC T23 ON T23.VISRESCODE =T20.U_FASE
group by
T20.DOCNUM,T20.[Status],T20.[U_PRG_AZS_Commessa], T23.RESNAME,T20.ITEMCODE, T20.PRODNAME,t20.linenum, T20.[wareHouse], t20.articolo, t20.[Desc articolo], t20.Disegno, T20.[ItmsGrpNam], t20.Quantita, t20.Trasferito, t20.[Da trasferire],t20.azione,   t22.docnum ,t22.cardname, t21.shipdate , T20.POSTDATE
)
as t30
left join owor t31 on t31.itemcode=t30.articolo and (T31.Status <> N'L' )  AND  (T31.Status <> N'C' ) and T31.[DueDate]=t30.[Cons ODP]

where t30.azione Like '%%" & TextBox4.Text & "%%' and T30.[U_PRG_AZS_Commessa] Like '%%" & TextBox7.Text & "%%' and T30.articolo Like '%%" & TextBox5.Text & "%%' and T30.[Desc articolo] Like '%%" & TextBox1.Text & "%%' and T30.docnum Like '%%" & TextBox2.Text & "%%' and T30.WAREHOUSE Like '%%" & TextBox3.Text & "%%' 

order by t30.POSTDATE"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()

            DataGridView_analisi_riga.Rows.Add(cmd_SAP_reader_2("LINENUM"), cmd_SAP_reader_2("Articolo"), cmd_SAP_reader_2("Desc articolo"), cmd_SAP_reader_2("Disegno"), cmd_SAP_reader_2("docnum"), cmd_SAP_reader_2("ITEMCODE"), cmd_SAP_reader_2("prodname"), cmd_SAP_reader_2("Status"), cmd_SAP_reader_2("Resname"), cmd_SAP_reader_2("U_PRG_AZS_Commessa"), cmd_SAP_reader_2("WAREHOUSE"), cmd_SAP_reader_2("Quantita"), cmd_SAP_reader_2("Trasferito"), cmd_SAP_reader_2("Da trasferire"), cmd_SAP_reader_2("Azione"), cmd_SAP_reader_2("ODP"), cmd_SAP_reader_2("Cons ODP"), cmd_SAP_reader_2("Commessa"), cmd_SAP_reader_2("Reparto"), cmd_SAP_reader_2("OA"), cmd_SAP_reader_2("Fornitore"), cmd_SAP_reader_2("Cons OA"))
        Loop
        cmd_SAP_reader_2.Close()
        cnn1.Close()

    End Sub

    Private Sub Button_commessa_Click(sender As Object, e As EventArgs) Handles Button_commessa.Click
        Homepage.mostra_dashboard()
        Mostra.Show()
        Mostra.Owner = Me
        Me.Hide()
    End Sub




    Sub filtra()
        Dim i = 0
        Dim parola1 As String
        Dim parola2 As String
        Dim parola4 As String
        Dim parola10 As String
        Dim parola14 As String
        Dim parola8 As String
        Do While i < DataGridView_analisi_riga.RowCount
            Try

                parola1 = UCase(DataGridView_analisi_riga.Rows(i).Cells(1).Value)
                parola2 = UCase(DataGridView_analisi_riga.Rows(i).Cells(2).Value)
                parola4 = UCase(DataGridView_analisi_riga.Rows(i).Cells(4).Value)
                parola10 = UCase(DataGridView_analisi_riga.Rows(i).Cells(10).Value)
                parola14 = UCase(DataGridView_analisi_riga.Rows(i).Cells(14).Value)
                parola8 = UCase(DataGridView_analisi_riga.Rows(i).Cells(8).Value)

                If parola2.Contains(UCase(TextBox1.Text)) Then
                    DataGridView_analisi_riga.Rows(i).Visible = True
                    If parola1.Contains(UCase(TextBox5.Text)) Then
                        DataGridView_analisi_riga.Rows(i).Visible = True


                        If parola4.Contains(UCase(TextBox2.Text)) Then
                            DataGridView_analisi_riga.Rows(i).Visible = True

                            If parola10.Contains(UCase(TextBox3.Text)) Then
                                DataGridView_analisi_riga.Rows(i).Visible = True


                                If parola14.Contains(UCase(TextBox4.Text)) Then
                                    DataGridView_analisi_riga.Rows(i).Visible = True


                                    If parola8.Contains(UCase(TextBox6.Text)) Then
                                        DataGridView_analisi_riga.Rows(i).Visible = True


                                    Else
                                        DataGridView_analisi_riga.Rows(i).Visible = False

                                    End If


                                Else
                                    DataGridView_analisi_riga.Rows(i).Visible = False

                                End If

                            Else
                                DataGridView_analisi_riga.Rows(i).Visible = False

                            End If

                        Else
                            DataGridView_analisi_riga.Rows(i).Visible = False

                        End If


                    Else
                        DataGridView_analisi_riga.Rows(i).Visible = False

                    End If

                Else
                    DataGridView_analisi_riga.Rows(i).Visible = False

                End If

            Catch ex As Exception
                DataGridView_analisi_riga.Rows(i).Visible = False
            End Try
            i = i + 1
        Loop
    End Sub



    Sub elimina_riga()
        Dim CNN3 As New SqlConnection
        CNN3.ConnectionString = Homepage.sap_tirelli
        cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = cnn3


        CMD_SAP_3.CommandText = "DELETE T0 FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T1.[DocNum] =" & DataGridView_analisi_riga.Rows(riga_1).Cells(4).Value & " AND  T0.[ItemCode] ='" & DataGridView_analisi_riga.Rows(riga_1).Cells(1).Value & "' AND  T0.[LineNum] =" & DataGridView_analisi_riga.Rows(riga_1).Cells(0).Value & ""

        CMD_SAP_3.ExecuteNonQuery()
        cnn3.Close()
        For Each Riga As DataGridViewRow In DataGridView_analisi_riga.SelectedRows
            DataGridView_analisi_riga.Rows.RemoveAt(riga_1)
        Next

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView_analisi_riga.Rows(riga_1).Cells(12).Value > 0 Then

            MsgBox("Trasferito >0 , rendere il pezzo prima di eliminare la riga")

        Else

            Dim Question
            Question = MsgBox("Sei sicuro di voler eliminare il codice " & DataGridView_analisi_riga.Rows(riga_1).Cells(1).Value & " nell'ODP " & DataGridView_analisi_riga.Rows(riga_1).Cells(4).Value & "?", vbYesNo)
            If Question = vbYes Then

                CODICE_confermato = DataGridView_analisi_riga.Rows(riga_1).Cells(1).Value
                elimina_riga()
                ripara_confermati()

            End If
        End If
    End Sub



    Private Sub DataGridView_analisi_riga_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_analisi_riga.CellClick
        If e.RowIndex >= 0 Then
            riga_1 = e.RowIndex
            If e.ColumnIndex = 1 Then


                Magazzino.Codice_SAP = DataGridView_analisi_riga.Rows(e.RowIndex).Cells(1).Value

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
                Me.WindowState = FormWindowState.Minimized
            End If
            If e.ColumnIndex = 3 Then

                Try
                    Process.Start(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_analisi_riga.Rows(e.RowIndex).Cells(3).Value & ".PDF")
                Catch ex As Exception
                    MsgBox("Il disegno " & DataGridView_analisi_riga.Rows(e.RowIndex).Cells(3).Value & " non è ancora stato processato")
                End Try

            End If

            If e.ColumnIndex = 4 Then






                ODP_Form.docnum_odp = DataGridView_analisi_riga.Rows(e.RowIndex).Cells(4).Value
                ODP_Form.Show()
                ODP_Form.inizializza_form(DataGridView_analisi_riga.Rows(e.RowIndex).Cells(4).Value)

            End If
        End If
    End Sub

    Sub ripara_confermati()
        Dim CNN3 As New SqlConnection
        CNN3.ConnectionString = Homepage.sap_tirelli
        cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = cnn3


        CMD_SAP_3.CommandText = "update t41 set t41.iscommited=t40.import_confermati
from
(
Select t30.itemcode, sum(t30.IMPORT_CONFERMATI) as 'Import_confermati', T30.WHSCODE
from
(
SELECT t20.itemcode,T20.CONFERMATI, t20.MAG, T21.WHSCODE, CASE WHEN T21.WHSCODE=t20.MAG THEN T20.CONFERMATI ELSE 0 END AS 'IMPORT_CONFERMATI'
FROM
(
SELECT t10.itemcode,sum(t10.confermati) AS 'CONFERMATI', t10.MAG
FROM
(
SELECT T0.[ItemCode], SUM(T0.[PlannedQty]) AS 'CONFERMATI', T0.[wareHouse] AS 'mag'
 FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE (T1.[STATUS] ='P' OR  T1.[STATUS] ='R') AND T1.[CmpltQty]< T1.[PlannedQty] 
GROUP BY T0.[ItemCode],T0.[wareHouse]

UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.[WhsCode] 
FROM RDR1 T0  INNER JOIN ORDR T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T1.[DocStatus] ='O' AND T0.[OpenQty]>0
GROUP BY 
T0.[ItemCode],  T0.[WhsCode]
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.FROMWHSCOD 
FROM WTQ1 T0  INNER JOIN OWTQ T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T0.[OpenQty] >0 AND  T1.[DocStatus] ='O'
GROUP BY 
T0.[ItemCode], T0.FROMWHSCOD
)
AS T10
group by t10.itemcode, t10.MAG
)
AS T20 LEFT JOIN OITW T21 ON T20.ITEMCODE=T21.ITEMCODE
WHERE (T21.WHSCODE='01' OR T21.WHSCODE='03' OR T21.WHSCODE='WIP' OR T21.WHSCODE='FERRETTO' OR T21.WHSCODE='MUT' OR T21.WHSCODE='SCA') 
)
as t30
group by t30.itemcode, T30.WHSCODE
)
as t40 inner join oitw t41 on t41.itemcode=t40.itemcode and t41.whscode=t40.whscode
where t40.itemcode='" & CODICE_confermato & "'"

        CMD_SAP_3.ExecuteNonQuery()
        cnn3.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Close()
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'filtra()
        Materiale_mancante()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub
End Class