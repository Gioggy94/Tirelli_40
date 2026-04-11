Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Materiale_MU
    Public RIGA As Integer
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Hide()
    End Sub

    Sub materiale()


        DataGridView_MATERIALE.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "Select T30.docentry, T30.[N° Ordine], T30.DOC, T30.COD, T30.NOME, T30.Q, T30.DIS, t30.produzione, t30.itemcode,t30.itemname,t30.u_ubicazione, T30.TRAS, T30.[DA TRAS], T30.A_mag, T30.Stato,  T30.LAV, T30.PRIO,T30.[DueDate], t30.Prog
from
(
Select T20.docentry, T20.[N° Ordine], T20.DOC, T20.COD, T20.NOME, T20.Q, T20.DIS, t20.produzione, t20.itemcode,t20.itemname,t20.u_ubicazione, T20.TRAS, T20.[DA TRAS], T20.A_mag, T20.Stato,  T20.LAV, T20.PRIO,T20.[DueDate], SUM(CASE WHEN T21.ITEMCODE='R00554' and t21.u_stato_lavorazione='O' THEN 1 ELSE 0 END) AS 'Prog'
from
(
SELECT T10.docentry, T10.[N° Ordine], T10.DOC, T10.COD, T10.NOME, T10.Q, T10.DIS, t10.produzione, t10.itemcode,t10.itemname,t10.u_ubicazione, T10.TRAS, T10.[DA TRAS], T10.A_mag, CASE WHEN T10.A_mag>=T10.[DA TRAS] THEN 'Trasferibile' when sum (CASE WHEN t11.ONORDER IS NULL THEN 0 ELSE T11.ONORDER END)+T10.A_mag >= T10.[DA TRAS] then 'IN APPROV' ELSE 'DA ORDINARE' end as 'Stato',  T10.LAV, T10.PRIO,T10.[DueDate]
FROM
(
SELECT T0.docentry, T0.[DocNum] as 'N° Ordine', 'ODP' as 'DOC', T0.[ItemCode] AS 'COD', T2.[ItemNAME] AS 'NOME', T0.[PlannedQty] AS 'Q', t2.u_disegno AS 'DIS', t3.itemcode,t3.itemname, t3.u_ubicazione, T1.[U_PRG_WIP_QtaSpedita] AS 'TRAS', T1.[U_PRG_WIP_QtaDaTrasf] AS 'DA TRAS', sum(CASE WHEN T4.onhand IS NULL THEN 0 ELSE T4.ONHAND END) as 'A_mag',  T0.[U_Lavorazione] AS 'LAV', T0.[U_Priorita_MES] AS 'PRIO', T0.[U_PRODUZIONE] 'Produzione',T0.[DueDate]
FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T2.ITEMCODE=T0.ITEMCODE 
inner join oitm t3 on t3.itemcode=t1.itemcode
left join oitw t4 on t4.itemcode=t3.itemcode

WHERE T0.[Status] ='R'  AND  T1.[ItemType] ='4' and T1.[U_PRG_WIP_QtaDaTrasf]>0 and 
( (substring(T0.[U_PRODUZIONE],1,3) ='INT' AND( t4.whscode='01' or t4.whscode='FERRETTO' or t4.whscode='03' or t4.whscode='MUT' or t4.whscode='SCA')) OR (T0.[U_PRODUZIONE] ='ASSEMBL' AND T3.DFLTWH ='MUT') ) AND (SUBSTRING(T0.[ItemCode],1,1)='C' OR SUBSTRING(T0.[ItemCode],1,1)='D' OR SUBSTRING(T0.[ItemCode],1,1)='0' OR SUBSTRING(T0.[ItemCode],1,1)='F')
group by 
T0.[DocNum] , T0.[ItemCode], T2.[ItemNAME] , T0.[PlannedQty] , t2.u_disegno , t3.itemcode,t3.itemname, T1.[U_PRG_WIP_QtaSpedita] , T1.[U_PRG_WIP_QtaDaTrasf] , T0.[U_Lavorazione] , T0.[U_Priorita_MES] ,T0.[U_PRODUZIONE],t3.u_ubicazione,T0.[DueDate],T0.docentry
)
AS T10 left join oitw t11 on t11.itemcode=t10.itemcode
GROUP BY 
T10.[N° Ordine], T10.DOC, T10.COD, T10.NOME, T10.Q, T10.DIS, t10.produzione, t10.itemcode,t10.itemname, T10.TRAS, T10.[DA TRAS], T10.A_mag,   T10.LAV, T10.PRIO,t10.u_ubicazione,T10.[DueDate],T10.docentry
)
as t20
LEFT JOIN WOR1 T21 ON T21.DOCENTRY=T20.DOCENTRY

group by T20.docentry, T20.[N° Ordine], T20.DOC, T20.COD, T20.NOME, T20.Q, T20.DIS, t20.produzione, t20.itemcode,t20.itemname,t20.u_ubicazione, T20.TRAS, T20.[DA TRAS], T20.A_mag, T20.Stato,  T20.LAV, T20.PRIO,T20.[DueDate]
)
as t30
where t30.prog is null or t30.prog=0
ORDER BY case when substring(t30.produzione,1,3)='int' then t30.prio else T30.[DueDate] end"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Try

            Do While cmd_SAP_reader_2.Read()


                DataGridView_MATERIALE.Rows.Add(cmd_SAP_reader_2("N° Ordine"), cmd_SAP_reader_2("DOC"), cmd_SAP_reader_2("COD"), cmd_SAP_reader_2("NOME"), cmd_SAP_reader_2("Q"), cmd_SAP_reader_2("DIS"), cmd_SAP_reader_2("Produzione"), cmd_SAP_reader_2("itemcode"), cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("u_ubicazione"), cmd_SAP_reader_2("TRAS"), cmd_SAP_reader_2("DA TRAS"), cmd_SAP_reader_2("STATO"), cmd_SAP_reader_2("LAV"), cmd_SAP_reader_2("PRIO"))

            Loop

        Catch ex As Exception
            MsgBox("L'ordine" & cmd_SAP_reader_2("N° Ordine") & "presenta un errore")
        End Try
        cmd_SAP_reader_2.Close()
        cnn1.Close()


    End Sub


    Private Sub DataGridView_MATERIALE_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_MATERIALE.CellClick
        If e.RowIndex >= 0 Then

            Dashboard_MU_New.mu = 1
            RIGA = e.RowIndex
            If e.ColumnIndex = 0 Then



                ODP_Form.docnum_odp = DataGridView_MATERIALE.Rows(RIGA).Cells(0).Value
                ODP_Form.Show()
                ODP_Form.inizializza_form(DataGridView_MATERIALE.Rows(RIGA).Cells(0).Value)


            End If

            If e.ColumnIndex = 2 Then
                    Magazzino.Codice_SAP = DataGridView_MATERIALE.Rows(e.RowIndex).Cells(2).Value

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
                If e.ColumnIndex = 7 Then
                    Magazzino.Codice_SAP = DataGridView_MATERIALE.Rows(e.RowIndex).Cells(7).Value
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

                If e.ColumnIndex = 5 Then

                    Try
                    Process.Start(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_MATERIALE.Rows(e.RowIndex).Cells(5).Value & ".PDF")
                Catch ex As Exception
                        MsgBox("Il disegno " & DataGridView_MATERIALE.Rows(e.RowIndex).Cells(5).Value & " non è ancora stato processato")
                    End Try

                End If

            End If
    End Sub



    Private Sub DataGridView_MATERIALE_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_MATERIALE.CellFormatting

        If Not DataGridView_MATERIALE.Rows(e.RowIndex).Cells(12).Value Is System.DBNull.Value Then
            If DataGridView_MATERIALE.Rows(e.RowIndex).Cells(12).Value = "Trasferibile" Then
                DataGridView_MATERIALE.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
            ElseIf DataGridView_MATERIALE.Rows(e.RowIndex).Cells(12).Value = "IN APPROV" Then
                DataGridView_MATERIALE.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Khaki
            Else
                DataGridView_MATERIALE.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.OrangeRed
            End If
        End If
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        filtra()

    End Sub


    Sub filtra()
        Dim i = 0
        Dim parola0 As String
        Dim parola2 As String
        Dim parola3 As String
        Dim parola6 As String




        Do While i < DataGridView_MATERIALE.RowCount

            Try

                parola0 = UCase(DataGridView_MATERIALE.Rows(i).Cells(0).Value)
                parola2 = UCase(DataGridView_MATERIALE.Rows(i).Cells(2).Value)
                parola3 = UCase(DataGridView_MATERIALE.Rows(i).Cells(3).Value)
                parola6 = UCase(DataGridView_MATERIALE.Rows(i).Cells(6).Value)



                If parola0.Contains(UCase(TextBox8.Text)) Then
                    DataGridView_MATERIALE.Rows(i).Visible = True
                    If parola2.Contains(UCase(TextBox7.Text)) Then
                        DataGridView_MATERIALE.Rows(i).Visible = True


                        If parola3.Contains(UCase(TextBox5.Text)) Then
                            DataGridView_MATERIALE.Rows(i).Visible = True


                            If parola6.Contains(UCase(TextBox1.Text)) Then
                                DataGridView_MATERIALE.Rows(i).Visible = True

                            Else
                                DataGridView_MATERIALE.Rows(i).Visible = False

                            End If


                        Else
                            DataGridView_MATERIALE.Rows(i).Visible = False

                        End If


                    Else
                        DataGridView_MATERIALE.Rows(i).Visible = False

                    End If

                Else
                    DataGridView_MATERIALE.Rows(i).Visible = False

                End If

            Catch ex As Exception
                DataGridView_MATERIALE.Rows(i).Visible = False
            End Try
            i = i + 1
        Loop
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        filtra()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        filtra()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        filtra()
    End Sub
End Class