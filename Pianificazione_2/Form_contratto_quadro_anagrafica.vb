Imports System.Data.SqlClient

Public Class Form_contratto_quadro_anagrafica
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Sub inizializza_form(par_docentry As Integer)
        anagrafica(par_docentry)
        DETTAGLI(par_docentry, DataGridView2)
        documenti(par_docentry, DataGridView1)

    End Sub

    Sub anagrafica(par_docentry As Integer)


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.AbsID, T0.BpCode, T0.BpName, T0.CntctCode, T0.StartDate, T0.EndDate, T0.TermDate, T0.Descript, T0.Type, T0.Status, T0.Owner, T0.Renewal, T0.UseDiscnt, T0.RemindVal, T0.RemindUnit, T0.Remarks, T0.AtchEntry, T0.LogInstanc, T0.UserSign, T0.UserSign2, T0.UpdtDate, T0.CreateDate, T0.Cancelled, T0.DataSource, T0.Transfered, T0.RemindFlg, T0.Attachment, T0.SettleProb, T0.UpdtTime, T0.Method, T0.GroupNum, T0.ListNum, T0.SignDate, T0.AmendedTo, T0.Series, T0.Number, T0.ObjType, T0.Handwrtten, T0.PIndicator, T0.BpType, T0.PayMethod, T0.NumAtCard, T0.BPCurr, T0.FixedRate, T0.TrnspCode, T0.Project, T0.PriceMode, T0.WddStatus, T0.FromStat, T0.DPPStatus, T0.SAPPassprt, T0.EncryptIV, T0.U_PRG_AZS_NoteChiu, T0.U_PRG_AZS_NoteAper 
FROM OOAT T0 
WHERE T0.[AbsID] =" & par_docentry & "
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            Label2.Text = cmd_SAP_reader_2("number")
            Label3.Text = cmd_SAP_reader_2("BpName")
            DateTimePicker1.Value = cmd_SAP_reader_2("StartDate")
            DateTimePicker2.Value = cmd_SAP_reader_2("EndDate")
            TextBox1.Text = cmd_SAP_reader_2("Descript")


        End If

        Cnn1.Close()




    End Sub

    Sub DETTAGLI(par_docentry As Integer, PAR_DATAGRIDVIEW As DataGridView)

        PAR_DATAGRIDVIEW.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.AgrNo, T0.AgrLineNum, T0.VisOrder, T0.ItemCode, T0.ItemName,  T0.PlanQty,T0.CumQty,T0.PlanQty-T0.CumQty as 'Q_aperto'
, T0.UnitPrice 
,T0.UNDLVQTY
,T0.PlanQty*T0.UnitPrice as 'Tot'
,(T0.PlanQty-T0.CumQty)*T0.UnitPrice as 'Tot_aperto'
, T0.CumAmntFC, T0.CumAmntLC, T0.FreeTxt, T0.InvntryUom, T0.LogInstanc,  T0.RetPortion, T0.WrrtyEnd, T0.LineStatus, T0.PlanAmtLC, T0.PlanAmtFC, T0.Discount, T0.UomEntry, T0.UomCode, T0.NumPerMsr, T0.UndlvQty, T0.UndlvAmntL, T0.UndlvAmntF, T0.TrnspCode, T0.Project, T0.TaxCode, T0.TAXRate, T0.PlVatAmtLC, T0.PlVatAmtFC, T0.CumVtAmtLC, T0.CumVtAmtFC, T0.EncryptIV 
FROM OAT1 T0
WHERE T0.AgrNo =" & par_docentry & "
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            PAR_DATAGRIDVIEW.Rows.Add(cmd_SAP_reader_2("visorder") + 1, cmd_SAP_reader_2("ItemCode"), cmd_SAP_reader_2("ItemName"), cmd_SAP_reader_2("PlanQty"), cmd_SAP_reader_2("UNDLVQTY"), cmd_SAP_reader_2("q_aperto"), cmd_SAP_reader_2("UnitPrice"), cmd_SAP_reader_2("Tot"), cmd_SAP_reader_2("Tot_aperto"))



        Loop

        Cnn1.Close()

        PAR_DATAGRIDVIEW.ClearSelection()


    End Sub

    Sub documenti(par_docentry As Integer, PAR_DATAGRIDVIEW As DataGridView)

        PAR_DATAGRIDVIEW.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "declare @contract as integer
set @contract=" & par_docentry & "

-- Ordini di Vendita
SELECT 
    'Ordine di Vendita' AS 'Tipo Documento',
    T0.DocEntry AS 'Numero Documento',
    T0.DocNum AS 'Numero Documento SAP',
    T0.DocDate AS 'Data Documento',
    T0.CardCode AS 'Codice Cliente/Fornitore',
    T0.CardName AS 'Nome Cliente/Fornitore',
    T0.DocTotal AS 'Totale Documento'
FROM 
    ORDR T0 inner join rdr1 t1 on t0.docentry=t1.docentry
WHERE 
    T1.AgrNo = 27 -- Sostituire con l'ID del Contratto Quadro

UNION ALL


-- Consegne
SELECT 
    'Consegna' AS 'Tipo Documento',
    T3.DocEntry AS 'Numero Documento',
    T3.DocNum AS 'Numero Documento SAP',
    T3.DocDate AS 'Data Documento',
    T3.CardCode AS 'Codice Cliente/Fornitore',
    T3.CardName AS 'Nome Cliente/Fornitore',
    T3.DocTotal AS 'Totale Documento'
FROM 
    ODLN T3 inner join DLN1 t1 on t3.docentry=t1.docentry
WHERE 
   T1.AgrNo = @contract



ORDER BY 
    'Data Documento' ASC;

"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            PAR_DATAGRIDVIEW.Rows.Add(cmd_SAP_reader_2("Tipo Documento"), cmd_SAP_reader_2("Numero Documento"), cmd_SAP_reader_2("Numero Documento SAP"), cmd_SAP_reader_2("Data Documento"), cmd_SAP_reader_2("Codice Cliente/Fornitore"), cmd_SAP_reader_2("Nome Cliente/Fornitore"), cmd_SAP_reader_2("Totale Documento"))



        Loop

        Cnn1.Close()

        PAR_DATAGRIDVIEW.ClearSelection()


    End Sub



End Class