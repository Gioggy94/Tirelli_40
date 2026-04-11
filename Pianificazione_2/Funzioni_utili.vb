Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Windows.Annotations
Imports Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class Funzioni_utili
    Public ok_1 As String = "NO"
    Public ok_2 As String = "NO"
    Public Excel As Excel.Application
    Public WB_Excel As Excel.Workbook

    Public docentry As String
    Public docnum As String
    Dim Series As String
    Dim pindicator As String
    Dim Versionnum As String
    Dim JrnlMemo As String
    Dim itemcode As String
    Dim quantità As Integer
    Dim itemname As String
    Dim disegno As String
    Dim revisione As String
    Dim Prezzo As String
    Dim min As String
    Dim q_min_ord As String
    Dim motivazione_stock As String
    Dim gestione_magazzino As String
    Public codice_brb As String
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Sub Ricerca_anagrafica(PAR_CODICE As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        CMD_SAP_2.Connection = Cnn

        CMD_SAP_2.CommandText = " SELECT T0.[ItemCode], T0.[ItemName] 
FROM oitm T0  WHERE T0.[ItemCode] ='" & PAR_CODICE & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        If cmd_SAP_reader_2.Read() Then

            Label1.Text = cmd_SAP_reader_2("ItemName")
            Numero_Codici_da_sostituire()
            righe_da_sostituire()

            ok_1 = "OK"
        Else
            ok_1 = "NO"

        End If
        Cnn.Close()

    End Sub

    Sub Ricerca_anagrafica_per_distinta(par_codice As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        CMD_SAP_2.Connection = Cnn

        CMD_SAP_2.CommandText = " SELECT T0.[ItemCode], T0.[ItemName] FROM oitm T0  WHERE T0.[ItemCode] ='" & par_codice & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        If cmd_SAP_reader_2.Read() Then

            Label7.Text = cmd_SAP_reader_2("ItemName")
            Numero_Codici_da_sostituire_db(cmd_SAP_reader_2("ItemCode"))
            righe_da_sostituire_db(cmd_SAP_reader_2("ItemCode"))
            ok_1 = "OK"
        Else
            ok_1 = "NO"

        End If
        Cnn.Close()

    End Sub



    Sub Ricerca_anagrafica_2(par_codice As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        CMD_SAP_2.Connection = Cnn

        CMD_SAP_2.CommandText = " SELECT T0.[ItemCode], T0.[ItemName], T0.[U_Gestione_magazzino], T0.[frozenFor] FROM oitm T0  
WHERE T0.[ItemCode] ='" & par_codice & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        If cmd_SAP_reader_2.Read() Then

            Label3.Text = cmd_SAP_reader_2("ItemName")
            Label4.Text = cmd_SAP_reader_2("frozenFor")
            Label5.Text = cmd_SAP_reader_2("U_Gestione_magazzino")
            If cmd_SAP_reader_2("frozenFor") = "N" Then
                ok_2 = "OK"
            Else
                ok_2 = "INATTIVO"
            End If
        Else
            ok_2 = "NO"

        End If

        Cnn.Close()

    End Sub

    Sub Ricerca_anagrafica_2_db(par_codice As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        CMD_SAP_2.Connection = Cnn

        CMD_SAP_2.CommandText = " SELECT T0.[ItemCode], T0.[ItemName], T0.[U_Gestione_magazzino], T0.[frozenFor] 
FROM oitm T0  WHERE T0.[ItemCode] ='" & par_codice & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        If cmd_SAP_reader_2.Read() Then

            Label9.Text = cmd_SAP_reader_2("ItemName")
            Label10.Text = cmd_SAP_reader_2("frozenFor")
            Label8.Text = cmd_SAP_reader_2("U_Gestione_magazzino")
            If cmd_SAP_reader_2("frozenFor") = "N" Then
                ok_2 = "OK"
            Else
                ok_2 = "INATTIVO"
            End If
        Else
            ok_2 = "NO"

        End If

        Cnn.Close()

    End Sub

    Sub Numero_Codici_da_sostituire()

        Dim Cnn2 As New SqlConnection
        Cnn2.ConnectionString = Homepage.sap_tirelli
        Cnn2.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn2

        CMD_SAP_1.CommandText = " SELECT T2.[ItemCode], T2.[ItemName], cast(sum(T0.[PlannedQty]) as decimal) as 'Tot' FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] inner join OITM T2 on t0.itemcode=t2.itemcode
WHERE T0.[ItemCode] ='" & TextBox1.Text & "' and (T1.[Status] ='P' or  T1.[Status] ='R') and  T0.[PlannedQty] = t0.u_prg_wip_qtadatrasf

group by T2.[ItemCode], T2.[ItemName]"

        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        If cmd_SAP_reader_1.Read() Then

            Label2.Text = cmd_SAP_reader_1("Tot")


        End If
        Cnn2.Close()

    End Sub

    Sub Numero_Codici_da_sostituire_db(PAR_CODICE As String)

        Dim Cnn2 As New SqlConnection
        Cnn2.ConnectionString = Homepage.sap_tirelli
        Cnn2.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn2

        CMD_SAP_1.CommandText = " SELECT COUNT(T0.CODE) as 'Tot'
FROM ITT1 T0
WHERE T0.CODE='" & PAR_CODICE & "'"

        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        If cmd_SAP_reader_1.Read() Then

            Label6.Text = cmd_SAP_reader_1("Tot")


        End If
        Cnn2.Close()

    End Sub

    Sub righe_da_sostituire()

        DataGridView4.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1

        CMD_SAP_1.CommandText = " SELECT T2.[ItemCode], T2.[ItemName], t1.docnum, T0.[PlannedQty]  as 'Plannedqty', T1.[U_PRG_AZS_Commessa], t1.u_utilizz FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] inner join OITM T2 on t0.itemcode=t2.itemcode
WHERE T0.[ItemCode] ='" & TextBox1.Text & "' and (T1.[Status] ='P' or  T1.[Status] ='R') and  T0.[PlannedQty] = t0.u_prg_wip_qtadatrasf"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()

            DataGridView4.Rows.Add(cmd_SAP_reader_1("docnum"), cmd_SAP_reader_1("Plannedqty"), cmd_SAP_reader_1("U_PRG_AZS_Commessa"), cmd_SAP_reader_1("u_utilizz"))

        Loop
        Cnn1.Close()

    End Sub

    Sub righe_da_sostituire_db(par_codice As String)

        DataGridView1.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1

        CMD_SAP_1.CommandText = " SELECT t0.father, t1.itemname, t0.Quantity
FROM ITT1 T0 inner join oitm t1 on t0.father=t1.itemcode
WHERE T0.CODE='" & par_codice & "'"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()

            DataGridView1.Rows.Add(cmd_SAP_reader_1("father"), cmd_SAP_reader_1("itemname"), cmd_SAP_reader_1("Quantity"))

        Loop
        Cnn1.Close()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Ricerca_anagrafica(TextBox1.Text)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Ricerca_anagrafica_2(TextBox2.Text)
    End Sub

    Sub sostituzione_codici()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "update wor1 set wor1.itemcode='" & TextBox2.Text & "'
from
(
SELECT T2.[ItemCode], T2.[ItemName], t1.docentry, t1.docnum, t0.linenum, T0.[PlannedQty]  as 'Plannedqty', T1.[U_PRG_AZS_Commessa], t1.u_utilizz FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] inner join OITM T2 on t0.itemcode=t2.itemcode
WHERE T0.[ItemCode] ='" & TextBox1.Text & "' and (T1.[Status] ='P' or  T1.[Status] ='R') and  T0.[PlannedQty] = t0.u_prg_wip_qtadatrasf
)
as t10 inner join wor1 t11 on t11.linenum=t10.linenum and t11.docentry=t10.docentry"
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Sub sostituzione_codici_in_distinta(par_codice_da_sostituire As String, par_codice_sostituente As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "
UPDATE T1 set t1.code='" & par_codice_sostituente & "', t1.price=t3.price, t1.[ItemName]=t2.itemname
FROM OITM T0 INNER JOIN ITT1 T1 ON T0.ITEMCODE=T1.CODE
inner join oitm t2 on t2.itemcode='" & par_codice_sostituente & "'
inner join itm1 t3 on t3.pricelist=2 and t3.itemcode=t2.itemcode
WHERE T0.ITEMCODE='" & par_codice_da_sostituire & "'"
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Sub Riparare_i_confermati_1()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "update t41 set t41.iscommited=t40.confermati
from
(

SELECT t10.itemcode,sum(t10.confermati) AS 'CONFERMATI', t10.MAG
FROM
(
SELECT T0.ITEMCODE, 0 AS 'CONFERMATI', T0.WHSCODE AS 'MAG'
FROM OITW T0
WHERE T0.ISCOMMITED>0 OR T0.ISCOMMITED<0
UNION ALL
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
as t40 inner join oitw t41 on t41.itemcode=t40.itemcode and t41.whscode=t40.mag
where t40.itemcode = '" & TextBox1.Text & "'
SELECT T0.DOCNUM FROM OWOR T0 WHERE T0.DOCNUM=5"
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Sub Riparare_i_confermati_2()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "update t41 set t41.iscommited=t40.confermati
from
(

SELECT t10.itemcode,sum(t10.confermati) AS 'CONFERMATI', t10.MAG
FROM
(
SELECT T0.ITEMCODE, 0 AS 'CONFERMATI', T0.WHSCODE AS 'MAG'
FROM OITW T0
WHERE T0.ISCOMMITED>0 OR T0.ISCOMMITED<0
UNION ALL
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
as t40 inner join oitw t41 on t41.itemcode=t40.itemcode and t41.whscode=t40.mag
where t40.itemcode = '" & TextBox2.Text & "'
SELECT T0.DOCNUM FROM OWOR T0 WHERE T0.DOCNUM=5"
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ok_1 = "OK" And ok_2 = "OK" Then
            sostituzione_codici()
            Riparare_i_confermati_1()
            Riparare_i_confermati_2()
            MsgBox("Articolo sostituito con successo")
        ElseIf ok_2 = "INATTIVO" Then
            MsgBox("Articolo da inserire inattivo, non è possibile proseguire")
        Else
            MsgBox("Non è possibile proseguire, controllare di aver inserito correttamente i codici")

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Stampante_3D.Show()
    End Sub

    Sub importazione_excel()
        Dim contatore_excel As Integer = 2
        Dim lunghezza_excel As Integer = 9486


        Dim NOME_file As String = "Match articoli Tirelli BRB"
        Dim NOME_FOGLIO As String = "BRB"


        Excel = CreateObject("Excel.application")
        WB_Excel = Excel.Workbooks.Open("C:\Users\giovannitirelli\Desktop\" & NOME_file & ".xlsx")

        Excel.Visible = True

        Dim desc_sup As String
        Dim catalogo_fornitore As String
        Dim fornitore As String

        Do While contatore_excel <= lunghezza_excel

            itemcode = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 1).value
            itemname = Replace(Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 2).value, "'", "")
            desc_sup = Replace(Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 3).value, "'", "")
            catalogo_fornitore = Replace(Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 5).value, "'", "")
            fornitore = Replace(Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 7).value, "'", "")

            ' gestione_magazzino = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 2).value
            ' min = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 6).value
            'q_min_ord = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 5).value
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli

            Cnn.Open()

            Dim Cmd_SAP As New SqlCommand

            Cmd_SAP.Connection = Cnn
            Cmd_SAP.CommandText = "
INSERT INTO [TIRELLI_40].[DBO].BRB_Codici
           ([Codice_BRB]
           ,[Descrizione_BRB]
           ,[Descrizione_supp_BRB]
           ,[Catalogo_fornitore]
           ,[Fornitore]
           ,[Ubicazione]
           ,[Costo])
     VALUES
           ('" & itemcode & "'
           ,'" & itemname & "'
           ,'" & desc_sup & "'
           ,'" & catalogo_fornitore & "'
           ,'" & fornitore & "'
           ,''
           ,0)"


            Cmd_SAP.ExecuteNonQuery()



            Cnn.Close()

            contatore_excel = contatore_excel + 1



        Loop
        MsgBox("importazione effettuata con successo")
    End Sub

    Sub importazione_matricole()
        Dim contatore_excel As Integer = 893

        Dim lunghezza_excel As Integer = 1398


        Dim NOME_file As String = "Importazione"
        Dim NOME_FOGLIO As String = "Foglio1"


        Excel = CreateObject("Excel.application")
        WB_Excel = Excel.Workbooks.Open("C:\Users\giovannitirelli\Desktop\" & NOME_file & ".xlsx")

        Excel.Visible = True

        Dim descrizione As String
        Dim Descrizione_supp As String
        Dim paese As String
        Dim agente As String
        Dim cliente As String
        Dim anno As String

        Do While contatore_excel <= lunghezza_excel



            itemcode = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 2).value
            descrizione = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 6).value
            Descrizione_supp = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 7).value
            cliente = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 10).value
            paese = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 11).value
            agente = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 9).value
            anno = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 14).value

            'MsgBox("ITEMCode " & itemcode)
            'MsgBox("Descrizione " & descrizione)
            'MsgBox("Descrizione sup " & Descrizione_supp)
            'MsgBox("Cliente " & cliente)
            'MsgBox("paese " & paese)
            'MsgBox("agente " & agente)

            'MsgBox("anno " & anno)


            ', TextBox_codice_sap.Text, Replace(TextBox_descrizione.Text, ",", " "), Replace(TextBox_DESC_SUPP.Text, ",", " "), Replace(TextBox_disegno.Text, ",", " "), Elenco_gruppi(ComboBox_gruppo_articoli.SelectedIndex), Elenco_fornitori(ComboBox5.SelectedIndex), Replace(TextBox1.Text, ",", " "), Elenco_produttori(ComboBox3.SelectedIndex + 1), ComboBox_tipo_montaggio.Text, Replace(TextBox10.Text, ",", " "), Replace(TextBox5.Text, ",", " "), SETTORE, paese, Replace(TextBox8.Text, ",", " "), brand

            'UT.inserisci_Nuovo_codice(Homepage.UTENTE_SAP_SALVATO, itemcode, descrizione, Descrizione_supp, "", 104, "", "", "-", "", "", cliente, "", paese, agente, "BRB")
            contatore_excel = contatore_excel + 1


        Loop
        MsgBox("importazione effettuata con successo")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        trova_dato_da_excel_pEr_importazionE("C:\Users\giovannitirelli\Desktop\Aggiornamento_dipendenti.xlsx", "Foglio1", 2, 327)
    End Sub

    Sub trova_dato_da_excel_pEr_importazionE(par_percorso_file As String, par_nome_foglio As String, par_riga_inizio As Integer, par_riga_fine As Integer)

        Dim colonna1 As String
        Dim colonna2 As String
        Dim colonna3 As String
        Dim colonna4 As String
        Dim colonna5 As String
        Dim colonna6 As String

        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True


        Do While par_riga_inizio <= par_riga_fine


            If Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value <> Nothing Then
                colonna1 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value
                colonna2 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value
                colonna3 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 3).value
                'colonna4 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value
                '  colonna5 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 5).value
                '  colonna6 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 6).value

                ' crea_codice_articolo(colonna1, colonna2, colonna3, colonna4, colonna5, colonna6)
                'insert_into_codici(colonna1, colonna2, colonna3, colonna4, colonna5, colonna6)
                UPDATE(colonna1, colonna2, colonna3, colonna4, colonna5)
            End If
            par_riga_inizio = par_riga_inizio + 1
        Loop
        Beep()
        MsgBox("Importazione effettuata con successo")


    End Sub

    Sub insert_into_codici(par_Colonna_1 As String, par_Colonna_2 As String, par_Colonna_3 As String, par_Colonna_4 As String, par_colonna_5 As String, par_colonna_6 As String)

        par_Colonna_1 = Replace(par_Colonna_1, "'", " ")
        par_Colonna_2 = Replace(par_Colonna_2, "'", " ")
        par_Colonna_3 = Replace(par_Colonna_3, "'", " ")
        par_Colonna_4 = Replace(par_Colonna_4, "'", " ")
        par_colonna_5 = Replace(par_colonna_5, "'", " ")
        par_colonna_6 = Replace(par_colonna_6, ",", ".")


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1
        Cmd_SAP.CommandText = "INSERT INTO [tirelli_40].[dbo].[Cassetto_codici]
           ([Codice]
           ,[Magazzino]
           ,[Cassetto])
     VALUES
           ('" & par_Colonna_1 & "'
           ," & par_Colonna_2 & "
           ," & par_Colonna_3 & "
           )
          "

        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()

    End Sub

    Sub crea_codice_articolo(par_Colonna_1 As String, par_Colonna_2 As String, par_Colonna_3 As String, par_Colonna_4 As String, par_colonna_5 As String, par_colonna_6 As String)
        UT.Show()
        UT.ComboBox_prima_lettera.SelectedIndex = 2
        UT.ComboBox_prima_lettera.SelectedIndex = 1
        UT.ComboBox_gruppo_articoli.Text = "Attrezzatura a formato"
        par_Colonna_1 = Replace(par_Colonna_1, "'", " ")
        par_Colonna_2 = Replace(par_Colonna_2, "'", " ")
        par_Colonna_3 = Replace(par_Colonna_3, "'", " ")
        par_Colonna_4 = Replace(par_Colonna_4, "'", " ")
        par_colonna_5 = Replace(par_colonna_5, "'", " ")
        par_colonna_6 = Replace(par_colonna_6, ",", ".")

        UT.TextBox_descrizione.Text = par_Colonna_2
        UT.TextBox_disegno.Text = par_Colonna_1
        UT.TextBox_DESC_SUPP.Text = par_Colonna_4

        UT.Aggiungi.PerformClick()

    End Sub

    Sub UPDATE(par_Colonna_1 As String, par_Colonna_2 As String, par_Colonna_3 As String, par_Colonna_4 As String, par_Colonna_5 As String)

        par_Colonna_1 = Replace(par_Colonna_1, "'", " ")
        par_Colonna_2 = Replace(par_Colonna_2, "'", " ")
        par_Colonna_3 = Replace(par_Colonna_3, "'", " ")
        par_Colonna_4 = Replace(par_Colonna_4, "'", " ")
        par_Colonna_5 = Replace(par_Colonna_5, "'", " ")

        par_Colonna_2 = Replace(par_Colonna_2, ",", ".")
        par_Colonna_3 = Replace(par_Colonna_3, ",", ".")


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1
        Cmd_SAP.CommandText = "
        UPDATE [Tirelli_40].[dbo].[OHEM]
        SET [Galileo]='" & par_Colonna_2 & "'
      

                where empid='" & par_Colonna_1 & "' "


        ', 
        'T0.MINLEVEL ='" & par_Colonna_3 & "',
        'T0.MINORDRQTY ='" & par_Colonna_2 & "',
        'U_UBIMAg ='" & par_Colonna_4 & "',
        'u_data_valutazione_stock = getdate()
        '        Cmd_SAP.CommandText = "update t0 
        'set t0.u_produzione='EST'
        'FROM owor T0 
        'WHERE T0.[docnum] ='" & par_Colonna_1 & "' "

        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Magazzino.Trova_NUMERATORE_OIVL()
        Magazzino.Trova_message_id()
        Magazzino.aggiusta_Numeratore_OIVL()
        Magazzino.Aggiusta_numeratore_messageid()
        MsgBox("Riparato")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Form_mail_ricambi.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        PDM_SAP.Show()

    End Sub







    Sub trasferito(par_codice_SAP As String)


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT  t0.[ID]
      ,t0.[Data_import]
      ,t0.[Data_export]
      ,t0.[Esportato]
      ,t0.[Codice_BRB]
      ,t0.[Codice_SAP]
      ,t0.[Descrizione]
      ,t0.[Descrizione_supp]
  FROM [TIRELLISRLDB].[dbo].[Frontiera_PDM_BRB_SAP] t0
  where t0.[Esportato]='N'
  order by t0.id"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()


            codice_brb = cmd_SAP_reader_2("Codice_BRB")
        Loop

        cmd_SAP_reader_2.Close()
        cnn1.Close()

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Ricerca_anagrafica_per_distinta(TextBox3.Text)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Ricerca_anagrafica_2_db(TextBox4.Text)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If ok_1 = "OK" And ok_2 = "OK" Then
            sostituzione_codici_in_distinta(TextBox3.Text, TextBox4.Text)
            MsgBox("Articolo sostituito con successo")
        ElseIf ok_2 = "INATTIVO" Then
            MsgBox("Articolo da inserire inattivo, non è possibile proseguire")
        Else
            MsgBox("Non è possibile proseguire, controllare di aver inserito correttamente i codici")

        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If TextBox6.Text = "" Then
            MsgBox("Inserire un lotto di prelievo")
        Else
            Beep()
            prelievo_a_ferretto_lotto_di_prelievo(TextBox6.Text, ComboBox1.Text, "ODP")
            MsgBox("Prelievo a FERRETTO lanciato con successo")
        End If

    End Sub

    ' Struttura per INPUT
    <StructLayout(LayoutKind.Sequential)>
    Private Structure INPUT
        Public type As Integer
        Public mi As MOUSEINPUT
    End Structure

    ' Struttura per MOUSEINPUT
    <StructLayout(LayoutKind.Sequential)>
    Private Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As Integer
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    ' Importiamo la funzione SendInput
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SendInput(nInputs As Integer, ByRef pInputs As INPUT, cbSize As Integer) As UInteger
    End Function

    ' Importiamo la funzione SetCursorPos per spostare il mouse
    <DllImport("user32.dll")>
    Private Shared Function SetCursorPos(x As Integer, y As Integer) As Boolean
    End Function

    ' Costanti per il click sinistro
    Private Const INPUT_MOUSE As Integer = 0
    Private Const MOUSEEVENTF_LEFTDOWN As Integer = &H2
    Private Const MOUSEEVENTF_LEFTUP As Integer = &H4

    ' Funzione per simulare un click sinistro del mouse
    Private Sub ClickMouse()
        Dim inputs(1) As INPUT

        ' Premere il pulsante sinistro
        inputs(0).type = INPUT_MOUSE
        inputs(0).mi.dwFlags = MOUSEEVENTF_LEFTDOWN

        ' Rilasciare il pulsante sinistro
        inputs(1).type = INPUT_MOUSE
        inputs(1).mi.dwFlags = MOUSEEVENTF_LEFTUP

        ' Inviare gli input
        SendInput(2, inputs(0), Marshal.SizeOf(GetType(INPUT)))
    End Sub

    Private Const MOUSEEVENTF_RIGHTDOWN As Integer = &H8
    Private Const MOUSEEVENTF_RIGHTUP As Integer = &H10

    Private Sub RightClickMouse()
        Dim inputs(1) As INPUT

        ' Premere il pulsante destro
        inputs(0).type = INPUT_MOUSE
        inputs(0).mi.dwFlags = MOUSEEVENTF_RIGHTDOWN

        ' Rilasciare il pulsante destro
        inputs(1).type = INPUT_MOUSE
        inputs(1).mi.dwFlags = MOUSEEVENTF_RIGHTUP

        ' Inviare l'input
        SendInput(2, inputs(0), Marshal.SizeOf(GetType(INPUT)))
    End Sub

    ' Sub per spostare il mouse e fare click dopo 3 secondi
    Private Sub lancia_ordini(par_odp As Integer, par_utilizzatore As String, par_tipo_ord As String)



        Dim x_binocolo   ' Coordinata X
        Dim y_binocolo   ' Coordinata Y

        Dim x_ricerca_odp ' Coordinata X
        Dim y_ricerca_odp  ' Coordinata Y

        Dim x_tasto_dx   ' Coordinata X
        Dim Y_tasto_dx   ' Coordinata
        ' 
        Dim x_trasferimento ' Coordinata X
        Dim Y_trasferimento   ' Coordinata Y



        Dim x_filtro_mag As Integer   ' Coordinata X
        Dim Y_filtro_mag As Integer   ' Coordinata

        Dim x_bottone_filtro ' Coordinata X
        Dim Y_bottone_filtro    ' Coordinata

        Dim x_bozza_trasferimento  ' Coordinata X
        Dim Y_bozza_trasferimento  ' Coordinata

        Dim x_aggiungere
        Dim y_aggiungere

        Dim x_esci As Integer
        Dim y_esci As Integer


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        CMD_SAP_2.Connection = Cnn

        CMD_SAP_2.CommandText = " SELECT TOP (1000) [Utente]
      ,[X_binocolo]
      ,[Y_Binocolo]
      ,[X_ricerca_odp]
      ,[Y_ricerca_odp]
      ,[X_tasto_DX]
      ,[Y_tasto_DX]
      ,[X_trasferimento_mag]
      ,[Y_trasferimento_mag]
      ,[X_filtro_mag]
      ,[Y_filtro_mag]
      ,[X_button_filtro]
      ,[Y_button_filtro]
      ,[X_bozza_trasf]
      ,[Y_bozza_trasf]
      ,[X_aggiungere]
      ,[Y_aggiungere]
      ,[X_esci]
      ,[Y_esci]
,[X_trasferimento_mag_oc]
      ,[Y_trasferimento_mag_oc]
  FROM [Tirelli_40].[dbo].[Trasferimento_auto_ferretto]
where utente='" & par_utilizzatore & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        If cmd_SAP_reader_2.Read() Then



            x_binocolo = cmd_SAP_reader_2("X_binocolo")  ' Coordinata X
            y_binocolo = cmd_SAP_reader_2("y_binocolo")   ' Coordinata Y

            x_ricerca_odp = cmd_SAP_reader_2("X_ricerca_odp")  ' Coordinata X
            y_ricerca_odp = cmd_SAP_reader_2("y_ricerca_odp")   ' Coordinata Y

            x_tasto_dx = cmd_SAP_reader_2("X_tasto_DX")  ' Coordinata X
            Y_tasto_dx = cmd_SAP_reader_2("Y_tasto_DX")   ' Coordinata
            ' 
            ' Coordinata X
            If par_tipo_ord = "ODP" Then
                Y_trasferimento = cmd_SAP_reader_2("Y_trasferimento_mag")
                x_trasferimento = cmd_SAP_reader_2("X_trasferimento_mag")
            Else par_tipo_ord = "OC"
                Y_trasferimento = cmd_SAP_reader_2("Y_trasferimento_mag_oc")
                x_trasferimento = cmd_SAP_reader_2("X_trasferimento_mag_oc")
            End If

            ' Coordinata Y

            x_filtro_mag = cmd_SAP_reader_2("X_filtro_mag")  ' Coordinata X
            Y_filtro_mag = cmd_SAP_reader_2("Y_filtro_mag")   ' Coordinata

            x_bottone_filtro = cmd_SAP_reader_2("X_button_filtro")  ' Coordinata X
            Y_bottone_filtro = cmd_SAP_reader_2("Y_button_filtro")   ' Coordinata

            x_bozza_trasferimento = cmd_SAP_reader_2("X_bozza_trasf")
            Y_bozza_trasferimento = cmd_SAP_reader_2("Y_bozza_trasf")

            x_aggiungere = cmd_SAP_reader_2("X_aggiungere")
            y_aggiungere = cmd_SAP_reader_2("Y_aggiungere")

            x_esci = cmd_SAP_reader_2("X_esci")
            y_esci = cmd_SAP_reader_2("Y_esci")

        End If
        Cnn.Close()




        SleepWithEscapeCheck(12000) ' Aspetta 3 secondi
        'mi posiziono su binoclo
        SetCursorPos(x_binocolo, y_binocolo)
        SleepWithEscapeCheck(1000) ' Aspetta 3 secondi
        ClickMouse()

        'mi posiziono su barra di ricerca ODP
        SetCursorPos(x_ricerca_odp, y_ricerca_odp)
        SleepWithEscapeCheck(1000) ' Aspetta 3 secondi
        ClickMouse()
        SleepWithEscapeCheck(2000)
        'inserisco valore
        SendKeys.SendWait(par_odp)
        SleepWithEscapeCheck(1000) ' 
        SendKeys.SendWait("{ENTER}") ' 
        If par_tipo_ord = "ODP" Then
            SleepWithEscapeCheck(15000)
        ElseIf par_tipo_ord = "OC" Then
            SleepWithEscapeCheck(25000)
        End If


        'mi metto in centro all'odp
        SetCursorPos(x_tasto_dx, Y_tasto_dx)

        RightClickMouse()
        SleepWithEscapeCheck(1500)
        'premo trasferimento
        SetCursorPos(x_trasferimento, Y_trasferimento)

        SleepWithEscapeCheck(3000)


        ClickMouse()
        'mi posiziono su barra filtro
        SetCursorPos(x_filtro_mag, Y_filtro_mag)
        SleepWithEscapeCheck(1000)
        ClickMouse()
        SleepWithEscapeCheck(1000)
        SendKeys.SendWait("FER") ' 
        SleepWithEscapeCheck(1500)

        'Premo il filtro
        SetCursorPos(x_bottone_filtro, Y_bottone_filtro)
        SleepWithEscapeCheck(500)
        ClickMouse()
        SleepWithEscapeCheck(3500)

        'Mi posiziono su trasferimento magazzino
        SetCursorPos(x_bozza_trasferimento, Y_bozza_trasferimento)
        SleepWithEscapeCheck(1500)
        ClickMouse()
        SleepWithEscapeCheck(10000)


        'Aggiungo
        SetCursorPos(x_aggiungere, y_aggiungere)
        Thread.Sleep(1000)
        ClickMouse()
        SleepWithEscapeCheck(8000)


        'Chiudo
        SetCursorPos(x_esci, y_esci)
        Thread.Sleep(500)
        ClickMouse()


    End Sub

    ' Importiamo la funzione GetCursorPos dalla libreria user32.dll
    <DllImport("user32.dll")>
    Private Shared Function GetCursorPos(ByRef lpPoint As POINT) As Boolean
    End Function

    ' Struttura per memorizzare le coordinate del mouse
    Private Structure POINT
        Public X As Integer
        Public Y As Integer
    End Structure

    ' Timer per aggiornare le coordinate
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim pos As POINT
        If GetCursorPos(pos) Then
            Label11.Text = "X: " & pos.X
            Label12.Text = "Y: " & pos.Y
        End If
    End Sub

    ' Avvia il timer quando si avvia il form
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Interval = 50  ' Aggiorna ogni 50 millisecondi
        Timer1.Start()

        Inserimento_postazioni(ComboBox1)
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(vKey As Integer) As Short
    End Function

    Private Sub SleepWithEscapeCheck(milliseconds As Integer)
        Dim start As Integer = Environment.TickCount
        Do While Environment.TickCount - start < milliseconds
            If GetAsyncKeyState(Keys.Escape) <> 0 Then
                Exit Sub ' Esce immediatamente dalla funzione
            End If
            Thread.Sleep(10) ' Pausa breve per non bloccare il sistema
        Loop
    End Sub

    Sub prelievo_a_ferretto_lotto_di_prelievo(par_lotto_prelievo As Integer, par_utilizzatore As String, par_tipo_documento As String)


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        CMD_SAP_2.Connection = Cnn1

        '        If par_tipo_documento = "ODP" Then
        '            CMD_SAP_2.CommandText = "SELECT COUNT(*) FROM (
        '    SELECT t10.docnum
        '    FROM (
        '        SELECT t1.docnum
        '        FROM [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 
        '        LEFT JOIN owor t1 ON t1.docnum = t0.docnum AND t0.tipo_doc = 'ODP'
        '        LEFT JOIN WOR1 T2 ON T2.DOCENTRY = T1.DOCENTRY
        '        INNER JOIN oitw t3 ON t3.itemcode = t2.itemcode AND t3.whscode = t2.wareHouse 
        '        LEFT JOIN wtq1 t4 ON t4.itemcode = t3.itemcode AND t4.FromWhsCod = t2.wareHouse AND T4.[OpenQty] > 0 AND T4.[LineStatus] = 'O'
        '        WHERE T0.ID = " & par_lotto_prelievo & " AND T1.STATUS = 'R' AND T2.wareHouse = 'FERRETTO' 
        '        AND coalesce(t2.U_PRG_WIP_QtaRichMagAuto,0) < coalesce(t2.U_PRG_WIP_QtaDaTrasf,0) AND COALESCE(t3.onhand, 0) >= 1
        '        GROUP BY t1.docnum
        '    ) AS t10
        ') AS t20"
        '        ElseIf par_tipo_documento = "OC" Then

        '            CMD_SAP_2.CommandText = "SELECT COUNT(*) FROM ( 
        '    SELECT t10.docnum, SUM(CASE WHEN t10.onhand - t10.richiedibile - t10.Richieste_aperte < 0 THEN 1 ELSE 0 END) AS 'Non_richiedibile' 
        '    FROM ( 
        '        SELECT t1.docnum, t3.onhand, coalesce(t2.U_Datrasferire,0) - coalesce(t2.U_PRG_WIP_QtaRichMagAuto,0) AS 'richiedibile', 
        '               SUM(COALESCE(T4.OPENQTY, 0)) AS 'Richieste_aperte' 
        '        FROM [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 
        '        LEFT JOIN ordr t1 ON t1.docnum = t0.docnum AND t0.tipo_doc = 'OC' 
        '        LEFT JOIN rdr1 T2 ON T2.DOCENTRY = T1.DOCENTRY 
        '        INNER JOIN oitw t3 ON t3.itemcode = t2.itemcode AND t3.whscode = t2.whscode
        '        LEFT JOIN wtq1 t4 ON t4.itemcode = t3.itemcode AND t4.FromWhsCod = t2.whscode AND T4.[OpenQty] > 0 AND T4.[LineStatus] = 'O' 
        '        WHERE T0.ID =  " & par_lotto_prelievo & " AND T1.DocStatus = 'O' AND T2.whscode = 'FERRETTO'  and t2.OpenQty>0
        '              AND coalesce(t2.U_PRG_WIP_QtaRichMagAuto,0) < t2.U_Datrasferire AND COALESCE(t3.onhand, 0) >= 1 
        '        GROUP BY t1.docnum, t2.U_Datrasferire - t2.U_PRG_WIP_QtaRichMagAuto, t3.onhand 
        '    ) AS T10 
        '    GROUP BY t10.docnum 
        ') AS t20 
        'WHERE t20.Non_richiedibile = 0"

        '        End If


        'Dim total As Integer = Convert.ToInt32(CMD_SAP_2.ExecuteScalar())

        Dim cmd_SAP_reader_2 As SqlDataReader
        If par_tipo_documento = "ODP" Then
            CMD_SAP_2.CommandText = "   select *
   from
   (
   select t20.docnum, sum(case when t20.Max_Trasferibile-richiesta_mag_auto-Richieste_aperte<0 then 1 else 0 end) 'Non_richiedibile'
   from
   (
   select t10.docnum, case when t10.onhand<=t10.Da_trasf then t10.onhand else t10.Da_trasf end as 'Max_Trasferibile'
   , t10.richiesta_mag_auto
   , t10.Richieste_aperte


   from
   (
   SELECT t1.docnum,t2.itemcode, t3.onhand, coalesce(t2.U_PRG_WIP_QtaDaTrasf,0) as 'Da_trasf' , coalesce(t2.U_PRG_WIP_QtaRichMagAuto,0) AS 'richiesta_mag_auto', 
               SUM(COALESCE(T4.OPENQTY, 0)) AS 'Richieste_aperte' 
        FROM [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 
        LEFT JOIN owor t1 ON t1.docnum = t0.docnum AND t0.tipo_doc = 'ODP' 
        LEFT JOIN WOR1 T2 ON T2.DOCENTRY = T1.DOCENTRY 
        INNER JOIN oitw t3 ON t3.itemcode = t2.itemcode AND t3.whscode = t2.wareHouse 
        LEFT JOIN wtq1 t4 ON t4.itemcode = t3.itemcode AND t4.FromWhsCod = t2.wareHouse AND T4.[OpenQty] > 0 AND T4.[LineStatus] = 'O' 
        WHERE T0.ID = " & par_lotto_prelievo & "
	AND T1.STATUS = 'R' AND T2.wareHouse = 'FERRETTO' 
          AND coalesce(t2.U_PRG_WIP_QtaRichMagAuto,0) < coalesce(t2.U_PRG_WIP_QtaDaTrasf,0) AND COALESCE(t3.onhand, 0) >= 1 
		  --and t1.docnum=229347 
		  --and t2.itemcode='D68295'

        GROUP BY t1.docnum,t2.itemcode, coalesce(t2.U_PRG_WIP_QtaDaTrasf,0) , coalesce(t2.U_PRG_WIP_QtaRichMagAuto,0), t3.onhand 
		)
		as t10
		)
		as t20
		
		GROUP BY t20.docnum 
	)
	as t30
	where t30.Non_richiedibile = 0"
        ElseIf par_tipo_documento = "OC" Then

            CMD_SAP_2.CommandText = "SELECT t20.docnum FROM ( 
    SELECT t10.docnum, SUM(CASE WHEN t10.onhand - t10.richiedibile - t10.Richieste_aperte < 0 THEN 1 ELSE 0 END) AS 'Non_richiedibile' 
    FROM ( 
        SELECT t1.docnum, t3.onhand, coalesce(t2.U_Datrasferire,0) - coalesce(t2.U_PRG_WIP_QtaRichMagAuto,0) AS 'richiedibile', 
               SUM(COALESCE(T4.OPENQTY, 0)) AS 'Richieste_aperte' 
        FROM [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 
        LEFT JOIN ordr t1 ON t1.docnum = t0.docnum AND t0.tipo_doc = 'OC' 
        LEFT JOIN rdr1 T2 ON T2.DOCENTRY = T1.DOCENTRY 
        INNER JOIN oitw t3 ON t3.itemcode = t2.itemcode AND t3.whscode = t2.whscode
        LEFT JOIN wtq1 t4 ON t4.itemcode = t3.itemcode AND t4.FromWhsCod = t2.whscode AND T4.[OpenQty] > 0 AND T4.[LineStatus] = 'O' 
        WHERE T0.ID = " & par_lotto_prelievo & " AND T1.DocStatus = 'O' AND T2.whscode = 'FERRETTO'  and t2.OpenQty>0
              AND coalesce(t2.U_PRG_WIP_QtaRichMagAuto,0) < coalesce(t2.U_Datrasferire,0) AND COALESCE(t3.onhand, 0) >= 1 
        GROUP BY t1.docnum, coalesce(t2.U_Datrasferire,0) - coalesce(t2.U_PRG_WIP_QtaRichMagAuto,0), t3.onhand 
    ) AS T10 
    GROUP BY t10.docnum 
) AS t20 
WHERE t20.Non_richiedibile = 0"
        End If


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader()

        Dim counter As Integer = 0

        Do While cmd_SAP_reader_2.Read()



            lancia_ordini(cmd_SAP_reader_2("docnum"), par_utilizzatore, par_tipo_documento)
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    Private Sub Banner_Tick(sender As Object, e As EventArgs) Handles Banner_timer.Tick
        'Banner.Close()
        'Banner_timer.Stop()
    End Sub

    Sub Trova_coordinate(PAR_utente As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        CMD_SAP_2.Connection = Cnn

        CMD_SAP_2.CommandText = " SELECT TOP (1000) [Utente]
      ,[X_binocolo]
      ,[Y_Binocolo]
      ,[X_ricerca_odp]
      ,[Y_ricerca_odp]
      ,[X_tasto_DX]
      ,[Y_tasto_DX]
      ,[X_trasferimento_mag]
      ,[Y_trasferimento_mag]
      ,[X_filtro_mag]
      ,[Y_filtro_mag]
      ,[X_button_filtro]
      ,[Y_button_filtro]
      ,[X_bozza_trasf]
      ,[Y_bozza_trasf]
      ,[X_aggiungere]
      ,[Y_aggiungere]
      ,[X_esci]
      ,[Y_esci]
 ,[X_trasferimento_mag_OC]
      ,[Y_trasferimento_mag_OC]
  FROM [Tirelli_40].[dbo].[Trasferimento_auto_ferretto]
where utente='" & PAR_utente & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        If cmd_SAP_reader_2.Read() Then

            TextBox7.Text = cmd_SAP_reader_2("X_binocolo")
            TextBox8.Text = cmd_SAP_reader_2("y_binocolo")
            TextBox10.Text = cmd_SAP_reader_2("X_ricerca_odp")
            TextBox9.Text = cmd_SAP_reader_2("y_ricerca_odp")
            TextBox12.Text = cmd_SAP_reader_2("X_tasto_DX")
            TextBox11.Text = cmd_SAP_reader_2("Y_tasto_DX")
            TextBox14.Text = cmd_SAP_reader_2("X_trasferimento_mag")
            TextBox13.Text = cmd_SAP_reader_2("Y_trasferimento_mag")
            TextBox16.Text = cmd_SAP_reader_2("X_filtro_mag")
            TextBox15.Text = cmd_SAP_reader_2("Y_filtro_mag")
            TextBox18.Text = cmd_SAP_reader_2("X_button_filtro")
            TextBox17.Text = cmd_SAP_reader_2("Y_button_filtro")
            TextBox20.Text = cmd_SAP_reader_2("X_bozza_trasf")
            TextBox19.Text = cmd_SAP_reader_2("Y_bozza_trasf")
            TextBox22.Text = cmd_SAP_reader_2("X_aggiungere")
            TextBox21.Text = cmd_SAP_reader_2("Y_aggiungere")
            TextBox24.Text = cmd_SAP_reader_2("X_esci")
            TextBox23.Text = cmd_SAP_reader_2("Y_esci")
            TextBox27.Text = cmd_SAP_reader_2("X_trasferimento_mag_OC")
            TextBox26.Text = cmd_SAP_reader_2("Y_trasferimento_mag_OC")

        End If
        Cnn.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Trova_coordinate(ComboBox1.Text)
        TextBox25.Text = ComboBox1.Text
    End Sub

    Sub aggiorna_coordinate(par_UTENTE As String,
                        X_binocolo As Integer, Y_binocolo As Integer,
                        X_ricerca_odp As Integer, Y_ricerca_odp As Integer,
                        X_tasto_DX As Integer, Y_tasto_DX As Integer,
                        X_trasferimento_mag As Integer, Y_trasferimento_mag As Integer,
                        X_filtro_mag As Integer, Y_filtro_mag As Integer,
                        X_button_filtro As Integer, Y_button_filtro As Integer,
                        X_bozza_trasf As Integer, Y_bozza_trasf As Integer,
                        X_aggiungere As Integer, Y_aggiungere As Integer,
                        X_esci As Integer, Y_esci As Integer,
                             X_trasferimento_mag_oc As Integer, Y_trasferimento_mag_oc As Integer)



        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1
        Cmd_SAP.CommandText = "DELETE [Tirelli_40].[dbo].[Trasferimento_auto_ferretto] WHERE UTENTE ='" & par_UTENTE & "' "

        Cmd_SAP.ExecuteNonQuery()
        Cmd_SAP.CommandText = "INSERT INTO [Tirelli_40].[dbo].[Trasferimento_auto_ferretto] 
                        (UTENTE, X_binocolo, Y_binocolo, 
                         X_ricerca_odp, Y_ricerca_odp, 
                         X_tasto_DX, Y_tasto_DX, 
                         X_trasferimento_mag, Y_trasferimento_mag, 
                         X_filtro_mag, Y_filtro_mag, 
                         X_button_filtro, Y_button_filtro, 
                         X_bozza_trasf, Y_bozza_trasf, 
                         X_aggiungere, Y_aggiungere, 
                         X_esci, Y_esci,
X_trasferimento_mag_oc, Y_trasferimento_mag_oc) 
                        VALUES 
                        (@UTENTE, @X_binocolo, @Y_binocolo, 
                         @X_ricerca_odp, @Y_ricerca_odp, 
                         @X_tasto_DX, @Y_tasto_DX, 
                         @X_trasferimento_mag, @Y_trasferimento_mag, 
                         @X_filtro_mag, @Y_filtro_mag, 
                         @X_button_filtro, @Y_button_filtro, 
                         @X_bozza_trasf, @Y_bozza_trasf, 
                         @X_aggiungere, @Y_aggiungere, 
                         @X_esci, @Y_esci,
@X_trasferimento_mag_oc, @Y_trasferimento_mag_oc)"

        Cmd_SAP.Parameters.AddWithValue("@UTENTE", par_UTENTE)
        Cmd_SAP.Parameters.AddWithValue("@X_binocolo", X_binocolo)
        Cmd_SAP.Parameters.AddWithValue("@Y_binocolo", Y_binocolo)
        Cmd_SAP.Parameters.AddWithValue("@X_ricerca_odp", X_ricerca_odp)
        Cmd_SAP.Parameters.AddWithValue("@Y_ricerca_odp", Y_ricerca_odp)
        Cmd_SAP.Parameters.AddWithValue("@X_tasto_DX", X_tasto_DX)
        Cmd_SAP.Parameters.AddWithValue("@Y_tasto_DX", Y_tasto_DX)
        Cmd_SAP.Parameters.AddWithValue("@X_trasferimento_mag", X_trasferimento_mag)
        Cmd_SAP.Parameters.AddWithValue("@Y_trasferimento_mag", Y_trasferimento_mag)
        Cmd_SAP.Parameters.AddWithValue("@X_filtro_mag", X_filtro_mag)
        Cmd_SAP.Parameters.AddWithValue("@Y_filtro_mag", Y_filtro_mag)
        Cmd_SAP.Parameters.AddWithValue("@X_button_filtro", X_button_filtro)
        Cmd_SAP.Parameters.AddWithValue("@Y_button_filtro", Y_button_filtro)
        Cmd_SAP.Parameters.AddWithValue("@X_bozza_trasf", X_bozza_trasf)
        Cmd_SAP.Parameters.AddWithValue("@Y_bozza_trasf", Y_bozza_trasf)
        Cmd_SAP.Parameters.AddWithValue("@X_aggiungere", X_aggiungere)
        Cmd_SAP.Parameters.AddWithValue("@Y_aggiungere", Y_aggiungere)
        Cmd_SAP.Parameters.AddWithValue("@X_esci", X_esci)
        Cmd_SAP.Parameters.AddWithValue("@Y_esci", Y_esci)
        Cmd_SAP.Parameters.AddWithValue("@X_trasferimento_mag_oc", X_trasferimento_mag_oc)
        Cmd_SAP.Parameters.AddWithValue("@Y_trasferimento_mag_oc", Y_trasferimento_mag_oc)

        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()

    End Sub

    Sub Elimina_coordinate(par_UTENTE As String)



        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1
        Cmd_SAP.CommandText = "DELETE [Tirelli_40].[dbo].[Trasferimento_auto_ferretto] WHERE UTENTE ='" & par_UTENTE & "' "
        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        aggiorna_coordinate(TextBox25.Text,
                        TextBox7.Text, TextBox8.Text,
                        TextBox10.Text, TextBox9.Text,
                        TextBox12.Text, TextBox11.Text,
                        TextBox14.Text, TextBox13.Text,
                        TextBox16.Text, TextBox15.Text,
                        TextBox18.Text, TextBox17.Text,
                        TextBox20.Text, TextBox19.Text,
                        TextBox22.Text, TextBox21.Text,
                        TextBox24.Text, TextBox23.Text,
TextBox27.Text, TextBox26.Text)
        Inserimento_postazioni(ComboBox1)
        ComboBox1.Text = TextBox25.Text

        MsgBox("Coordinate aggiornatee con successo")
    End Sub


    Sub Inserimento_postazioni(par_combobox As ComboBox)
        par_combobox.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT [Utente]  
        FROM [Tirelli_40].[dbo].[Trasferimento_auto_ferretto]
group by [Utente]
"
        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            par_combobox.Items.Add(cmd_SAP_reader("utente"))

        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Elimina_coordinate(TextBox25.Text)
        Inserimento_postazioni(ComboBox1)

        MsgBox("Coordinate eliminate con successo")

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        SetCursorPos(TextBox12.Text, TextBox11.Text)
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        aggiorna_coordinate_KPI(TextBox34.Text, TextBox46.Text, TextBox29.Text, TextBox28.Text, TextBox31.Text, TextBox30.Text, TextBox33.Text, TextBox32.Text, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        '     Inserimento_postazioni(ComboBox1)
        '  ComboBox1.Text = TextBox25.Text

        MsgBox("Coordinate aggiornatee con successo")
    End Sub

    Sub aggiorna_coordinate_KPI(par_reparto As String, par_UTENTE As String,
                       X_1 As Integer, Y_1 As Integer,
                      X_2 As Integer, Y_2 As Integer,
                                X_3 As Integer, Y_3 As Integer,
                                X_4 As Integer, Y_4 As Integer,
                                X_5 As Integer, Y_5 As Integer,
                                X_6 As Integer, Y_6 As Integer,
                                X_7 As Integer, Y_7 As Integer,
                                X_8 As Integer, Y_8 As Integer,
                                X_9 As Integer, Y_9 As Integer,
                                X_10 As Integer, Y_10 As Integer)



        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1
        Cmd_SAP.CommandText = "DELETE [Tirelli_40].[dbo].[KPI_automatici] 
WHERE UTENTE ='" & par_UTENTE & "' and reparto='" & par_reparto & "'"

        Cmd_SAP.ExecuteNonQuery()
        Cmd_SAP.CommandText = "INSERT INTO [Tirelli_40].[dbo].[Kpi_automatici]
           ([Reparto]
           ,[Utente]
           ,[X_1]
,[Y_1]
           ,[X_2]
           ,[Y_2]
           ,[X_3]
           ,[Y_3]
           ,[X_4]
           ,[Y_4]
           ,[X_5]
           ,[Y_5]
           ,[X_6]
           ,[Y_6]
           ,[X_7]
           ,[Y_7]
           ,[X_8]
           ,[Y_8]
           ,[X_9]
           ,[Y_9]
           ,[X_10]
           ,[Y_10])
     VALUES
    ('" & par_reparto & "'
    ,'" & par_UTENTE & "'
    ," & X_1 & "
    ," & Y_1 & "
    ," & X_2 & "
    ," & Y_2 & "
    ," & X_3 & "
    ," & Y_3 & "
    ," & X_4 & "
    ," & Y_4 & "
    ," & X_5 & "
    ," & Y_5 & "
    ," & X_6 & "
    ," & Y_6 & "
    ," & X_7 & "
    ," & Y_7 & "
    ," & X_8 & "
    ," & Y_8 & "
    ," & X_9 & "
    ," & Y_9 & "
    ," & X_10 & "
    ," & Y_10 & ")"

        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        SetCursorPos(269, 990)
        SleepWithEscapeCheck(3000) ' Aspetta 3 secondi
        ClickMouse()


        SetCursorPos(425, 990)
        SleepWithEscapeCheck(3000) ' Aspetta 3 secondi
        ClickMouse()


    End Sub
End Class