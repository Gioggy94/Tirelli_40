Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib

Public Class Distinta_base_form
    Public riga As Integer
    Public itemcode_riga As String
    Public VALIDfor_riga As String
    Public itemtype_riga As String
    Public itemname_riga As String
    Public price_riga As Decimal
    Public DfltWH_riga As String
    Public contatore As Integer
    Public Stop_codice_non_presente As Integer


    Public quantità_itt1 As String
    Public prezzo_unitario_itt1 As String
    Public attrezzaggio_itt1 As String
    Public itemcode_itt1 As String

    Public visorder_riga As Integer
    Public childnum_riga As Integer

    Public visorder_riga_sopra As Integer
    Public childnum_riga_sopra As Integer
    Public itemtype_itt1_sopra As String
    Public descrizione_itt1_sopra As String
    Public magazzino_itt1_sopra As String
    Public quantità_itt1_sopra As Decimal
    Public prezzo_unitario_itt1_sopra As Decimal
    Public attrezzaggio_itt1_sopra As Decimal
    Public itemcode_itt1_sopra As String



    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If e.RowIndex >= 0 Then


            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Codice) Then


                'Try
                itemcode_riga = UCase(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice").Value)
                informazioni_articolo(itemcode_riga, DataGridView1, riga)
                'Catch ex As Exception
                '    MsgBox("C'è un errore nell'articolo riga")
                'End Try



            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(Quantità) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Attrezzaggio) Then

                Try



                    contatore = e.RowIndex

                    If InStr(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value, ",") > 1 Then
                        quantità_itt1 = LSet(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value, InStr(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value), InStr(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value), ",") - 1))
                    Else
                        quantità_itt1 = Replace(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value, ",", ".")
                    End If

                    If InStr(DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value, ",") > 1 Then


                        prezzo_unitario_itt1 = LSet(DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value, InStr(DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value), InStr(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value), ",") - 1))

                    Else
                        prezzo_unitario_itt1 = Replace(DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value, ",", ".")
                    End If

                    If InStr(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, ",") > 1 Then
                        attrezzaggio_itt1 = LSet(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, InStr(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value), InStr(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value), ",") - 1))

                    Else
                        attrezzaggio_itt1 = Replace(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, ",", ".")

                    End If


                    DataGridView1.Rows(riga).Cells(columnName:="Totale").Value = (DataGridView1.Rows(riga).Cells(columnName:="Quantità").Value + DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value) * DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value

                    Dim costoTotale As Decimal = 0

                    ' Scorre tutte le righe del DataGridView e somma i valori della colonna "Totale"
                    For Each row As DataGridViewRow In DataGridView1.Rows
                        If Not IsDBNull(row.Cells("Totale").Value) Then
                            costoTotale += Convert.ToDecimal(row.Cells("Totale").Value)
                        End If
                    Next

                    ' Imposta il costo totale nella TextBox
                    TextBox5.Text = costoTotale.ToString("C2") / TextBox2.Text ' Format currency

                    'Dim CNN As New SqlConnection
                    'CNN.ConnectionString = Homepage.sap_tirelli
                    'cnn.Open()

                    'Dim CMD_SAP As New SqlCommand
                    'Dim cmd_SAP_reader As SqlDataReader
                    'CMD_SAP.Connection = cnn

                    'CMD_SAP.CommandText = "SELECT (" & quantità_itt1 + attrezzaggio_itt1 & ")*" & prezzo_unitario_itt1 & " as 'Totale' "

                    'cmd_SAP_reader = CMD_SAP.ExecuteReader
                    'If cmd_SAP_reader.Read() = True Then

                    '    DataGridView1.Rows(riga).Cells(columnName:="Totale").Value = cmd_SAP_reader("Totale")


                    'End If
                    'cmd_SAP_reader.Close()
                    'cnn.Close()



                    'DataGridView1.Rows(riga).Cells(columnName:="Totale").Value = (quantità_itt1 + attrezzaggio_itt1) * prezzo_unitario_itt1
                Catch ex As Exception

                End Try


            End If
        End If

    End Sub

    Sub informazioni_articolo(par_codice_articolo As String, par_datagridview As DataGridView, par_riga As Integer)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "Select Case When T2.[VisResCode] Is null Then T0.objTYPE Else '290' end as 'objtype'
,coalesce(t0.u_codice_brb,'') as 'Codice_BRB'
,T0.[ItemName],

 Case When T2.[VisResCode] Is null Then case when T0.[DfltWH] is null then '01' else T0.[DfltWH] end else 'RIS' END as 'DfltWH',

T1.[Price], T0.VALIDFOR 
FROM OITM T0  INNER JOIN ITM1 T1 ON T0.[ItemCode] = T1.[ItemCode] 
left join orsc t2 on T2.[VisResCode]=t0.itemcode 
WHERE T0.[ItemCode] ='" & par_codice_articolo & "' AND  T1.[PriceList] =2"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            VALIDfor_riga = cmd_SAP_reader("VALIDFOR")
            itemtype_riga = cmd_SAP_reader("objTYPE")
            itemname_riga = cmd_SAP_reader("ItemName")
            price_riga = cmd_SAP_reader("Price")
            DfltWH_riga = cmd_SAP_reader("DfltWH")


            If VALIDfor_riga = "N" Then
                MsgBox("Articolo inattivo")
            Else
                par_datagridview.Rows(par_riga).Cells(columnName:="codice_BRB").Value = cmd_SAP_reader("Codice_BRB")
                par_datagridview.Rows(par_riga).Cells(columnName:="Descrizione").Value = itemname_riga
                par_datagridview.Rows(par_riga).Cells(columnName:="Itemtype").Value = itemtype_riga
                par_datagridview.Rows(par_riga).Cells(columnName:="magazzino_").Value = DfltWH_riga
                par_datagridview.Rows(par_riga).Cells(columnName:="Prezzo_unitario").Value = price_riga
                par_datagridview.Rows(par_riga).Cells(columnName:="Quantità").Value = 1
                par_datagridview.Rows(par_riga).Cells(columnName:="Attrezzaggio").Value = 0
                par_datagridview.Rows(par_riga).Cells(columnName:="Totale").Value = (par_datagridview.Rows(par_riga).Cells(columnName:="Quantità").Value + par_datagridview.Rows(par_riga).Cells(columnName:="Attrezzaggio").Value) * par_datagridview.Rows(par_riga).Cells(columnName:="Prezzo_unitario").Value
            End If
        Else
            MsgBox("Articolo non esistente")

        End If
        cmd_SAP_reader.Close()
        CNN.Close()


    End Sub

    Sub INSERT_INTO_OITT(par_codice_padre As String, par_quantità As String, par_descrizione As String, PAR_UTENTE_SAP As String, par_loginstanc As Integer)
        If Check_distinta_base_già_presente(par_codice_padre) = "N" Then


            par_quantità = Replace(par_quantità, ",", ".")
            par_descrizione = Replace(par_descrizione, "'", " ")
            Dim CNN6 As New SqlConnection
            CNN6.ConnectionString = Homepage.sap_tirelli
            CNN6.Open()

            Dim CMD_SAP_5 As New SqlCommand

            CMD_SAP_5.Connection = CNN6
            CMD_SAP_5.CommandText = "INSERT INTO OITT (OITT.Code, OITT.TreeType, OITT.PriceList, OITT.Qauntity, OITT.CreateDate, OITT.UpdateDate, OITT.Transfered, OITT.DispCurr, OITT.ToWH, OITT.Object, OITT.LogInstac, OITT.UserSign2, OITT.OcrCode, OITT.HideComp, OITT.OcrCode2, OITT.OcrCode3, OITT.OcrCode4, OITT.OcrCode5, OITT.UpdateTime, OITT.Project, OITT.PlAvgSize, OITT.Name, OITT.CreateTS,OITT.UPDATETS,OITT.UserSign)
                                  VALUES('" & par_codice_padre & "','P','2','" & par_quantità & "',CONVERT(date, GETDATE()),CONVERT(date, GETDATE()),'N','','01','66'," & par_loginstanc & ",'" & PAR_UTENTE_SAP & "','','N','','','','',cast(CONCAT(DATEPART(HOUR, GETDATE()),DATEPART(MINUTE, GETDATE())) as integer),'','1','" & par_descrizione & "',CONCAT(DATEPART(HOUR, GETDATE()),DATEPART(MINUTE, GETDATE()),DATEPART(second, GETDATE())),CONCAT(DATEPART(HOUR, GETDATE()),DATEPART(MINUTE, GETDATE()),DATEPART(second, GETDATE())),'" & PAR_UTENTE_SAP & "')"
            CMD_SAP_5.ExecuteNonQuery()


            CNN6.Close()
        End If
    End Sub

    Sub INSERT_INTO_AITT(par_codice_padre As String, par_quantità As String, par_descrizione As String, PAR_UTENTE_SAP As String, par_loginstanc As Integer)



        par_quantità = Replace(par_quantità, ",", ".")
        par_descrizione = Replace(par_descrizione, "'", " ")
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = CNN6
        CMD_SAP_5.CommandText = "INSERT INTO AITT (AITT.Code, AITT.TreeType, AITT.PriceList, AITT.Qauntity, AITT.CreateDate, AITT.UpdateDate, AITT.Transfered, AITT.DispCurr, AITT.ToWH, AITT.Object, AITT.LogInstac, AITT.UserSign2, AITT.OcrCode, AITT.HideComp, AITT.OcrCode2, AITT.OcrCode3, AITT.OcrCode4, AITT.OcrCode5, AITT.UpdateTime, AITT.Project, AITT.PlAvgSize, AITT.Name, AITT.CreateTS,AITT.UPDATETS,AITT.UserSign)
                                  VALUES('" & par_codice_padre & "','P','2','" & par_quantità & "',CONVERT(date, GETDATE()),CONVERT(date, GETDATE()),'N','','01','66'," & par_loginstanc & ",'" & PAR_UTENTE_SAP & "','','N','','','','',cast(CONCAT(DATEPART(HOUR, GETDATE()),DATEPART(MINUTE, GETDATE())) as integer),'','1','" & par_descrizione & "',CONCAT(DATEPART(HOUR, GETDATE()),DATEPART(MINUTE, GETDATE()),DATEPART(second, GETDATE())),CONCAT(DATEPART(HOUR, GETDATE()),DATEPART(MINUTE, GETDATE()),DATEPART(second, GETDATE())),'" & PAR_UTENTE_SAP & "')"
        CMD_SAP_5.ExecuteNonQuery()


        CNN6.Close()

    End Sub

    Sub delete_itt1(par_codice_padre As String, par_posizione As Integer, par_codice_figlio As String)


        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        cnn6.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = cnn6
        CMD_SAP_5.CommandText = "DELETE ITT1 WHERE CODE='" & par_codice_figlio & "' AND FATHER = '" & par_codice_padre & "' AND (CHILDNUM=" & par_posizione & " OR VISORDER = " & par_posizione & ")"
        CMD_SAP_5.ExecuteNonQuery()


        cnn6.Close()

    End Sub

    Function Check_distinta_base_già_presente(Par_codice_articolo As String)
        Dim CNN As New SqlConnection
        Dim risposta As String = "N"
        cnn.Close()
        cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "SELECT coalesce(T0.code,'') as 'DB'
from OITt T0 
where t0.code='" & Par_codice_articolo & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            If cmd_SAP_reader("DB") = "" Then
                risposta = "N"
            Else
                risposta = "Y"
            End If


        Else
            risposta = "N"
        End If
        cmd_SAP_reader.Close()
        cnn.Close()

        Return risposta

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        If Stop_codice_non_presente = 1 Then
            MsgBox("Creare il codice " & " " & TextBox1.Text & " prima di crearne la distinta base")
        Else

            If TextBox2.Text >= 1 Then
                delete_oitt(TextBox1.Text)
                delete_itt1(TextBox1.Text)
                Dim loginstanc As Integer = Trova_ultimo_loginstanc_distinta(TextBox1.Text)
                INSERT_INTO_OITT(TextBox1.Text, TextBox2.Text, TextBox3.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, loginstanc)
                INSERT_INTO_AITT(TextBox1.Text, TextBox2.Text, TextBox3.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, loginstanc)
                contatore = 0
                Do While contatore <= DataGridView1.Rows.Count - 2


                    INSERT_INTO_ITT1(TextBox1.Text, DataGridView1.Rows(contatore).Cells(columnName:="Codice").Value, DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value, DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value, DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, contatore, DataGridView1.Rows(contatore).Cells(columnName:="Magazzino_").Value, DataGridView1.Rows(contatore).Cells(columnName:="Descrizione").Value, loginstanc, DataGridView1.Rows(contatore).Cells(columnName:="Importazione").Value)
                    INSERT_INTO_ATT1(TextBox1.Text, DataGridView1.Rows(contatore).Cells(columnName:="Codice").Value, DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value, DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value, DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, contatore, DataGridView1.Rows(contatore).Cells(columnName:="Magazzino_").Value, DataGridView1.Rows(contatore).Cells(columnName:="Descrizione").Value, loginstanc, DataGridView1.Rows(contatore).Cells(columnName:="Importazione").Value)
                    contatore = contatore + 1

                Loop
                MEttere_db_produzione_oitm(TextBox1.Text)
                MsgBox("Distinta base creata con successo")



                Me.Close()
            Else
                MsgBox("Selezionare una quantità > 0")

            End If
        End If





    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If Not IsNumeric(TextBox2.Text) Then
            MsgBox("Selezionare un valore numerico valido")
            TextBox2.Text = 1
        End If

    End Sub

    Sub INSERT_INTO_ITT1(par_codice_padre As String, par_codice_articolo As String, par_quantità As String, par_prezzo_unitario As String, par_attrezzaggio As String, par_riga As Integer, par_magazzino As String, par_descrizione_articolo As String, PAR_LOGINSTANC As Integer, PAR_PDM As String)

        par_descrizione_articolo = Replace(par_descrizione_articolo, "'", " ")
        par_quantità = Replace(par_quantità, ",", ".")
        itemcode_itt1 = UCase(par_codice_articolo)



        If InStr(par_quantità, ",") > 1 Then
            quantità_itt1 = LSet(par_quantità, InStr(par_quantità, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value), InStr(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value), ",") - 1))
        Else
            quantità_itt1 = Replace(par_quantità, ",", ".")
        End If

        If InStr(par_prezzo_unitario, ",") > 1 Then


            prezzo_unitario_itt1 = LSet(par_prezzo_unitario, InStr(par_prezzo_unitario, ",") - 1) & "." & StrReverse(LSet(StrReverse(par_prezzo_unitario), InStr(StrReverse(par_prezzo_unitario), ",") - 1))

        Else
            prezzo_unitario_itt1 = Replace(par_prezzo_unitario, ",", ".")
        End If

        If InStr(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, ",") > 1 Then
            attrezzaggio_itt1 = LSet(par_attrezzaggio, InStr(par_attrezzaggio, ",") - 1) & "." & StrReverse(LSet(StrReverse(par_attrezzaggio), InStr(StrReverse(par_attrezzaggio), ",") - 1))

        Else
            attrezzaggio_itt1 = Replace(par_attrezzaggio, ",", ".")

        End If
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand

        Dim TIPO_PRELIEVO As String

        If Magazzino.OttieniDettagliAnagrafica(itemcode_itt1).Gruppo = "Lastre" Then
            TIPO_PRELIEVO = "M"
        Else
            TIPO_PRELIEVO = "B"
        End If

        CMD_SAP_5.Connection = CNN6
        CMD_SAP_5.CommandText = "INSERT INTO ITT1 (ITT1.Father, ITT1.ChildNum, ITT1.VisOrder, ITT1.Code, ITT1.Quantity, ITT1.Warehouse, ITT1.Price, ITT1.Currency, ITT1.PriceList, ITT1.OrigPrice, ITT1.OrigCurr, ITT1.IssueMthd, ITT1.Uom, ITT1.Comment, ITT1.LogInstanc, ITT1.Object, ITT1.OcrCode, ITT1.OcrCode2, ITT1.OcrCode3, ITT1.OcrCode4, ITT1.OcrCode5, ITT1.PrncpInput, ITT1.Project, ITT1.Type, ITT1.WipActCode, ITT1.AddQuantit, ITT1.LineText,  ITT1.ItemName, ITT1.U_TEMPOME, ITT1.U_UBIMAG, ITT1.U_PRG_AZS_PhanFat, ITT1.U_PRG_TIR_LPN, ITT1.U_Tempo_preassemblaggio, ITT1.U_Tempo_preassemb, ITT1.U_PRG_TIR_Importazione)
VALUES ('" & par_codice_padre & "','" & par_riga & "','" & par_riga & "','" & itemcode_itt1 & "'," & quantità_itt1 & ",'" & par_magazzino & "'," & prezzo_unitario_itt1 & ",'EUR','2'," & prezzo_unitario_itt1 & " ,'EUR','" & TIPO_PRELIEVO & "','',''," & PAR_LOGINSTANC & ",'66','','','','','','N','',case when substring('" & itemcode_itt1 & "',1,1)='R' then 290 else 4 end,''," & attrezzaggio_itt1 & ",'',substring('" & par_descrizione_articolo & "',1,200),'','','','N','','','" & PAR_PDM & "')"
        CMD_SAP_5.ExecuteNonQuery()




        CNN6.Close()


    End Sub

    Sub INSERT_INTO_ATT1(par_codice_padre As String, par_codice_articolo As String, par_quantità As String, par_prezzo_unitario As String, par_attrezzaggio As String, par_riga As Integer, par_magazzino As String, par_descrizione_articolo As String, PAR_LOGINSTANC As Integer, par_importazione As String)

        par_descrizione_articolo = Replace(par_descrizione_articolo, "'", " ")
        itemcode_itt1 = UCase(par_codice_articolo)
        par_quantità = Replace(par_quantità, ",", ".")


        If InStr(par_quantità, ",") > 1 Then
            quantità_itt1 = LSet(par_quantità, InStr(par_quantità, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value), InStr(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value), ",") - 1))
        Else
            quantità_itt1 = Replace(par_quantità, ",", ".")
        End If

        If InStr(par_prezzo_unitario, ",") > 1 Then


            prezzo_unitario_itt1 = LSet(par_prezzo_unitario, InStr(par_prezzo_unitario, ",") - 1) & "." & StrReverse(LSet(StrReverse(par_prezzo_unitario), InStr(StrReverse(par_prezzo_unitario), ",") - 1))

        Else
            prezzo_unitario_itt1 = Replace(par_prezzo_unitario, ",", ".")
        End If

        If InStr(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, ",") > 1 Then
            attrezzaggio_itt1 = LSet(par_attrezzaggio, InStr(par_attrezzaggio, ",") - 1) & "." & StrReverse(LSet(StrReverse(par_attrezzaggio), InStr(StrReverse(par_attrezzaggio), ",") - 1))

        Else
            attrezzaggio_itt1 = Replace(par_attrezzaggio, ",", ".")

        End If
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = CNN6
        CMD_SAP_5.CommandText = "INSERT INTO ATT1 (ATT1.Father, ATT1.ChildNum, ATT1.VisOrder, ATT1.Code, ATT1.Quantity, ATT1.Warehouse, ATT1.Price, ATT1.Currency, ATT1.PriceList, ATT1.OrigPrice, ATT1.OrigCurr, ATT1.IssueMthd, ATT1.Uom, ATT1.Comment, ATT1.LogInstanc, ATT1.Object, ATT1.OcrCode, ATT1.OcrCode2, ATT1.OcrCode3, ATT1.OcrCode4, ATT1.OcrCode5, ATT1.PrncpInput, ATT1.Project, ATT1.Type, ATT1.WipActCode, ATT1.AddQuantit, ATT1.LineText,  ATT1.ItemName, ATT1.U_TEMPOME, ATT1.U_UBIMAG, ATT1.U_PRG_AZS_PhanFat, ATT1.U_PRG_TIR_LPN, ATT1.U_Tempo_preassemblaggio, ATT1.U_Tempo_preassemb, ATT1.U_PRG_TIR_Importazione)
VALUES ('" & par_codice_padre & "','" & par_riga & "','" & par_riga & "','" & itemcode_itt1 & "'," & quantità_itt1 & ",'" & par_magazzino & "'," & prezzo_unitario_itt1 & ",'EUR','2'," & prezzo_unitario_itt1 & " ,'EUR','B','',''," & PAR_LOGINSTANC & ",'66','','','','','','N','',case when substring('" & itemcode_itt1 & "',1,1)='R' then 290 else 4 end,''," & attrezzaggio_itt1 & ",'','" & par_descrizione_articolo & "','','','','N','','','" & par_importazione & "')"
        CMD_SAP_5.ExecuteNonQuery()




        CNN6.Close()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Close()

    End Sub

    Sub Check_se_distinta_base_già_presente()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "SELECT T0.itemcode,t0.itemname, case when t1.code is null then '' else t1.code end as 'code', case when t1.u_sottocodice is null then 0 else t1.u_sottocodice end as 'u_sottocodice' 
, t1.qauntity
from OITM T0 left join oitt t1 on t0.itemcode=t1.code where t0.itemcode='" & TextBox1.Text & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            If cmd_SAP_reader("code") = "" Then
                Stop_codice_non_presente = 0
                TextBox2.Text = 1
                TextBox3.Text = cmd_SAP_reader("itemname")
                TextBox4.Text = cmd_SAP_reader("u_sottocodice")
                Button1.Text = "Aggiungere"

            Else
                Stop_codice_non_presente = 0
                TextBox3.Text = cmd_SAP_reader("itemname")
                TextBox4.Text = cmd_SAP_reader("u_sottocodice")
                TextBox2.Text = cmd_SAP_reader("qauntity")
                Riempi_distinta_base(DataGridView1, TextBox1.Text, TextBox5, TextBox2)
                ' Button1.Text = "Aggiornare"
                Button1.Text = "Aggiungere"
            End If

        Else
            Stop_codice_non_presente = 1
            TextBox2.Text = 1

        End If
        cmd_SAP_reader.Close()
        cnn.Close()


    End Sub
    Public Function Trova_ultimo_loginstanc_distinta(par_Codice_sap As String)

        Dim loginstanc As Integer = 0
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "SELECT coalescE(max(coalesce(T0.[LogInstac],0)),0) as 'LogInstac'
from AITT t0
where t0.code='" & par_Codice_sap & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            loginstanc = cmd_SAP_reader("LogInstac") + 1


        Else
            loginstanc = 0

        End If
        cmd_SAP_reader.Close()
        CNN.Close()

        Return loginstanc
    End Function



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Check_se_distinta_base_già_presente()
    End Sub

    Sub MEttere_db_produzione_oitm(par_codice_padre As String)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        cnn6.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = cnn6
        CMD_SAP_5.CommandText = "UPDATE T0 SET T0.TREETYPE ='P' FROM OITM T0 WHERE T0.[ItemCode] ='" & par_codice_padre & "'"
        CMD_SAP_5.ExecuteNonQuery()


        cnn6.Close()

    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Button1.Text = "Aggiornare"
        If e.ColumnIndex < 0 Then
            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            riga = e.RowIndex

        Else
            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End If


        If e.RowIndex = 0 Then
            Button3.Visible = False
            Button4.Visible = True
        ElseIf e.RowIndex > 0 And e.RowIndex < DataGridView1.RowCount - 2 Then
            Button3.Visible = True
            Button4.Visible = True
        ElseIf e.RowIndex = DataGridView1.RowCount - 2 Then
            Button3.Visible = True
            Button4.Visible = False
        End If


    End Sub



    Sub Riempi_distinta_base(par_datagridview As DataGridView, par_codice_padre As String, par_textbox_prezzo As TextBox, par_textbox_quantita As TextBox)
        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn1

        CMD_SAP.CommandText = "SELECT T0.[VisOrder], T0.[ChildNum], T0.[Type],
COALESCE(T0.[Code],'') AS 'CODE', 
coalesce(t1.u_codice_brb,'') as 'Codice_BRB',
COALESCE(T0.[ItemName],'') AS 'ITEMNAME'
, COALESCE(T0.[Warehouse],'') AS 'WAREHOUSE', T0.[Quantity], CASE WHEN T0.[AddQuantit] IS NULL THEN 0 ELSE T0.[AddQuantit] END AS 'AddQuantit' 
, COALESCE(T0.[Price],0) AS 'PRICE' 
,coalesce(t0.u_prg_tir_importazione,'') as 'importazione'
, coalesce(sum(t2.onhand),0) as 'Mag_tot'
, coalesce(sum(case when t2.whscode='01' THEN t2.onhand ELSE 0 END),0) as 'Mag_01'
, coalesce(sum(case when t2.whscode='FERRETTO' THEN t2.onhand ELSE 0 END),0) as 'Mag_FER'
, coalesce(sum(case when t2.whscode='03' THEN t2.onhand ELSE 0 END),0) as 'Mag_03'
, coalesce(sum(case when t2.whscode='SCA' THEN t2.onhand ELSE 0 END),0) as 'Mag_SCA'
, coalesce(sum(case when t2.whscode='WIP' THEN t2.onhand ELSE 0 END),0) as 'Mag_WIP'



FROM ITT1 T0 left join oitm t1 on t0.code=t1.itemcode
left join oitw t2 on t2.itemcode=t1.itemcode 

WHERE T0.[Father] ='" & par_codice_padre & "' 
group by T0.[VisOrder], T0.[ChildNum], T0.[Type],T0.[Code],t1.u_codice_brb,T0.[ItemName],T0.[Warehouse],T0.[Price],t0.Quantity,T0.[AddQuantit],t0.u_prg_tir_importazione
ORDER BY T0.VISORDER"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader("visorder"), cmd_SAP_reader("Childnum"), cmd_SAP_reader("Type"), cmd_SAP_reader("Code"), cmd_SAP_reader("Codice_BRB"), cmd_SAP_reader("ItemName"), cmd_SAP_reader("Warehouse"), cmd_SAP_reader("Quantity"), cmd_SAP_reader("AddQuantit"), cmd_SAP_reader("Price"), (cmd_SAP_reader("Quantity") + cmd_SAP_reader("AddQuantit")) * cmd_SAP_reader("Price"), cmd_SAP_reader("Mag_tot"), cmd_SAP_reader("Mag_01"), cmd_SAP_reader("Mag_FER"), cmd_SAP_reader("Mag_03"), cmd_SAP_reader("Mag_SCA"), cmd_SAP_reader("Mag_WIP"), cmd_SAP_reader("IMPORTAZIONE"))



        Loop


        cmd_SAP_reader.Close()
        Cnn1.Close()

        par_datagridview.ClearSelection()

        Dim costoTotale As Decimal = 0

        ' Scorre tutte le righe del DataGridView e somma i valori della colonna "Totale"
        For Each row As DataGridViewRow In par_datagridview.Rows
            If Not IsDBNull(row.Cells("Totale").Value) Then
                costoTotale += Convert.ToDecimal(row.Cells("Totale").Value)
            End If
        Next

        ' Imposta il costo totale nella TextBox
        Try
            par_textbox_prezzo.Text = costoTotale.ToString("C2") / par_textbox_quantita.Text ' Format currency
        Catch ex As Exception
            par_textbox_prezzo.Text = 0
        End Try

    End Sub

    Sub compila_odp()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click



        If riga = 0 Then
            Button3.Visible = False
        End If
        SpostaRigaSu(DataGridView1, riga)
        riga -= 1


    End Sub

    Sub SpostaRigaSu(par_datagridview As DataGridView, riga As Integer)
        ' Controlla se la riga corrente è valida e non è la prima riga
        If riga <= 0 Or riga >= par_datagridview.Rows.Count Then Exit Sub

        ' Crea array per memorizzare i valori della riga corrente e della riga precedente
        Dim valoriRigaCorrente As Object() = New Object(par_datagridview.Columns.Count - 1) {}
        Dim valoriRigaPrecedente As Object() = New Object(par_datagridview.Columns.Count - 1) {}

        ' Memorizza i valori della riga corrente
        For i As Integer = 0 To par_datagridview.Columns.Count - 1
            valoriRigaCorrente(i) = par_datagridview.Rows(riga).Cells(i).Value
        Next

        ' Memorizza i valori della riga precedente
        For i As Integer = 0 To par_datagridview.Columns.Count - 1
            valoriRigaPrecedente(i) = par_datagridview.Rows(riga - 1).Cells(i).Value
        Next

        ' Scambia i valori tra la riga corrente e la riga precedente
        For i As Integer = 0 To par_datagridview.Columns.Count - 1
            par_datagridview.Rows(riga).Cells(i).Value = valoriRigaPrecedente(i)
            par_datagridview.Rows(riga - 1).Cells(i).Value = valoriRigaCorrente(i)
        Next

        ' Aggiorna la selezione visiva
        par_datagridview.Rows(riga - 1).Selected = True
        par_datagridview.Rows(riga).Selected = False
    End Sub



    Public Sub delete_itt1(par_codice_articolo As String)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        cnn6.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = cnn6

        CMD_SAP_5.CommandText = "
delete itt1 where itt1.father='" & par_codice_articolo & "'"
        CMD_SAP_5.ExecuteNonQuery()

        cnn6.Close()

    End Sub

    Public Sub delete_oitt(par_codice_articolo)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        cnn6.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = cnn6

        CMD_SAP_5.CommandText = "
delete oitt where oitt.code='" & par_codice_articolo & "'"
        CMD_SAP_5.ExecuteNonQuery()

        cnn6.Close()

    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim par_datagridview As DataGridView = DataGridView1




        'Aggiorna_numero_combinazione(DataGridView1, par_id_combinazione As Integer)
        SpostaRigaGiù(DataGridView1, riga)
        riga += 1
        If riga = par_datagridview.RowCount - 2 Then
            Button4.Visible = False
            Button3.Visible = True
        End If


    End Sub



    Public Sub SpostaRigaGiù(par_datagridview As DataGridView, riga As Integer)
        ' Controlla se la riga corrente è valida e non è l'ultima riga
        If riga < 0 Or riga >= par_datagridview.Rows.Count - 1 Then Exit Sub

        ' Crea array per memorizzare i valori della riga corrente e della riga successiva
        Dim valoriRigaCorrente As Object() = New Object(par_datagridview.Columns.Count - 1) {}
        Dim valoriRigaSuccessiva As Object() = New Object(par_datagridview.Columns.Count - 1) {}

        ' Memorizza i valori della riga corrente
        For i As Integer = 0 To par_datagridview.Columns.Count - 1
            valoriRigaCorrente(i) = par_datagridview.Rows(riga).Cells(i).Value
        Next

        ' Memorizza i valori della riga successiva
        For i As Integer = 0 To par_datagridview.Columns.Count - 1
            valoriRigaSuccessiva(i) = par_datagridview.Rows(riga + 1).Cells(i).Value
        Next

        ' Scambia i valori tra la riga corrente e la riga successiva
        For i As Integer = 0 To par_datagridview.Columns.Count - 1
            par_datagridview.Rows(riga).Cells(i).Value = valoriRigaSuccessiva(i)
            par_datagridview.Rows(riga + 1).Cells(i).Value = valoriRigaCorrente(i)
        Next

        ' Aggiorna la selezione visiva
        par_datagridview.Rows(riga + 1).Selected = True
        par_datagridview.Rows(riga).Selected = False
    End Sub



    Private Sub DataGridView1_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        If e.RowIndex >= 0 Then
            riga = e.RowIndex

        End If
    End Sub





    Sub Check_se_distinta_base_sottocodice()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "SELECT T0.itemcode,t0.itemname, case when t1.code is null then '' else t1.code end as 'code' from OITM T0 left join oitt t1 on t0.itemcode=t1.code where t0.itemcode='" & TextBox1.Text & "_'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            If cmd_SAP_reader("code") = "" Then
                Stop_codice_non_presente = 0
                TextBox3.Text = cmd_SAP_reader("itemname")

                Button1.Text = "Aggiungere"

            Else
                Stop_codice_non_presente = 0
                TextBox3.Text = cmd_SAP_reader("itemname")
                Riempi_distinta_base(DataGridView1, TextBox1.Text, TextBox5, TextBox2)
                ' Button1.Text = "Aggiornare"
                Button1.Text = "Aggiungere"
            End If

        Else
            Stop_codice_non_presente = 1

        End If
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Sub Check_se_sottocodice_già_presente_OITM()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "SELECT T0.itemcode,t0.itemname, case when t1.code is null then '' else t1.code end as 'code', case when t1.u_sottocodice is null then 0 else t1.u_sottocodice end as 'u_sottocodice' from OITM T0 left join oitt t1 on t0.itemcode=t1.code where t0.itemcode='" & TextBox1.Text & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            If cmd_SAP_reader("code") = "" Then
                Stop_codice_non_presente = 0
                TextBox3.Text = cmd_SAP_reader("itemname")
                TextBox4.Text = cmd_SAP_reader("u_sottocodice")
                Button1.Text = "Aggiungere"

            Else
                Stop_codice_non_presente = 0
                TextBox3.Text = cmd_SAP_reader("itemname")
                TextBox4.Text = cmd_SAP_reader("u_sottocodice")
                Riempi_distinta_base(DataGridView1, TextBox1.Text, TextBox5, TextBox2)
                ' Button1.Text = "Aggiornare"
                Button1.Text = "Aggiungere"
            End If

        Else
            Stop_codice_non_presente = 1

        End If
        cmd_SAP_reader.Close()
        cnn.Close()


    End Sub





    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Albero.Show()
        Albero.commessa = TextBox1.Text
        Albero.TextBox1.Text = Albero.commessa
        Albero.inizializza_albero(TextBox1.Text)
    End Sub

    Private Sub CancellareRigaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CancellareRigaToolStripMenuItem.Click



        DataGridView1.Rows.RemoveAt(riga)

    End Sub

    Private Sub DatiAnagraficiArticoloToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DatiAnagraficiArticoloToolStripMenuItem.Click

        Magazzino.Codice_SAP = DataGridView1.Rows(riga).Cells(columnName:="Codice").Value
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

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Distinta_base_form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DistintaBaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DistintaBaseToolStripMenuItem.Click

        Dim new_form_distinta_form As New Distinta_base_form

        ' Imposta il valore della TextBox
        new_form_distinta_form.TextBox1.Text = DataGridView1.Rows(riga).Cells("Codice").Value.ToString()

        ' Calcola 1 cm in pixel (conversione da cm a pixel, 96 dpi)
        Dim oneCmInPixels As Integer = CInt(1 / 2.54 * 96) ' 1 cm ≈ 37,8 pixel a 96 dpi

        ' Imposta la posizione della nuova form rispetto alla posizione dell'attuale form
        new_form_distinta_form.StartPosition = FormStartPosition.Manual
        new_form_distinta_form.Location = New Point(Me.Location.X + oneCmInPixels, Me.Location.Y + oneCmInPixels)

        ' Mostra la form
        new_form_distinta_form.Show()


    End Sub

    Private Sub DataGridView1_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDown
        ' Controlla se il tasto destro del mouse è stato premuto
        If e.Button = MouseButtons.Right Then
            ' Ottiene l'indice della riga in base alla posizione del mouse
            Dim hit As DataGridView.HitTestInfo = DataGridView1.HitTest(e.X, e.Y)

            ' Verifica se una cella valida è stata cliccata
            If hit.RowIndex >= 0 Then
                ' Seleziona la riga corrispondente
                DataGridView1.ClearSelection()
                DataGridView1.Rows(hit.RowIndex).Selected = True
            End If
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        ' Verifica se c'è una riga selezionata
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Ottieni l'indice della prima riga selezionata
            Dim rowIndex As Integer = DataGridView1.SelectedRows(0).Index

            ' Fai qualcosa con rowIndex, ad esempio visualizzalo o utilizzalo per ulteriori operazioni
            riga = rowIndex.ToString()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Il nodo è un nodo figlio: apri la parte di magazzino
        Magazzino.Codice_SAP = TextBox1.Text

        ' Ripristina la finestra se è minimizzata e portala in primo piano
        If Magazzino.WindowState = FormWindowState.Minimized Then
            Magazzino.WindowState = FormWindowState.Normal
        End If
        Magazzino.BringToFront()
        Magazzino.Activate()
        Magazzino.Show()

        ' Imposta il codice SAP e aggiorna i dettagli
        Magazzino.TextBox2.Text = Magazzino.Codice_SAP
        Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Dim par_datagridview As DataGridView = DataGridView1
        ' Creare un'applicazione Excel
        Dim excelApp As New Excel.Application
        excelApp.Visible = True ' Mostrare Excel all'utente

        ' Creare un nuovo foglio di lavoro
        Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add
        Dim excelWorksheet As Excel.Worksheet = CType(excelWorkbook.Worksheets(1), Excel.Worksheet)

        ' Aggiungere intestazioni alla prima riga del foglio di lavoro (facoltativo)
        For col As Integer = 1 To par_datagridview.Columns.Count
            excelWorksheet.Cells(1, col) = par_datagridview.Columns(col - 1).HeaderText
        Next

        ' Aggiungere dati dalla DataGridView al foglio di lavoro
        For row As Integer = 0 To par_datagridview.Rows.Count - 1
            For col As Integer = 0 To par_datagridview.Columns.Count - 1
                ' Imposta il formato della cella come testo
                excelWorksheet.Cells(row + 2, col + 1).NumberFormat = "@"
                ' Inserisce il valore nella cella
                excelWorksheet.Cells(row + 2, col + 1) = par_datagridview.Rows(row).Cells(col).Value
            Next
        Next

        ' Salvare il file Excel
        Dim saveFileDialog As New SaveFileDialog
        saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            excelWorkbook.SaveAs(saveFileDialog.FileName)
            MessageBox.Show("Esportazione completata con successo!", "Successo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        ' Chiudere Excel
        excelApp.Quit()
        ReleaseComObject(excelApp)

    End Sub

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class