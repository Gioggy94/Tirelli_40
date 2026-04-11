Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Windows.Annotations
Imports foxit.pdf

Public Class PDM_SAP
    Public Excel As Excel.Application
    Public WB_Excel As Excel.Workbook
    Sub leggi_file_anagrafica()
        ' Percorso del file
        Dim percorsoFile_anagrafica As String = Homepage.percorso_PDM_BRB & "\anagrafica.txt"

        ' Verifica se il file esiste
        If File.Exists(percorsoFile_anagrafica) Then
            ' Leggi tutte le righe del file
            Dim righe As String() = File.ReadAllLines(percorsoFile_anagrafica)

            ' Verifica se ci sono righe nel file
            If righe.Length > 0 Then
                For Each riga As String In righe
                    ' Trova gli indici dei separatori "|"
                    Dim indice1Pipe As Integer = riga.IndexOf("|")
                    Dim indice2Pipe As Integer = riga.IndexOf("|", indice1Pipe + 1)
                    Dim indice3Pipe As Integer = riga.IndexOf("|", indice2Pipe + 1)
                    Dim indice4Pipe As Integer = riga.IndexOf("|", indice3Pipe + 1)
                    Dim indice5Pipe As Integer = riga.IndexOf("|", indice4Pipe + 1)
                    Dim indice6Pipe As Integer = riga.IndexOf("|", indice5Pipe + 1)
                    Dim indice7Pipe As Integer = riga.IndexOf("|", indice6Pipe + 1)
                    Dim indice8Pipe As Integer = riga.IndexOf("|", indice7Pipe + 1)
                    Dim indice9Pipe As Integer = riga.IndexOf("|", indice8Pipe + 1)
                    Dim indice10Pipe As Integer = riga.IndexOf("|", indice9Pipe + 1)
                    Dim indice11Pipe As Integer = riga.IndexOf("|", indice10Pipe + 1)
                    Dim indice12Pipe As Integer = riga.IndexOf("|", indice11Pipe + 1)
                    Dim indice13Pipe As Integer = riga.IndexOf("|", indice12Pipe + 1)
                    Dim indice14Pipe As Integer = riga.IndexOf("|", indice13Pipe + 1)


                    If indice1Pipe >= 0 Then

                        Dim codice_BRB As String = riga.Substring(0, indice1Pipe)


                        Dim descrizione As String = riga.Substring(indice1Pipe + 1, indice2Pipe - indice1Pipe - 1)


                        Dim descrizione_supp As String = riga.Substring(indice3Pipe + 1, indice4Pipe - indice3Pipe - 1)

                        Dim udm As String = riga.Substring(indice4Pipe + 1, indice5Pipe - indice4Pipe - 1)

                        Dim ubicazione As String = riga.Substring(indice5Pipe + 1, indice6Pipe - indice5Pipe - 1)


                        Dim codice_barre As String = riga.Substring(indice6Pipe + 1, indice7Pipe - indice6Pipe - 1)



                        Dim revisione = riga.Substring(indice7Pipe + 1, indice8Pipe - indice7Pipe - 1)



                        ' Dim revisione = riga.Substring(indice8Pipe + 1, indice9Pipe - indice8Pipe - 1)



                        Dim creato_da = riga.Substring(indice9Pipe + 1, indice10Pipe - indice9Pipe - 1)


                        Dim modificato_da = riga.Substring(indice10Pipe + 1, indice11Pipe - indice10Pipe - 1)


                        Dim materiale = riga.Substring(indice11Pipe + 1, indice12Pipe - indice11Pipe - 1)



                        Dim trattamento = riga.Substring(indice12Pipe + 1, indice13Pipe - indice12Pipe - 1)


                        Dim disegno = riga.Substring(indice13Pipe + 1)

                        Dim fornitore As String = ""






                        importazione_PDM_SAP_ANAGRAFICA(codice_BRB, Replace(Replace(descrizione, ",", "."), "'", " "), Replace(Replace(descrizione_supp, ",", "."), "'", " "), revisione, 100, creato_da, modificato_da, Replace(udm, "'", " "), Replace(materiale, "'", " "), Replace(trattamento, "'", " "), Replace(ubicazione, "'", " "), Replace(codice_barre, "'", " "), Replace(disegno, "'", " "), fornitore)
                        ' Puoi eseguire ulteriori operazioni con le sottostringhe qui


                    End If
                Next
            Else
                Console.WriteLine("Il file è vuoto.")
            End If
        Else
            Console.WriteLine("Il file non esiste.")
        End If

        Console.ReadLine()
    End Sub

    Sub leggi_file_distinta()
        ' Percorso del file
        Dim percorsoFile_legami As String = Homepage.percorso_PDM_BRB & "\legami.txt"

        ' Verifica se il file esiste
        If File.Exists(percorsoFile_legami) Then
            ' Leggi tutte le righe del file
            Dim righe As String() = File.ReadAllLines(percorsoFile_legami)

            ' Verifica se ci sono righe nel file
            If righe.Length > 0 Then
                For Each riga As String In righe
                    ' Trova gli indici dei separatori "|"
                    Dim indice1Pipe As Integer = riga.IndexOf("|")
                    Dim indice2Pipe As Integer = riga.IndexOf("|", indice1Pipe + 1)
                    Dim indice3Pipe As Integer = riga.IndexOf("|", indice2Pipe + 1)
                    Dim indice4Pipe As Integer = riga.IndexOf("|", indice3Pipe + 1)
                    Dim indice5Pipe As Integer = riga.IndexOf("|", indice4Pipe + 1)
                    Dim indice6Pipe As Integer = riga.IndexOf("|", indice5Pipe + 1)
                    Dim indice7Pipe As Integer = riga.IndexOf("|", indice6Pipe + 1)
                    Dim indice8Pipe As Integer = riga.IndexOf("|", indice7Pipe + 1)
                    Dim indice9Pipe As Integer = riga.IndexOf("|", indice8Pipe + 1)
                    Dim indice10Pipe As Integer = riga.IndexOf("|", indice9Pipe + 1)
                    Dim indice11Pipe As Integer = riga.IndexOf("|", indice10Pipe + 1)
                    Dim indice12Pipe As Integer = riga.IndexOf("|", indice11Pipe + 1)
                    Dim indice13Pipe As Integer = riga.IndexOf("|", indice12Pipe + 1)
                    Dim indice14Pipe As Integer = riga.IndexOf("|", indice13Pipe + 1)
                    Dim indice15Pipe As Integer = riga.IndexOf("|", indice14Pipe + 1)
                    Dim indice16Pipe As Integer = riga.IndexOf("|", indice15Pipe + 1)
                    Dim indice17Pipe As Integer = riga.IndexOf("|", indice16Pipe + 1)
                    Dim indice18Pipe As Integer = riga.IndexOf("|", indice17Pipe + 1)


                    If indice1Pipe >= 0 Then

                        Dim codice_padre As String = riga.Substring(0, indice1Pipe)

                        ' Estrai le sottostringhe tra i separatori "|"
                        Dim codice_figlio As String = riga.Substring(indice1Pipe + 1, indice2Pipe - indice1Pipe - 1)



                        Dim posizione As Integer = riga.Substring(indice2Pipe + 1, indice3Pipe - indice2Pipe - 1)

                        Dim Q As String = riga.Substring(indice3Pipe + 1, indice4Pipe - indice3Pipe - 1)

                        Dim data_creazione As String = riga.Substring(indice4Pipe + 1, indice5Pipe - indice4Pipe - 1)

                        Dim revisione As Integer = riga.Substring(indice5Pipe + 1, indice6Pipe - indice5Pipe - 1)

                        Dim descrizione As String = riga.Substring(indice7Pipe + 1, indice8Pipe - indice7Pipe - 1)

                        Dim descrizione_sup As String = riga.Substring(indice8Pipe + 1, indice9Pipe - indice8Pipe - 1)

                        Dim creato_da As String = riga.Substring(indice9Pipe + 1, indice10Pipe - indice9Pipe - 1)

                        Dim modificato_da As String = riga.Substring(indice10Pipe + 1, indice11Pipe - indice10Pipe - 1)

                        Dim UDM As String = riga.Substring(indice11Pipe + 1, indice12Pipe - indice11Pipe - 1)

                        Dim materiale As String = riga.Substring(indice12Pipe + 1, indice13Pipe - indice12Pipe - 1)

                        Dim trattamento As String = riga.Substring(indice13Pipe + 1, indice14Pipe - indice13Pipe - 1)

                        Dim ubicazione As String = riga.Substring(indice14Pipe + 1, indice15Pipe - indice14Pipe - 1)

                        Dim codice_a_barre As String = riga.Substring(indice15Pipe + 1, indice16Pipe - indice15Pipe - 1)

                        Dim disegno As String = riga.Substring(indice16Pipe + 1)







                        importazione_PDM_SAP_distinta(codice_padre, codice_figlio, posizione, Replace(Q, ",", "."), revisione, Replace(descrizione, ",", "."), Replace(descrizione_sup, ",", "."), creato_da, modificato_da, UDM, materiale, trattamento, ubicazione, codice_a_barre, disegno)
                        ' Puoi eseguire ulteriori operazioni con le sottostringhe qui

                    End If
                Next
            Else
                Console.WriteLine("Il file è vuoto.")
            End If
        Else
            Console.WriteLine("Il file non esiste.")
        End If

        Console.ReadLine()
    End Sub



    Sub importazione_PDM_SAP_ANAGRAFICA(par_codice_BRB As String, par_descrizione As String, par_descrizione_supp As String, par_revisione As String, par_Gruppo_articoli As String, par_Creato_da As String, par_Modificato_da As String, par_UDM As String, par_Materiale As String, par_Trattamento As String, par_Ubicazione As String, par_Codice_a_barre As String, par_Disegno As String, par_fornitore As String)

        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_ANAGRAFICA
           ([Data_import]
,esportato
           ,[Codice_BRB]
           ,[Descrizione]
           ,[Descrizione_supp]
 ,[revisione]
           ,[Gruppo_articoli]
           ,[Creato_da]
           ,[Modificato_da]
           ,[UDM]
           ,[Materiale]
           ,[Trattamento]
           ,[Ubicazione]
           ,[Codice_a_barre]
           ,[Disegno]
,fornitore)

     VALUES
           (getdate()
,'N'
           ,'" & par_codice_BRB & "'
           ,'" & par_descrizione & "'
           ,'" & par_descrizione_supp & "'
,' " & par_revisione & "'
,'" & par_Gruppo_articoli & "'
,' " & par_Creato_da & "'
,'" & par_Modificato_da & "'
,'" & par_UDM & "'
,'" & par_Materiale & "'
,'" & par_Trattamento & "'
,'" & par_Ubicazione & "'
,'" & par_Codice_a_barre & "'
,'" & par_Disegno & "'
,'" & par_fornitore & "')"


        Cmd_SAP.ExecuteNonQuery()



        Cnn.Close()


    End Sub

    Sub delete_frontiera_anagrafica()


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "DELETE [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_ANAGRAFICA"


        Cmd_SAP.ExecuteNonQuery()



        Cnn.Close()


    End Sub

    Sub delete_frontiera_distinta()
        Dim Cnn As New SqlConnection


        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "DELETE [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_DISTINTA"


        Cmd_SAP.ExecuteNonQuery()



        Cnn.Close()


    End Sub

    Sub importazione_PDM_SAP_distinta(par_codice_padre As String, par_codice_figlio As String, par_posizione As String, par_q As String, par_revisione As Integer, par_descrizione As String, par_descrizione_supp As String, par_creato_da As String, par_modificato_da As String, par_UDM As String, par_materiale As String, par_trattamento As String, par_ubicazione As String, par_codice_a_barre As String, par_disegno As String)

        par_descrizione = Replace(par_descrizione, "'", " ")
        par_descrizione_supp = Replace(par_descrizione_supp, "'", " ")
        par_UDM = Replace(par_UDM, "'", " ")
        par_materiale = Replace(par_materiale, "'", " ")
        par_trattamento = Replace(par_trattamento, "'", " ")
        par_codice_a_barre = Replace(par_codice_a_barre, "'", " ")
        par_disegno = Replace(par_disegno, "'", " ")
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_DISTINTA
           ([Data_import]
,esportato
,esportato_distinta
           ,[Codice_PADRE]
           ,[Codice_FIGLIO]
           ,[Posizione]
,Q
           ,[Descrizione]
           ,[Descrizione_supp]

           ,[Data_Creazione]
           ,[Revisione]
           ,[Gruppo_articoli] 
           ,[Creato_da]
           ,[Modificato_da]
           ,[UDM]
           ,[Materiale]
           ,[Trattamento]
           ,[Ubicazione]
           ,[Codice_a_barre]
           ,[Disegno]



)
     VALUES
           (getdate()
,'N'
,'N'
,'" & par_codice_padre & "'
,'" & par_codice_figlio & "'
,'" & par_posizione & "'
,'" & par_q & "'
,'" & par_descrizione & "'
,'" & par_descrizione_supp & "'
,getdate()
           ,'" & par_revisione & "'
           ,100
           ,'" & par_creato_da & "'
           ,'" & par_modificato_da & "'
           ,'" & par_UDM & "'
           ,'" & par_materiale & "'
           ,'" & par_trattamento & "'
           ,'" & par_ubicazione & "'
           ,'" & par_codice_a_barre & "'
           ,'" & par_disegno & "')"



        Cmd_SAP.ExecuteNonQuery()



        Cnn.Close()


    End Sub



    Sub Esportazione_anagrafica()
        Dim Cnn5 As New SqlConnection
        Cnn5.ConnectionString = Homepage.sap_tirelli
        Cnn5.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        Dim nuovo_Codice_tirelli As String
        CMD_SAP_2.Connection = Cnn5
        CMD_SAP_2.CommandText = "SELECT  t0.[ID]
      ,t0.[Data_import]
      ,t0.[Data_export]
      ,t0.[Esportato]
      ,t0.[Codice_BRB]
      ,COALESCE(T1.ITEMNAME,COALESCE(t0.[Descrizione],'')) AS 'Descrizione'
      ,case when t1.itemcode is null then coalesce(t0.[Descrizione_supp],'') else coalesce(t1.frgnname,'')  end as 'Descrizione_supp'
      ,case when t1.itemcode is null then coalesce(t0.[revisione],99) else coalesce(t1.u_prg_tir_rev,0) end as 'Revisione'
      ,case when t1.itemcode is null then T1.[ItmsGrpCod] else t0.[Gruppo_articoli] end as 'Gruppo_articoli_PDM'
      ,t0.[Creato_da] as 'Creato_da_pdm'
      ,t0.[Modificato_da] as 'Modificato_da_pdm'
      ,t0.[UDM]
      ,t0.[Materiale]
      ,t0.[Trattamento]
      ,case when (t0.[Ubicazione] ='' or t0.[Ubicazione] is null) then coalesce(t5.ubicazione,'') else coalesce(t0.ubicazione,'') end as 'Ubicazione'
      ,t0.[Codice_a_barre]
      ,COALESCE(T1.U_DISEGNO,COALESCE(T0.DISEGNO,'')) AS 'Disegno'
,coalesce(t0.phantom,'N') as 'phantom'
,coalesce(t1.itemcode,'') as 'Codice_SAP'
, coalesce(t2.userid,1) as 'Creato_da' 
, coalesce(t3.userid,1) as 'Modificato_da'
,t4.tirelli
,t4.Gruppo_articoli
, coalesce(COALESCE(t8.cardcode,t6.[Codice_tirelli]),'') as 'Codice_bp_Tirelli'
,coalesce(t5.costo,0) as 'Costo'
, t8.cardcode
,t8.cardname as 'nome_fornitore'
,coalesce(T1.ITEMNAME,'') AS 'Descrizione_tirelli'
,coalesce(t1.frgnname,'') as 'Desc_supp_tirelli'
,coalesce(T1.SUPPCATNUM,'') AS 'Catalogo_fornitore_tir'
,coalesce(T1.firmcode,'-1') AS 'produttore_tir'
, case when LEN(t0.[Codice_BRB])='6' then 'Uso_tir' else '' end as 'Uso_tir'

  FROM [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_ANAGRAFICA t0 
  left join [TIRELLISRLDB].[dbo].[OITM] t1 on case when LEN(t0.[Codice_BRB])='6' OR LEN(t0.[Codice_BRB])='7' then t1.itemcode else t1.u_codice_Brb end =t0.[Codice_BRB]
LEFT JOIN [TIRELLI_40].[DBO].OHEM T2 ON T2.U_CODICE_PDM=substring(t0.[Creato_da],2,1000)
LEFT JOIN [TIRELLI_40].[DBO].OHEM T3 ON T3.U_CODICE_PDM=t0.[modificato_da]
left join [TIRELLI_40].[DBO].[Frontiera_PDM_BRB_SAP_Prima_Lettera] t4 on t4.brb=substring(t0.codice_BRB,1,1)
LEFT JOIN [TIRELLI_40].[dbo].[BRB_Codici] T5 ON T5.[Codice_BRB]=t0.[Codice_BRB]
LEFT JOIN [TIRELLI_40].[DBO].[Frontiera_PDM_SAP_BPB_BP_Code] T6 ON T6.[Codice_BRB]=T5.[codice_fornitore]
LEFT JOIN [TIRELLI_40].[dbo].[Frontiera_PDM_SAP_BPB_BP_Code] T7 ON T7.[Codice_BRB]=T0.[FORNITORE]
left join [TIRELLISRLDB].[dbo].[ocrd] t8 on t7.codice_tirelli=t8.cardcode
 where t0.Esportato='N' and substring(t0.[Codice_BRB],1,1)<>'F'
 order by t0.[ID]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            If cmd_SAP_reader_2("Codice_SAP") = "" Then

                If check_esistenza_codice(cmd_SAP_reader_2("Codice_BRB")) = "Y" Then
                    nuovo_Codice_tirelli = UT.Dammi_codice(cmd_SAP_reader_2("Tirelli"))
                    If UT.check_non_duplicazione_codici(nuovo_Codice_tirelli, cmd_SAP_reader_2("Disegno"), cmd_SAP_reader_2("Codice_BRB")) = "N" Then
                        UT.inserisci_Nuovo_codice(cmd_SAP_reader_2("Creato_da"), nuovo_Codice_tirelli, cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("Descrizione_supp"), cmd_SAP_reader_2("Disegno"), cmd_SAP_reader_2("Gruppo_articoli"), cmd_SAP_reader_2("Codice_bp_tirelli"), "", "-1", "", "", "", "", "", "", "", cmd_SAP_reader_2("Codice_BRB"), cmd_SAP_reader_2("Revisione"), cmd_SAP_reader_2("phantom"), cmd_SAP_reader_2("costo"), cmd_SAP_reader_2("Ubicazione"), "M")

                    Else
                        MsgBox("Il disegno " & cmd_SAP_reader_2("Disegno") & " codice BRB " & cmd_SAP_reader_2("Codice_BRB") & " è già presente a Database ")
                    End If



                End If

            Else

                Magazzino.UPDATE_OITM(cmd_SAP_reader_2("modificato_da"), cmd_SAP_reader_2("Codice_Sap"), cmd_SAP_reader_2("Descrizione_tirelli"), cmd_SAP_reader_2("Desc_supp_tirelli"), cmd_SAP_reader_2("Disegno"), cmd_SAP_reader_2("Codice_bp_tirelli"), cmd_SAP_reader_2("Catalogo_fornitore_tir"), cmd_SAP_reader_2("produttore_tir"), "", "", "", "", "", "", "", "", cmd_SAP_reader_2("Revisione"))


            End If
            aggiorna_esportato_anagrafica(cmd_SAP_reader_2("id"))


        Loop

        Cnn5.Close()

    End Sub

    Private Function check_esistenza_codice(PAR_CODICE_BRB)
        Dim risultato As String = "Y"
        Dim Cnn4 As New SqlConnection
        Cnn4.ConnectionString = Homepage.sap_tirelli
        Cnn4.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn4
        CMD_SAP_2.CommandText = "SELECT  
      t0.[Codice_BRB]
      

  FROM [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_ANAGRAFICA t0 
INNER join [TIRELLISRLDB].[dbo].[OITM] t1 ON T0.CODICE_BRB=T1.U_CODICE_BRB
 where t0.[Codice_BRB]='" & PAR_CODICE_BRB & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then

            risultato = "N"


        End If

        Cnn4.Close()

        Return risultato

    End Function


    Sub Esportazione_distinta_anagrafiche()
        Dim Cnn5 As New SqlConnection
        Cnn5.ConnectionString = Homepage.sap_tirelli
        Cnn5.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        Dim nuovo_Codice_tirelli As String
        CMD_SAP_2.Connection = Cnn5
        CMD_SAP_2.CommandText = "SELECT  t0.[ID]
      ,t0.[Data_import]
      ,t0.[Data_export]
      ,t0.[Esportato]
,t0.[Esportato_distinta]
	  ,t0.[Codice_padre]
      ,t0.[Codice_figlio]
	  ,t0.posizione
	  ,t0.q
      ,coalesce(t1.itemname,coalesce(t0.[Descrizione],''),'') as 'Descrizione'
      ,case when t1.itemcode is null then coalesce(t0.[Descrizione_supp],'') else t1.frgnname end as 'Descrizione_supp'
      ,case when t1.itemcode is null then t0.[revisione] else t1.u_prg_tir_rev end as 'Revisione'
      ,case when t1.itemcode is null then T1.[ItmsGrpCod] else t0.[Gruppo_articoli] end as 'Gruppo_articoli_PDM'
      ,t0.[Creato_da] as 'Creato_da_pdm'
      ,t0.[Modificato_da] as 'modificato_da_pdm'
      ,t0.[UDM]
      ,t0.[Materiale]
      ,t0.[Trattamento]
      ,case when (t0.[Ubicazione]='' or t0.[Ubicazione] is null) then coalesce(t5.ubicazione,'') else coalesce(t0.ubicazione,'') end as 'Ubicazione'
      ,t0.[Codice_a_barre]
      ,COALESCE(T1.U_DISEGNO,COALESCE(T0.DISEGNO,'')) AS 'Disegno'
,coalesce(t0.phantom,'N') as 'phantom'
,coalesce(t1.itemcode,'') as 'Codice_SAP'
, coalesce(t2.userid,1) as 'Creato_da' 
, coalesce(t3.userid,1) as 'Modificato_da'
,t4.tirelli
,t4.Gruppo_articoli
,coalesce(t5.costo,0) as 'Costo'
,t1.itemname as 'Descrizione_Tirelli'
,t1.frgnname as 'Desc_supp_Tirelli'

  FROM [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_DISTINTA t0 
 left join [TIRELLISRLDB].[dbo].[OITM] t1 on case when LEN(t0.[Codice_figlio])='6' OR LEN(t0.[Codice_figlio])='7' then t1.itemcode else t1.u_codice_Brb end =t0.[Codice_figlio]
LEFT JOIN [TIRELLI_40].[DBO].OHEM T2 ON T2.U_CODICE_PDM=substring(t0.[Creato_da],2,1000)
LEFT JOIN [TIRELLI_40].[DBO].OHEM T3 ON T3.U_CODICE_PDM=t0.[modificato_da]
left join [TIRELLI_40].[DBO].[Frontiera_PDM_BRB_SAP_Prima_Lettera] t4 on t4.brb=substring(t0.codice_figlio,1,1)
left join [TIRELLI_40].[DBO].[brb_codici] t5 on t0.[Codice_figlio]=t5.[Codice_BRB]
  where t0.Esportato='N'
  order by t0.[ID]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            If cmd_SAP_reader_2("Codice_SAP") = "" Then
                nuovo_Codice_tirelli = UT.Dammi_codice(cmd_SAP_reader_2("Tirelli"))

                UT.inserisci_Nuovo_codice(cmd_SAP_reader_2("Creato_da"), nuovo_Codice_tirelli, cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("Descrizione_supp"), cmd_SAP_reader_2("Disegno"), cmd_SAP_reader_2("Gruppo_articoli"), "", "", "-1", "", "", "", "", "", "", "", cmd_SAP_reader_2("Codice_figlio"), cmd_SAP_reader_2("Revisione"), cmd_SAP_reader_2("phantom"), cmd_SAP_reader_2("costo"), cmd_SAP_reader_2("ubicazione"), "M")

            Else

                Magazzino.UPDATE_OITM(cmd_SAP_reader_2("modificato_da"), cmd_SAP_reader_2("Codice_Sap"), cmd_SAP_reader_2("Descrizione_tirelli"), cmd_SAP_reader_2("Desc_supp_tirelli"), cmd_SAP_reader_2("Disegno"), "", "", "-1", "", "", "", "", "", "", "", "", cmd_SAP_reader_2("Revisione"))


            End If

            aggiorna_esportato_distinta(cmd_SAP_reader_2("id"))
        Loop

        Cnn5.Close()

    End Sub

    Sub Esportazione_distinta_legami()
        delete_itt1_distinte_esportate()
        delete_oitt_distinte_esportate()
        Dim Cnn5 As New SqlConnection
        Cnn5.ConnectionString = Homepage.sap_tirelli
        Cnn5.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        Dim Distinta_base_precedente As String = "PREC"

        Dim codice_tirelli_figlio As String

        CMD_SAP_2.Connection = Cnn5
        CMD_SAP_2.CommandText = "SELECT  t0.[ID]
      ,t0.[Data_import]
      ,t0.[Data_export]
      ,t0.[Esportato]
,t0.[Esportato_distinta]
	  ,t0.[Codice_padre]
      ,t0.[Codice_figlio]
	  ,t0.posizione
	  ,t0.q
      ,t0.[Descrizione]
      ,t0.[Descrizione_supp]
      ,t0.[revisione]
      ,t0.[Gruppo_articoli] as 'Gruppo_articoli_PDM'
      ,t0.[Creato_da] as 'Creato_da_pdm'
      ,t0.[Modificato_da] as 'modificato_da_pdm'
      ,t0.[UDM]
      ,t0.[Materiale]
      ,t0.[Trattamento]
      ,case when (t0.[Ubicazione]='' or t0.[Ubicazione] is null) then coalesce(t6.ubicazione,'') else t0.ubicazione end as 'Ubicazione'
      ,t0.[Codice_a_barre]
      ,t0.[Disegno]
,coalesce(t0.phantom,'N') as 'Phantom'
,coalesce(t1.itemcode,'') as 'Codice_SAP_figlio'
,coalesce(t5.itemcode,'') as 'Codice_SAP_padre'
,coalesce(t5.itemname,'') as 'nome_SAP_padre'

, coalesce(t2.userid,1) as 'Creato_da' 
, coalesce(t3.userid,1) as 'Modificato_da'
,t4.tirelli
,t4.Gruppo_articoli
,coalesce(t6.costo,0) as 'Costo'
, COALESCE(T7.CODE,'') AS 'Replica_figlio'
,T7.FATHER

  FROM [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_DISTINTA t0 
 left join [TIRELLISRLDB].[dbo].[OITM] t1 on case when len(t0.[Codice_figlio])='6' OR len(t0.[Codice_figlio])='7' then t1.itemcode else t1.u_codice_Brb end=t0.[Codice_figlio]
LEFT JOIN [TIRELLI_40].[DBO].OHEM T2 ON T2.U_CODICE_PDM=substring(t0.[Creato_da],2,1000)
LEFT JOIN [TIRELLI_40].[DBO].OHEM T3 ON T3.U_CODICE_PDM=t0.[modificato_da]
left join [TIRELLI_40].[DBO].[Frontiera_PDM_BRB_SAP_Prima_Lettera] t4 on t4.brb=substring(t0.codice_figlio,1,1)
 left join [TIRELLISRLDB].[dbo].[OITM] t5 on case when len(t0.[Codice_padre])='6' OR len(t0.[Codice_padre])='7' then t5.itemcode else t5.u_codice_Brb end=t0.[Codice_padre]
left join [TIRELLI_40].[DBO].[brb_codici] t6 on t0.[Codice_figlio]=t6.[Codice_BRB]  
LEFT JOIN [TIRELLISRLDB].[dbo].[ITT1] t7 on t5.itemcode=t7.[FATHER]  AND (T7.CHILDNUM=T0.POSIZIONE OR T7.VisOrder=T0.POSIZIONE)
where t0.Esportato_distinta='N'  AND COALESCE(t0.[Codice_figlio],'')<>''
  order by t0.[ID]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Dim contatore As String = 0
        Do While cmd_SAP_reader_2.Read()

            If cmd_SAP_reader_2("codice_sap_padre") = "" Then
                MsgBox("Creare il codice padre" & " " & cmd_SAP_reader_2("codice_padre") & " prima di crearne la distinta base")

            Else
                If cmd_SAP_reader_2("codice_sap_figlio") = "" Then
                    MsgBox("Creare il codice figlio" & " " & cmd_SAP_reader_2("codice_sap_figlio") & " per poterlo inserire nelle righe della DB")
                Else

                    If cmd_SAP_reader_2("q") >= 1 Then


                        If cmd_SAP_reader_2("codice_sap_padre") = cmd_SAP_reader_2("codice_sap_figlio") Then

                            MsgBox("La distinta " & cmd_SAP_reader_2("codice_sap_padre") & " Contiene nelle righe se stesso. Controllare")

                        Else
                            Dim LOGINSTANC As Integer = Distinta_base_form.Trova_ultimo_loginstanc_distinta(cmd_SAP_reader_2("codice_sap_padre"))
                            Distinta_base_form.INSERT_INTO_OITT(cmd_SAP_reader_2("codice_sap_padre"), 1, cmd_SAP_reader_2("nome_sap_padre"), cmd_SAP_reader_2("creato_da"), LOGINSTANC)


                            Distinta_base_form.MEttere_db_produzione_oitm(cmd_SAP_reader_2("codice_sap_padre"))



                            Distinta_base_form.delete_itt1(cmd_SAP_reader_2("codice_sap_padre"), cmd_SAP_reader_2("posizione"), cmd_SAP_reader_2("codice_sap_figlio"))


                            If cmd_SAP_reader_2("Codice_SAP_figlio") = "" Then
                                codice_tirelli_figlio = UT.Dammi_codice(cmd_SAP_reader_2("Tirelli"))

                                UT.inserisci_Nuovo_codice(cmd_SAP_reader_2("Creato_da"), codice_tirelli_figlio, cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("Descrizione_supp"), cmd_SAP_reader_2("Disegno"), cmd_SAP_reader_2("Gruppo_articoli"), "", "", "-1", "", "", "", "", "", "", "", cmd_SAP_reader_2("Codice_figlio"), cmd_SAP_reader_2("Revisione"), cmd_SAP_reader_2("phantom"), cmd_SAP_reader_2("costo"), cmd_SAP_reader_2("ubicazione"), "M")


                                Distinta_base_form.INSERT_INTO_ITT1(cmd_SAP_reader_2("codice_sap_padre"), codice_tirelli_figlio, cmd_SAP_reader_2("q"), cmd_SAP_reader_2("costo"), 0, cmd_SAP_reader_2("posizione"), "B01", cmd_SAP_reader_2("descrizione"), LOGINSTANC, "IVSAPLINK")

                            Else

                                Distinta_base_form.INSERT_INTO_ITT1(cmd_SAP_reader_2("codice_sap_padre"), cmd_SAP_reader_2("Codice_SAP_figlio"), cmd_SAP_reader_2("q"), cmd_SAP_reader_2("q"), 0, cmd_SAP_reader_2("posizione"), "B01", cmd_SAP_reader_2("descrizione"), LOGINSTANC, "IVSAPLINK")

                            End If

                        End If



                    Else
                        MsgBox("Selezionare una quantità > 0 PER IL CODICE " & cmd_SAP_reader_2("codice_FIGLIO"))

                    End If
                End If

                aggiorna_esportato_distinta_legami(cmd_SAP_reader_2("id"))
            End If


        Loop

        cnn5.Close()

    End Sub

    Sub riempi_datagridview_anagrafica()
        DataGridView_anagrafica.Rows.Clear()
        Dim Cnn5 As New SqlConnection
        Cnn5.ConnectionString = homepage.sap_tirelli
        cnn5.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn5
        CMD_SAP_2.CommandText = "SELECT  t0.[ID]
      ,t0.[Data_import]
      ,t0.[Data_export]
      ,t0.[Esportato]
      ,t0.[Codice_BRB]
      ,t0.[Descrizione]
      ,t0.[Descrizione_supp]
      ,t0.[revisione]
      ,t0.[Gruppo_articoli]
      ,t0.[Creato_da]
      ,t0.[Modificato_da]
      ,t0.[UDM]
      ,t0.[Materiale]
      ,t0.[Trattamento]
      ,case when t0.[Ubicazione] ='' or t0.[Ubicazione] is null then coalesce(t5.ubicazione,'') end as 'Ubicazione'
      ,t0.[Codice_a_barre]
      ,t0.[Disegno]
,coalesce(t1.itemcode,'') as 'Codice_SAP'
, coalesce(t2.userid,0) as 'Creato_da_2' 
, coalesce(t3.userid,0) as 'Modificato_da_2'
,t4.tirelli
,t4.Gruppo_articoli as 'Gruppo_articoli_2'
,T7.CARDCODE AS 'Codice_fornitore'
,T7.cardname AS 'Nome_fornitore'

  FROM [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_ANAGRAFICA t0 
  left join [TIRELLISRLDB].[dbo].[OITM] t1 on case when len(t0.[Codice_BRB])='6' OR  len(t0.[Codice_BRB])='7' then t1.itemcode else t1.u_codice_Brb end=t0.[Codice_BRB]
LEFT JOIN [TIRELLI_40].[DBO].OHEM T2 ON T2.U_CODICE_PDM=substring(t0.[Creato_da],2,1000)
LEFT JOIN [TIRELLI_40].[DBO].OHEM T3 ON T3.U_CODICE_PDM=t0.[modificato_da]
left join [TIRELLI_40].[DBO].[Frontiera_PDM_BRB_SAP_Prima_Lettera] t4 on t4.brb=substring(t0.codice_BRB,1,1)
left join [TIRELLI_40].[dbo].[BRB_codici] t5 on t5.codice_brb =t0.codice_brb
left join [TIRELLI_40].[DBO].[Frontiera_PDM_SAP_BPB_BP_Code] T6 ON T0.FORNITORE=T6.CODICE_BRB
LEFT JOIN OCRD T7 ON T7.CARDCODE=T6.CODICE_TIRELLI
  
  order by t0.[ID] DESC"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            DataGridView_anagrafica.Rows.Add(
    cmd_SAP_reader_2("ID"),
    cmd_SAP_reader_2("Data_import"),
    cmd_SAP_reader_2("Data_export"),
    cmd_SAP_reader_2("Esportato"),
    cmd_SAP_reader_2("Codice_BRB"),
    cmd_SAP_reader_2("Descrizione"),
    cmd_SAP_reader_2("Descrizione_supp"),
    cmd_SAP_reader_2("revisione"),
    cmd_SAP_reader_2("Gruppo_articoli"),
    cmd_SAP_reader_2("Creato_da"),
    cmd_SAP_reader_2("Modificato_da"),
    cmd_SAP_reader_2("UDM"),
    cmd_SAP_reader_2("Materiale"),
    cmd_SAP_reader_2("Trattamento"),
    cmd_SAP_reader_2("Ubicazione"),
    cmd_SAP_reader_2("Codice_a_barre"),
    cmd_SAP_reader_2("Disegno"),
    cmd_SAP_reader_2("Codice_SAP"),
     cmd_SAP_reader_2("Creato_Da_2"),
    cmd_SAP_reader_2("Modificato_Da_2"),
    cmd_SAP_reader_2("Tirelli"),
    cmd_SAP_reader_2("Gruppo_articoli_2"),
    cmd_SAP_reader_2("Nome_fornitore"))




        Loop

        cnn5.Close()

    End Sub

    Private Sub tabpage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter

        riempi_datagridview_distinta()



    End Sub

    Sub riempi_datagridview_distinta()
        DataGridView_distinta.Rows.Clear()
        Dim Cnn5 As New SqlConnection
        Cnn5.ConnectionString = homepage.sap_tirelli
        cnn5.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn5
        CMD_SAP_2.CommandText = "SELECT  t0.[ID]
      ,t0.[Data_import]
      ,t0.[Data_export]
      ,t0.[Esportato]
,t0.[Esportato_distinta]
	  ,t0.[Codice_padre]
      ,t0.[Codice_figlio]
	  ,t0.posizione
	  ,t0.q
      ,t0.[Descrizione]
      ,t0.[Descrizione_supp]
      ,t0.[revisione]
      ,t0.[Gruppo_articoli] as 'Gruppo_articoli_PDM'
      ,t0.[Creato_da] as 'Creato_da_pdm'
      ,t0.[Modificato_da] as 'modificato_da_pdm'
      ,t0.[UDM]
      ,t0.[Materiale]
      ,t0.[Trattamento]
      ,case when t0.[Ubicazione] ='' or t0.[Ubicazione] is null then coalesce(t6.ubicazione,'') end as 'Ubicazione'
      ,t0.[Codice_a_barre]
      ,t0.[Disegno]
,coalesce(t1.itemcode,'') as 'Codice_SAP_figlio'
, coalesce(t2.userid,1) as 'Creato_da' 
, coalesce(t3.userid,1) as 'Modificato_da'
,t4.tirelli
,t4.Gruppo_articoli
,coalesce(t5.itemcode,'') as 'Codice_SAP_padre'

  FROM [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_DISTINTA t0 
  left join [TIRELLISRLDB].[dbo].[OITM] t1 on case when len(t0.[Codice_figlio])='6' OR len(t0.[Codice_figlio])='7' then t1.itemcode else t1.u_codice_Brb end=t0.[Codice_figlio]
LEFT JOIN [TIRELLI_40].[DBO].OHEM T2 ON T2.U_CODICE_PDM=substring(t0.[Creato_da],2,1000)
LEFT JOIN [TIRELLI_40].[DBO].OHEM T3 ON T3.U_CODICE_PDM=t0.[modificato_da]
left join [TIRELLI_40].[DBO].[Frontiera_PDM_BRB_SAP_Prima_Lettera] t4 on t4.brb=substring(t0.codice_figlio,1,1)
 left join [TIRELLISRLDB].[dbo].[OITM] t5 on case when len(t0.[Codice_padre])='6' then t5.itemcode else t5.u_codice_Brb end=t0.[Codice_padre]
left join [TIRELLI_40].[DBO].[BRB_codici] t6 on t6.codice_brb =t0.codice_figlio

  order by t0.[ID] DESC"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            DataGridView_distinta.Rows.Add(
    cmd_SAP_reader_2("ID"),
    cmd_SAP_reader_2("Data_import"),
    cmd_SAP_reader_2("Data_export"),
    cmd_SAP_reader_2("Esportato"),
    cmd_SAP_reader_2("esportato_distinta"),
    cmd_SAP_reader_2("codice_padre"),
    cmd_SAP_reader_2("codice_figlio"),
    cmd_SAP_reader_2("posizione"),
    cmd_SAP_reader_2("q"),
    cmd_SAP_reader_2("Descrizione"),
    cmd_SAP_reader_2("Descrizione_supp"),
    cmd_SAP_reader_2("revisione"),
    cmd_SAP_reader_2("Gruppo_articoli"),
    cmd_SAP_reader_2("Creato_da"),
    cmd_SAP_reader_2("Modificato_da"),
    cmd_SAP_reader_2("UDM"),
    cmd_SAP_reader_2("Materiale"),
    cmd_SAP_reader_2("Trattamento"),
    cmd_SAP_reader_2("Ubicazione"),
    cmd_SAP_reader_2("Codice_a_barre"),
    cmd_SAP_reader_2("Disegno"),
    cmd_SAP_reader_2("Codice_SAP_padre"),
    cmd_SAP_reader_2("Codice_SAP_figlio"),
     cmd_SAP_reader_2("Creato_Da"),
    cmd_SAP_reader_2("Modificato_Da"),
    cmd_SAP_reader_2("Tirelli"),
    cmd_SAP_reader_2("Gruppo_articoli"))

        Loop

        cnn5.Close()

    End Sub

    Sub aggiorna_esportato_anagrafica(par_id As Integer)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = " UPDATE [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_ANAGRAFICA SET ESPORTATO='Y', Data_export=getdate() where id=" & par_id & "

"
        Cmd_SAP.ExecuteNonQuery()
        cnn.Close()
    End Sub

    Sub delete_itt1_distinte_esportate()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = cnn
        Cmd_SAP.CommandText = " delete t11
from
(
SELECT  coalesce(t5.itemcode,'') as 'Codice_SAP_padre'


  FROM [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_DISTINTA t0 
  left join [TIRELLISRLDB].[dbo].[OITM] t1 on t1.u_codice_brb=t0.[Codice_figlio]

left join [TIRELLISRLDB].[dbo].[OITM] t5 on t5.u_codice_brb =t0.[Codice_padre]

where t0.Esportato_distinta='N' AND COALESCE(t0.[Codice_figlio],'')<>''
group by coalesce(t5.itemcode,'')
)
as t10 left join itt1 t11 on t11.father=t10.Codice_SAP_padre
"
        Cmd_SAP.ExecuteNonQuery()
        cnn.Close()
    End Sub

    Sub delete_oitt_distinte_esportate()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = cnn
        Cmd_SAP.CommandText = "delete t11
from
(
SELECT  coalesce(t5.itemcode,'') as 'Codice_SAP_padre'


  FROM [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_DISTINTA t0 
  left join [TIRELLISRLDB].[dbo].[OITM] t1 on t1.u_codice_brb=t0.[Codice_figlio]

left join [TIRELLISRLDB].[dbo].[OITM] t5 on t5.u_codice_brb =t0.[Codice_padre]

where t0.Esportato_distinta='N' AND COALESCE(t0.[Codice_figlio],'')<>''
group by coalesce(t5.itemcode,'')
)
as t10 left join oitt t11 on t11.code=t10.Codice_SAP_padre"
        Cmd_SAP.ExecuteNonQuery()
        cnn.Close()
    End Sub

    Sub aggiorna_esportato_distinta(par_id As Integer)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = " UPDATE [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_DISTINTA SET ESPORTATO='Y', Data_export=getdate() where id=" & par_id & "

"
        Cmd_SAP.ExecuteNonQuery()
        cnn.Close()
    End Sub

    Sub aggiorna_esportato_distinta_legami(par_id As Integer)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = " UPDATE [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_DISTINTA SET ESPORTATO_distinta='Y', Data_export=getdate() where id=" & par_id & "

"
        Cmd_SAP.ExecuteNonQuery()
        cnn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        leggi_file_anagrafica()
        leggi_file_distinta()

        MsgBox("Importazione effettuata con successo")
        mostra_file()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Esportazione_anagrafica()
        Esportazione_distinta_anagrafiche()
        Esportazione_distinta_legami()
        MsgBox("Esportazione effettuata con successo")
        mostra_file()
    End Sub

    Sub Comprimi_anagrafica()
        ' Percorso del file originale
        Dim percorsoFile_comprimere As String = Homepage.percorso_PDM_BRB & "\anagrafica.txt"

        ' Verifica se il file esiste
        If File.Exists(percorsoFile_comprimere) Then
            ' Leggi tutte le righe del file
            Dim righe As List(Of String) = File.ReadAllLines(percorsoFile_comprimere).ToList()

            ' Verifica se ci sono righe nel file
            If righe.Count > 0 Then
                For i As Integer = 0 To righe.Count - 1
                    ' Dividi la riga in base al carattere "|" e prendi solo la parte prima del quarto "|"
                    Dim partiRiga As String() = righe(i).Split("|"c)
                    If partiRiga.Length >= TextBox1.Text Then
                        Dim parteDesiderata As String = String.Join("|", partiRiga.Take(TextBox1.Text))
                        righe(i) = parteDesiderata ' Sovrascrivi la riga nel file
                    Else
                        Console.WriteLine("La riga " & (i + 1) & " non ha abbastanza elementi separati da '|'.")
                    End If
                Next

                ' Sovrascrivi il file originale con le nuove righe modificate
                File.WriteAllLines(percorsoFile_comprimere, righe)

            Else
                Console.WriteLine("Il file è vuoto.")
            End If
        Else
            Console.WriteLine("Il file non esiste.")
        End If

        Console.ReadLine()
    End Sub

    Sub Comprimi_legami()
        ' Percorso del file originale
        Dim percorsoFile_comprimere As String = Homepage.percorso_PDM_BRB & "\legami.txt"

        ' Verifica se il file esiste
        If File.Exists(percorsoFile_comprimere) Then
            ' Leggi tutte le righe del file
            Dim righe As List(Of String) = File.ReadAllLines(percorsoFile_comprimere).ToList()

            ' Verifica se ci sono righe nel file
            If righe.Count > 0 Then
                For i As Integer = 0 To righe.Count - 1
                    ' Dividi la riga in base al carattere "|" e prendi solo la parte prima del quarto "|"
                    Dim partiRiga As String() = righe(i).Split("|"c)
                    If partiRiga.Length >= TextBox2.Text Then
                        Dim parteDesiderata As String = String.Join("|", partiRiga.Take(TextBox2.Text))
                        righe(i) = parteDesiderata ' Sovrascrivi la riga nel file
                    Else
                        Console.WriteLine("La riga " & (i + 1) & " non ha abbastanza elementi separati da '|'.")
                    End If
                Next

                ' Sovrascrivi il file originale con le nuove righe modificate
                File.WriteAllLines(percorsoFile_comprimere, righe)

            Else
                Console.WriteLine("Il file è vuoto.")
            End If
        Else
            Console.WriteLine("Il file non esiste.")
        End If

        Console.ReadLine()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Comprimi_anagrafica()
        Comprimi_legami()
        MsgBox("Il file è stato modificato con successo.")
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub PDM_SAP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LinkLabel1.Text = Homepage.percorso_PDM_BRB
        mostra_file()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Sub mostra_file()

        ' Verifica se rootDirectory esiste
        Dim rootDirectoryPath As String = Homepage.percorso_PDM_BRB
        If Not Directory.Exists(rootDirectoryPath) Then

            Return ' Esce dalla sub

        End If

        ' Cancella tutti i nodi precedenti nel TreeView
        TreeView1.Nodes.Clear()

        ' Aggiunge il nodo radice al TreeView
        Dim rootDirectory As New DirectoryInfo(rootDirectoryPath)

        Dim rootNode As New TreeNode(rootDirectory.Name)
        rootNode.Tag = rootDirectory
        TreeView1.Nodes.Add(rootNode)

        ' Aggiunge tutti i nodi figli del nodo radice
        AddDirectories(rootNode)
        Addfiles(rootNode)
        ' Espande tutti i nodi del TreeView
        TreeView1.ExpandAll()

        ' Abilita il trascinamento dei file sulla TreeView
        TreeView1.AllowDrop = True


    End Sub

    Public Sub AddDirectories(parentNode As TreeNode)
        Dim parentDirectory As DirectoryInfo = TryCast(parentNode.Tag, DirectoryInfo)

        ' Aggiunge tutte le cartelle come nodi figli
        For Each directory As DirectoryInfo In parentDirectory.GetDirectories()
            Dim directoryNode As New TreeNode(directory.Name)
            directoryNode.Tag = directory

            ' Aggiunge l'icona della cartella alla ImageList
            Dim folderIcon As Icon = Icon.ExtractAssociatedIcon("C:\Windows\System32\shell32.dll")
            ImageList1.Images.Add("folder", folderIcon)

            ' Imposta l'icona della cartella sulla chiave "folder" nell'ImageList
            directoryNode.ImageKey = "folder"

            parentNode.Nodes.Add(directoryNode)

            ' Aggiunge tutti i file come nodi figli della cartella
            For Each file As FileInfo In directory.GetFiles()
                Dim fileNode As New TreeNode(file.Name & " (" & file.LastWriteTime.ToString() & ")")
                fileNode.Tag = file

                ' Ottiene l'icona del file
                Dim fileIcon As Icon = SystemIcons.WinLogo
                Dim filepath As String
                filepath = file.FullName

                filepath = Replace(filepath, Homepage.percorso_server & "disco_e", "Y:")
                Try
                    fileIcon = Icon.ExtractAssociatedIcon(filepath)
                Catch ex As Exception

                End Try


                ' Aggiunge l'icona alla ImageList e imposta la proprietà ImageKey del nodo file
                If Not ImageList1.Images.ContainsKey(file.Extension) Then
                    ImageList1.Images.Add(file.Extension, fileIcon)
                End If
                fileNode.ImageKey = file.Extension

                directoryNode.Nodes.Add(fileNode)
            Next
            ' Ricorsivamente aggiunge tutti i nodi figli della cartella
            AddDirectories(directoryNode)

        Next





    End Sub


    Public Sub Addfiles(parentNode As TreeNode)
        Dim parentDirectory As DirectoryInfo = TryCast(parentNode.Tag, DirectoryInfo)
        Try
            For Each file As FileInfo In parentDirectory.GetFiles()
                Dim fileNode As New TreeNode(file.Name & " (" & file.LastWriteTime.ToString() & ")")
                fileNode.Tag = file

                ' Ottiene l'icona del file
                Dim fileIcon As Icon = SystemIcons.WinLogo
                Dim filepath As String
                filepath = file.FullName

                filepath = Replace(filepath, Homepage.percorso_server & "disco_e", "Y:")
                Try
                    fileIcon = Icon.ExtractAssociatedIcon(filepath)
                Catch ex As Exception

                End Try


                ' Aggiunge l'icona alla ImageList e imposta la proprietà ImageKey del nodo file
                If Not ImageList1.Images.ContainsKey(file.Extension) Then
                    ImageList1.Images.Add(file.Extension, fileIcon)
                End If
                fileNode.ImageKey = file.Extension

                parentNode.Nodes.Add(fileNode)
            Next
        Catch ex As Exception

        End Try
        ' Aggiunge anche i file direttamente presenti nella cartella padre




    End Sub



    Private Sub Apri_file_Click_1(sender As Object, e As EventArgs) Handles Apri_file.Click
        ' Verifica se il nodo selezionato è un file
        If TypeOf TreeView1.SelectedNode.Tag Is FileInfo Then
            ' Se il nodo selezionato è un file, apri il file
            Dim file As FileInfo = DirectCast(TreeView1.SelectedNode.Tag, FileInfo)
            Process.Start(file.FullName)
        ElseIf TypeOf TreeView1.SelectedNode.Tag Is DirectoryInfo Then
            ' Se il nodo selezionato è una directory, apri la cartella
            Dim directory As DirectoryInfo = DirectCast(TreeView1.SelectedNode.Tag, DirectoryInfo)
            Process.Start("explorer.exe", directory.FullName)
        End If
    End Sub

    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        If e.Button = MouseButtons.Right Then
            TreeView1.SelectedNode = TreeView1.GetNodeAt(e.X, e.Y)
        End If
    End Sub

    Private Sub TreeView1_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick
        ' Verifica se il nodo selezionato è una directory
        If TypeOf e.Node.Tag Is FileInfo Then
            ' Apri il file con l'applicazione predefinita
            Dim file As FileInfo = DirectCast(e.Node.Tag, FileInfo)
            Process.Start(file.FullName)
        ElseIf TypeOf e.Node.Tag Is DirectoryInfo Then
            ' Se il nodo selezionato è una directory, apri la cartella
            Dim directory As DirectoryInfo = DirectCast(e.Node.Tag, DirectoryInfo)
            Process.Start("explorer.exe", directory.FullName)
        End If


    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start(LinkLabel1.Text)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        riempi_datagridview_anagrafica()
        riempi_datagridview_distinta()
        Beep()
    End Sub

    Private Sub DataGridView_anagrafica_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_anagrafica.CellContentClick

    End Sub

    Private Sub DataGridView_anagrafica_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_anagrafica.CellFormatting

        If DataGridView_anagrafica.Rows(e.RowIndex).Cells(columnName:="Esportato").Value = "N" Then
            DataGridView_anagrafica.Rows(e.RowIndex).Cells(columnName:="Esportato").Style.BackColor = Color.Yellow
        ElseIf DataGridView_anagrafica.Rows(e.RowIndex).Cells(columnName:="Esportato").Value = "Y" Then
            DataGridView_anagrafica.Rows(e.RowIndex).Cells(columnName:="Esportato").Style.BackColor = Color.Lime

        End If


    End Sub


    Private Sub DataGridView_distinta_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_distinta.CellFormatting

        If DataGridView_distinta.Rows(e.RowIndex).Cells(columnName:="Esportato_anagrafica").Value = "N" Then
            DataGridView_distinta.Rows(e.RowIndex).Cells(columnName:="Esportato_anagrafica").Style.BackColor = Color.Yellow
        ElseIf DataGridView_distinta.Rows(e.RowIndex).Cells(columnName:="Esportato_anagrafica").Value = "Y" Then
            DataGridView_distinta.Rows(e.RowIndex).Cells(columnName:="Esportato_anagrafica").Style.BackColor = Color.Lime

        End If

        If DataGridView_distinta.Rows(e.RowIndex).Cells(columnName:="Esportato_distinta").Value = "N" Then
            DataGridView_distinta.Rows(e.RowIndex).Cells(columnName:="Esportato_distinta").Style.BackColor = Color.Yellow
        ElseIf DataGridView_distinta.Rows(e.RowIndex).Cells(columnName:="Esportato_distinta").Value = "Y" Then
            DataGridView_distinta.Rows(e.RowIndex).Cells(columnName:="Esportato_distinta").Style.BackColor = Color.Lime

        End If
    End Sub


    Sub importazione_distinta_da_Excel()

        Dim contatore_excel As Integer = 45

        Dim lunghezza_excel As Integer = 1398


        Dim NOME_file As String = "Report_Matr1429"
        Dim NOME_FOGLIO As String = "TT_REPORT1"


        Excel = CreateObject("Excel.application")
        WB_Excel = Excel.Workbooks.Open("\\tirfs01\00-Tirelli 4.0\BRB\" & NOME_file & ".xls")

        Excel.Visible = True

        Dim codice_nonno As String
        Dim descrizione As String
        Dim Descrizione_supp As String
        Dim paese As String
        Dim agente As String
        Dim cliente As String
        Dim anno As String
        codice_nonno = Excel.Sheets(NOME_FOGLIO).Cells(40, 2).value

        Do While contatore_excel <= lunghezza_excel



            'itemcode = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 2).value

            'descrizione = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 6).value
            'Descrizione_supp = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 7).value
            'cliente = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 10).value
            'paese = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 11).value
            'agente = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 9).value
            'anno = Excel.Sheets(NOME_FOGLIO).Cells(contatore_excel, 14).value

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

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim percorso_file As String

        Using openFileDialog1 As New OpenFileDialog()
            openFileDialog1.Title = "Seleziona un file"
            openFileDialog1.Filter = "Tutti i file (*.*)|*.*"

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                percorso_file = openFileDialog1.FileName
                ' Puoi fare qualcos'altro con il percorso_file qui
            Else
                MsgBox("Scegliere un file valido")
            End If
        End Using
        If ComboBox1.Text = "Business" Then
            trova_dato_da_excel_pEr_importazionE(percorso_file, TextBox3.Text, TextBox5.Text, TextBox4.Text)
        Else
            trova_dato_da_excel_pEr_importazionE_da_pdm(percorso_file, TextBox3.Text, TextBox5.Text, TextBox4.Text)
        End If


        Beep()
        MsgBox("Importazione completata")


    End Sub

    Sub trova_dato_da_excel_pEr_importazionE_quadro(par_percorso_file As String, par_nome_foglio As String, par_riga_inizio As Integer, par_riga_fine As Integer)

        TextBox6.Text = Replace(TextBox6.Text, "'", "")
        TextBox6.Text = Replace(TextBox6.Text, ",", ".")
        Dim figlio As String
        Dim padre As String
        Dim Descrizione As String
        'Dim desc_supp As String

        Dim um As String
        Dim codice_fornitore As String
        Dim fornitore As String
        Dim produttore As String



        Dim posizione As String
        Dim quantita As String
        Dim note As String


        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True


        Do While par_riga_inizio <= par_riga_fine

            If Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value <> Nothing Then


                figlio = "E_" & Replace(TextBox6.Text & "-" & Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value, " ", "")
                Descrizione = "IMPIANTO ELETTRICO " & Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value & " " & TextBox6.Text

                If check_anagrafica_gia_importata(figlio) = "NO" Then

                    importazione_PDM_SAP_ANAGRAFICA(figlio, Replace(Descrizione, "'", " "), Replace("", "'", " "), 0, 100, "G.T.", "G.T.", Replace("", "'", " "), Replace("", "'", " "), Replace("", "'", " "), Replace("", "'", " "), Replace("", "'", " "), Replace("", "'", " "), "")

                End If

                Descrizione = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value
                figlio = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 6).value

                codice_fornitore = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value
                produttore = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 3).value
                quantita = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 5).value
                note = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value
                fornitore = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 10).value
                padre = "E_" & Replace(TextBox6.Text & "-" & Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value, " ", "")

                Dim materiale As String = ""
                Dim trattamento As String = ""
                Dim disegno As String = figlio
                Dim creato_da As String = "G.T."
                Dim modificato_da As String = "G.T."
                Dim revisione As Integer = 0
                Dim ubicazione As String = ""
                Dim codice_a_barre As String = ""

                If check_anagrafica_gia_importata(figlio) = "NO" Then
                    importazione_PDM_SAP_ANAGRAFICA(figlio, Replace(Descrizione, "'", " "), Replace(note, "'", " "), revisione, 100, "G.T.", "G.T.", Replace(um, "'", " "), Replace(materiale, "'", " "), Replace(trattamento, "'", " "), Replace(materiale, "'", " "), Replace(trattamento, "'", " "), Replace(disegno, "'", " "), fornitore)
                End If
                importazione_PDM_SAP_distinta(padre, figlio, trova_posizione(padre), Replace(quantita, ",", "."), revisione, Descrizione, note, creato_da, modificato_da, um, materiale, trattamento, ubicazione, codice_a_barre, disegno)
            End If
                par_riga_inizio = par_riga_inizio + 1
        Loop


    End Sub

    Sub trova_dato_da_excel_pEr_importazionE(par_percorso_file As String, par_nome_foglio As String, par_riga_inizio As Integer, par_riga_fine As Integer)

        Dim figlio As String
        Dim padre As String
        Dim Descrizione As String
        'Dim desc_supp As String
        Dim fantasma As String
        Dim um As String
        Dim fornitore As String


        Dim posizione As String
        Dim quantita As String
        Dim note As String


        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True


        Do While par_riga_inizio <= par_riga_fine


            If Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value <> Nothing Then
                figlio = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value
                padre = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 3).value
                Descrizione = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value
                fantasma = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 5).value
                um = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 6).value

                quantita = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 7).value
                posizione = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 8).value
                note = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 9).value
                fornitore = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 10).value

                Dim materiale As String = ""
                Dim trattamento As String = ""
                Dim disegno As String = figlio
                Dim creato_da As String = "G.T."
                Dim modificato_da As String = "G.T."
                Dim revisione As Integer = 99
                Dim ubicazione As String = ""
                Dim codice_a_barre As String = ""


                importazione_PDM_SAP_ANAGRAFICA(figlio, Replace(Descrizione, "'", " "), Replace(note, "'", " "), revisione, 100, "G.T.", "G.T.", Replace(um, "'", " "), Replace(materiale, "'", " "), Replace(trattamento, "'", " "), Replace(materiale, "'", " "), Replace(trattamento, "'", " "), Replace(disegno, "'", " "), fornitore)
                importazione_PDM_SAP_distinta(padre, figlio, posizione, Replace(quantita, ",", "."), revisione, Descrizione, note, creato_da, modificato_da, um, materiale, trattamento, ubicazione, codice_a_barre, disegno)
            End If
            par_riga_inizio = par_riga_inizio + 1
        Loop


    End Sub

    Sub trova_dato_da_excel_pEr_importazionE_da_pdm(par_percorso_file As String, par_nome_foglio As String, par_riga_inizio As Integer, par_riga_fine As Integer)

        Dim figlio As String
        Dim padre As String
        Dim Descrizione As String
        Dim desc_supp As String
        Dim fantasma As String
        Dim um As String
        Dim fornitore As String


        Dim posizione As String
        Dim quantita As String
        Dim note As String
        Dim ubicazione As String = ""

        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True


        Do While par_riga_inizio <= par_riga_fine


            If Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value <> Nothing Then


                figlio = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value

                padre = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 13).value
                    Descrizione = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 3).value
                    desc_supp = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 5).value
                    'fantasma = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 5).value
                    um = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 10).value

                    quantita = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 8).value
                    posizione = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 6).value
                    'note = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 9).value
                    ubicazione = ""

                    Dim materiale As String = ""
                    Dim trattamento As String = ""
                    Dim disegno As String = figlio
                    Dim creato_da As String = "G.T."
                    Dim modificato_da As String = "G.T."
                    Dim revisione As Integer = 99

                    Dim codice_a_barre As String = ""


                    importazione_PDM_SAP_ANAGRAFICA(figlio, Replace(Descrizione, "'", " "), Replace(note, "'", " "), revisione, 100, "G.T.", "G.T.", Replace(um, "'", " "), Replace(materiale, "'", " "), Replace(trattamento, "'", " "), Replace(materiale, "'", " "), Replace(trattamento, "'", " "), Replace(disegno, "'", " "), fornitore)
                    importazione_PDM_SAP_distinta(padre, figlio, posizione, Replace(quantita, ",", "."), revisione, Descrizione, note, creato_da, modificato_da, um, materiale, trattamento, ubicazione, codice_a_barre, disegno)
                End If
                par_riga_inizio = par_riga_inizio + 1
        Loop


    End Sub




    Private Function check_anagrafica_gia_importata(par_codice As String)



        Dim esiste As String = "NO"
        Dim Cnn4 As New SqlConnection
        Cnn4.ConnectionString = Homepage.sap_tirelli
        Cnn4.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn4

        CMD_SAP_2.CommandText = "select t0.codice_brb
from [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_ANAGRAFICA t0
where t0.codice_brb='" & par_codice & "'

"
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            esiste = "SI"
        Else
            esiste = "NO"
        End If


        Cnn4.Close()
        Return esiste
    End Function

    Private Function trova_posizione(par_padre As String)

        Dim posizione As Integer = 0
        Dim Cnn4 As New SqlConnection
        Cnn4.ConnectionString = Homepage.sap_tirelli
        Cnn4.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn4

        CMD_SAP_2.CommandText = "select coalesce(sum(case when t0.codice_padre is null then 0 else 1 end ),0) as 'N'
from [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_DISTINTA t0
where t0.codice_padre='" & par_padre & "'


"
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            posizione = cmd_SAP_reader_2("N")
        Else
            posizione = 0
        End If


        Cnn4.Close()
        Return posizione
    End Function

    Sub riattiva_esportato_distinta()

        Dim Cnn6 As New SqlConnection

        Cnn6.ConnectionString = homepage.sap_tirelli
        cnn6.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = cnn6
        CMD_SAP_5.CommandText = "update t0 set  t0.[Data_export]=null
      ,t0.[Esportato_distinta]='N'
	 

  FROM [TIRELLI_40].[DBO].Frontiera_PDM_BRB_SAP_DISTINTA t0 

where t0.Esportato_distinta='Y'"
        CMD_SAP_5.ExecuteNonQuery()


        cnn6.Close()

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        delete_frontiera_anagrafica()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        delete_frontiera_distinta()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        riattiva_esportato_distinta()
        riempi_datagridview_distinta()
        MsgBox("Distinte riattivate")
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        delete_frontiera_anagrafica()
        delete_frontiera_distinta()
        If TextBox6.Text = "" Then
            MsgBox("Indicare la commessa")
        Else


            Dim percorso_file As String = ""

            Using openFileDialog1 As New OpenFileDialog()
                openFileDialog1.Title = "Seleziona un file"
                openFileDialog1.Filter = "Tutti i file (*.*)|*.*"

                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    percorso_file = openFileDialog1.FileName
                    ' Puoi fare qualcos'altro con il percorso_file qui
                Else
                    MsgBox("Scegliere un file valido")
                End If
            End Using

            'percorso_file = "C:\Users\giovannitirelli\Desktop\E_M04042.xlsx"
            trova_dato_da_excel_pEr_importazionE_quadro(percorso_file, TextBox7.Text, TextBox8.Text, TextBox9.Text)

            riempi_datagridview_anagrafica()
            riempi_datagridview_distinta()

            Beep()
            MsgBox("Importazione completata")
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        delete_frontiera_anagrafica()
        delete_frontiera_distinta()
        MsgBox("Record cancellati")
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Esportazione_anagrafica()
        Esportazione_distinta_anagrafiche()
        Esportazione_distinta_legami()
        MsgBox("Esportazione effettuata con successo")
        mostra_file()
    End Sub
End Class

