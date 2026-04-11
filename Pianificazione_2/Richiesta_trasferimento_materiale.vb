Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports Word = Microsoft.Office.Interop.Word
Imports System.Data.Odbc

Public Class Richiesta_trasferimento_materiale

    Public docnum_odp As Integer
    Public codice_sap As String
    Public Descrizione_riga As String

    Public magazzino_destinazione As String
    Public magazzino_partenza As String
    Public docentry_odp As Integer
    Public docentry_oc As Integer
    Public linenum_odp As Integer


    Public numero_RT As Integer
    Public docentry_RT As Integer
    Public linenumber_RT As Integer
    Public codice_riga As String
    'Public quantità_trasferibile As Decimal
    Public series_rt As Integer
    Public ref1 As String
    Public stringa_trasferimento As String
    Public documento As String
    Public codicedip As Integer
    Public listtype As Integer
    Public movimento As String
    Public magazzino_riga As String





    Public i As Integer
    Private controllo_rt As Integer
    Private stringa_controllo_rt As String

    Sub Inizializzazione_rt()
        ref1 = docnum_odp
        stringa_trasferimento = "Trasferimento"
        documento = "ODP"
        codicedip = 1
        magazzino_destinazione = "WIP"
        docentry_oc = 0

    End Sub

    Sub riempi_datagridview_rt(par_docnum_odp As Integer)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand
        Dim cmd_SAP_reader_7 As SqlDataReader
        CMD_SAP_7.Connection = cnn

        CMD_SAP_7.CommandText = "
select t10.linenum, t10.visorder, T10.[ItemCode], T10.[ItemName],t10.u_disegno, T10.[PlannedQty], t10.U_PRG_WIP_QtaSpedita , t10.U_PRG_WIP_QtaDaTrasf ,t10.U_PRG_WIP_QtaRichMagAuto, T10.Da_magazzino ,t10.A_Magazzino , t10.onhand, t10.Flaggato,cast(case when t10.onhand>=T10.[U_PRG_WIP_QtaDaTrasf]-t10.U_PRG_WIP_QtaRichMagAuto then T10.[U_PRG_WIP_QtaDaTrasf]-t10.U_PRG_WIP_QtaRichMagAuto else t10.onhand end as decimal) as 'Q_RT', t10.docentry, t10.docnum, t10.u_prg_azs_commessa
from
(
SELECT t1.linenum, t1.visorder, T2.[ItemCode], T2.[ItemName],t2.u_disegno, T1.[PlannedQty], case when T1.[U_PRG_WIP_QtaSpedita] is null then 0 else T1.[U_PRG_WIP_QtaSpedita] end as 'U_PRG_WIP_QtaSpedita' , case when T1.[U_PRG_WIP_QtaDaTrasf] is null then 0 else T1.[U_PRG_WIP_QtaDaTrasf] end as 'U_PRG_WIP_QtaDaTrasf' , T1.[wareHouse] as 'Da_magazzino' ,'WIP' AS 'A_Magazzino' , cast(t3.onhand as decimal) as 'onhand', case when t3.onhand>0 then 'True' else 'false' end as 'Flaggato',
case when T1.[U_PRG_WIP_QtaRichMagAuto] is null then 0 else T1.[U_PRG_WIP_QtaRichMagAuto] end as 'U_PRG_WIP_QtaRichMagAuto' ,t1.docentry,t0.docnum, case when t0.u_prg_azs_commessa is null then '' else t0.u_prg_azs_commessa end as 'u_prg_azs_commessa'
FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
inner join OITM T2 on t2.itemcode=t1.itemcode 
inner join oitw t3 on t3.itemcode=t1.itemcode and t3.whscode=T1.[wareHouse]

WHERE T0.[DocNum] ='" & par_docnum_odp & "'  and  case when T1.[U_PRG_WIP_QtaDaTrasf] is null then 0 else T1.[U_PRG_WIP_QtaDaTrasf] end >0 and T1.[wareHouse]='Ferretto' and (case when T1.[U_PRG_WIP_QtaRichMagAuto] is null then 0 else T1.[U_PRG_WIP_QtaRichMagAuto] end<case when T1.[U_PRG_WIP_QtaDaTrasf] is null then 0 else T1.[U_PRG_WIP_QtaDaTrasf] end or  (T1.[U_PRG_WIP_QtaRichMagAuto] is null and case when T1.[U_PRG_WIP_QtaDaTrasf] is null then 0 else T1.[U_PRG_WIP_QtaDaTrasf] end>0))
)
as t10"

        cmd_SAP_reader_7 = CMD_SAP_7.ExecuteReader

        Do While cmd_SAP_reader_7.Read() = True


            '            DataGridView.Rows.Add(cmd_SAP_reader_7("linenum"), cmd_SAP_reader_7("visorder"), cmd_SAP_reader_7("Flaggato"), cmd_SAP_reader_7("Itemcode"), cmd_SAP_reader_7("Itemname"), cmd_SAP_reader_7("u_disegno"), cmd_SAP_reader_7("PlannedQty"), cmd_SAP_reader_7("U_PRG_WIP_QtaSpedita"), cmd_SAP_reader_7("U_PRG_WIP_QtaDaTrasf"), cmd_SAP_reader_7("U_PRG_WIP_QtaRichMagAuto"), cmd_SAP_reader_7("da_magazzino"), cmd_SAP_reader_7("a_magazzino"), cmd_SAP_reader_7("onhand"), cmd_SAP_reader_7("Q_RT"), cmd_SAP_reader_7("docentry"), "", cmd_SAP_reader_7("docnum"), cmd_SAP_reader_7("u_prg_azs_commessa"))

            DataGridView.Rows.Add(cmd_SAP_reader_7("linenum"), cmd_SAP_reader_7("visorder"), False, cmd_SAP_reader_7("Itemcode"), cmd_SAP_reader_7("Itemname"), cmd_SAP_reader_7("u_disegno"), cmd_SAP_reader_7("PlannedQty"), cmd_SAP_reader_7("U_PRG_WIP_QtaSpedita"), cmd_SAP_reader_7("U_PRG_WIP_QtaDaTrasf"), cmd_SAP_reader_7("U_PRG_WIP_QtaRichMagAuto"), cmd_SAP_reader_7("da_magazzino"), cmd_SAP_reader_7("a_magazzino"), cmd_SAP_reader_7("onhand"), cmd_SAP_reader_7("Q_RT"), cmd_SAP_reader_7("docentry"), "", cmd_SAP_reader_7("docnum"), cmd_SAP_reader_7("u_prg_azs_commessa"))

        Loop
        cmd_SAP_reader_7.Close()
        cnn.Close()


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim contatore As Integer = 0
        i = 0


        Do While i < DataGridView.RowCount - 1

            If DataGridView.Rows(i).Cells(columnName:="FLAG").Value = True Then



                codice_sap = UCase(DataGridView.Rows(i).Cells(columnName:="Codice").Value)
                Descrizione_riga = UCase(DataGridView.Rows(i).Cells(columnName:="Descrizione").Value)
                'quantità_trasferibile = DataGridView.Rows(i).Cells(columnName:="Q_RT").Value
                magazzino_partenza = DataGridView.Rows(i).Cells(columnName:="Da_Magazzino").Value
                Magazzino.Codice_SAP = UCase(DataGridView.Rows(i).Cells(columnName:="Codice").Value)
                docentry_odp = DataGridView.Rows(i).Cells(columnName:="docentry_odp_").Value
                linenum_odp = DataGridView.Rows(i).Cells(columnName:="linenum").Value
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)


                check_RT(codice_sap, magazzino_partenza, docentry_odp, linenum_odp, DataGridView.Rows(i).Cells(columnName:="Q_RT").Value)
                If controllo_rt = 0 Then
                    If contatore = 0 Then
                        TROVA_docnum_RT()
                        TROVA_docentry_RT()
                        trova_serie()
                        Magazzino.Trova_PERIODO_contabile()
                        'Inserisci_documento__RICHIESTA_trasferimento()
                        'Magazzino.stringa_trasferimento = "Richiesta trasferimento materiali con 4.0"
                        Magazzino.Inserisci_documento_richiesta_trasferimento(docentry_RT, numero_RT, "ODP", docnum_odp, 0, 0, "Ferretto", "WIP", docentry_odp, 0, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, "Richiesta trasferimento materiali con 4.0")
                        AGGIUSTA_NUMERATORE()
                        contatore = contatore + 1
                    End If

                    Inserisci_righe_richiesta_trasferimento(DataGridView.Rows(i).Cells(columnName:="Q_RT").Value, Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP).Prezzo_listino_acquisto)
                    aggiorna_quantità_rich_mag_auto(DataGridView.Rows(i).Cells(columnName:="Q_RT").Value)
                    FERRETTO_Insert_into_lists(DataGridView.Rows(i).Cells(columnName:="Docnum").Value, DataGridView.Rows(i).Cells(columnName:="commessa").Value, DataGridView.Rows(i).Cells(columnName:="Q_RT").Value)
                    DataGridView.Rows.RemoveAt(i)
                Else
                    MsgBox(stringa_controllo_rt)
                End If
            End If
            i = i + 1

        Loop

        MsgBox("RICHIESTA TRASFERIMENTO ESEGUITA CON SUCCESSO")
    End Sub

    Sub Inserisci_righe_richiesta_trasferimento(par_quantità_trasferibile As Decimal, par_prezzo_listino_acquisto As String)



        magazzino_destinazione = "WIP"
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli

        cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = cnn
        Cmd_SAP.CommandText = "INSERT INTO WTQ1 (WTQ1.Quantity,WTQ1.DocEntry, WTQ1.LineNum, WTQ1.TargetType, WTQ1.TrgetEntry, WTQ1.BaseRef, WTQ1.BaseType, WTQ1.BaseEntry, WTQ1.BaseLine, WTQ1.LineStatus, WTQ1.ItemCode, WTQ1.Dscription,  WTQ1.ShipDate, WTQ1.OpenQty, WTQ1.Price, WTQ1.Currency, WTQ1.Rate, WTQ1.DiscPrcnt, WTQ1.LineTotal, WTQ1.TotalFrgn, WTQ1.OpenSum, WTQ1.OpenSumFC, WTQ1.VendorNum, WTQ1.SerialNum, WTQ1.WhsCode, WTQ1.SlpCode, WTQ1.Commission, WTQ1.TreeType, WTQ1.AcctCode, WTQ1.TaxStatus, WTQ1.GrossBuyPr, WTQ1.PriceBefDi, WTQ1.DocDate, WTQ1.OpenCreQty, WTQ1.UseBaseUn, WTQ1.SubCatNum, WTQ1.BaseCard, WTQ1.TotalSumSy, WTQ1.OpenSumSys, WTQ1.InvntSttus, WTQ1.OcrCode, WTQ1.Project, WTQ1.CodeBars, WTQ1.VatPrcnt, WTQ1.VatGroup, WTQ1.PriceAfVAT, WTQ1.Height1, WTQ1.Hght1Unit, WTQ1.Height2, WTQ1.Hght2Unit, WTQ1.Width1, WTQ1.Wdth1Unit, WTQ1.Width2, WTQ1.Wdth2Unit, WTQ1.Length1, WTQ1.Len1Unit, WTQ1.length2, WTQ1.Len2Unit, WTQ1.Volume, WTQ1.VolUnit, WTQ1.Weight1, WTQ1.Wght1Unit, WTQ1.Weight2, WTQ1.Wght2Unit, WTQ1.Factor1, WTQ1.Factor2, WTQ1.Factor3, WTQ1.Factor4, WTQ1.PackQty, WTQ1.UpdInvntry, WTQ1.BaseDocNum, WTQ1.BaseAtCard, WTQ1.SWW, WTQ1.VatSum, WTQ1.VatSumFrgn, WTQ1.VatSumSy, WTQ1.FinncPriod, WTQ1.ObjType, WTQ1.BlockNum, WTQ1.ImportLog, WTQ1.DedVatSum, WTQ1.DedVatSumF, WTQ1.DedVatSumS, WTQ1.IsAqcuistn, WTQ1.DistribSum, WTQ1.DstrbSumFC, WTQ1.DstrbSumSC, WTQ1.GrssProfit, WTQ1.GrssProfSC, WTQ1.GrssProfFC, WTQ1.VisOrder, WTQ1.INMPrice, WTQ1.PoTrgNum, WTQ1.PoTrgEntry, WTQ1.DropShip, WTQ1.PoLineNum, WTQ1.Address, WTQ1.TaxCode, WTQ1.TaxType, WTQ1.OrigItem, WTQ1.BackOrdr, WTQ1.FreeTxt, WTQ1.PickStatus, WTQ1.PickOty, WTQ1.PickIdNo, WTQ1.TrnsCode, WTQ1.VatAppld, WTQ1.VatAppldFC, WTQ1.VatAppldSC, WTQ1.BaseQty, WTQ1.BaseOpnQty, WTQ1.VatDscntPr, WTQ1.WtLiable, WTQ1.DeferrTax, WTQ1.EquVatPer, WTQ1.EquVatSum, WTQ1.EquVatSumF, WTQ1.EquVatSumS, WTQ1.LineVat, WTQ1.LineVatlF, WTQ1.LineVatS, WTQ1.unitMsr, WTQ1.NumPerMsr, WTQ1.CEECFlag, WTQ1.ToStock, WTQ1.ToDiff, WTQ1.ExciseAmt, WTQ1.TaxPerUnit, WTQ1.TotInclTax, WTQ1.CountryOrg, WTQ1.StckDstSum, WTQ1.ReleasQtty, WTQ1.LineType, WTQ1.TranType, WTQ1.Text, WTQ1.OwnerCode, WTQ1.StockPrice, WTQ1.ConsumeFCT, WTQ1.LstByDsSum, WTQ1.StckINMPr, WTQ1.LstBINMPr, WTQ1.StckDstFc, WTQ1.StckDstSc, WTQ1.LstByDsFc, WTQ1.LstByDsSc, WTQ1.StockSum, WTQ1.StockSumFc, WTQ1.StockSumSc, WTQ1.StckSumApp, WTQ1.StckAppFc, WTQ1.StckAppSc, WTQ1.ShipToCode, WTQ1.ShipToDesc, WTQ1.StckAppD, WTQ1.StckAppDFC, WTQ1.StckAppDSC, WTQ1.BasePrice, WTQ1.GTotal, WTQ1.GTotalFC, WTQ1.GTotalSC, WTQ1.DistribExp, WTQ1.DescOW, WTQ1.DetailsOW, WTQ1.GrossBase, WTQ1.VatWoDpm, WTQ1.VatWoDpmFc, WTQ1.VatWoDpmSc, WTQ1.CFOPCode, WTQ1.CSTCode, WTQ1.Usage, WTQ1.TaxOnly, WTQ1.WtCalced, WTQ1.QtyToShip, WTQ1.DelivrdQty, WTQ1.OrderedQty, WTQ1.CogsOcrCod, WTQ1.CiOppLineN, WTQ1.CogsAcct, WTQ1.ChgAsmBoMW, WTQ1.ActDelDate, WTQ1.OcrCode2, WTQ1.OcrCode3, WTQ1.OcrCode4, WTQ1.OcrCode5, WTQ1.TaxDistSum, WTQ1.TaxDistSFC, WTQ1.TaxDistSSC, WTQ1.PostTax, WTQ1.Excisable, WTQ1.AssblValue, WTQ1.RG23APart1, WTQ1.RG23APart2, WTQ1.RG23CPart1, WTQ1.RG23CPart2, WTQ1.CogsOcrCo2, WTQ1.CogsOcrCo3, WTQ1.CogsOcrCo4, WTQ1.CogsOcrCo5, WTQ1.LnExcised, WTQ1.LocCode, WTQ1.StockValue, WTQ1.GPTtlBasPr, WTQ1.unitMsr2, WTQ1.NumPerMsr2, WTQ1.SpecPrice, WTQ1.CSTfIPI, WTQ1.CSTfPIS, WTQ1.CSTfCOFINS, WTQ1.ExLineNo, WTQ1.isSrvCall, WTQ1.PQTReqQty, WTQ1.PQTReqDate, WTQ1.PcDocType, WTQ1.PcQuantity, WTQ1.LinManClsd, WTQ1.VatGrpSrc, WTQ1.NoInvtryMv, WTQ1.ActBaseEnt, WTQ1.ActBaseLn, WTQ1.ActBaseNum, WTQ1.OpenRtnQty, WTQ1.AgrNo, WTQ1.AgrLnNum, WTQ1.CredOrigin, WTQ1.Surpluses, WTQ1.DefBreak, WTQ1.Shortages, WTQ1.UomEntry, WTQ1.UomEntry2, WTQ1.UomCode, WTQ1.UomCode2, WTQ1.FromWhsCod, WTQ1.NeedQty, WTQ1.PartRetire, WTQ1.RetireQty, WTQ1.RetireAPC, WTQ1.RetirAPCFC, WTQ1.RetirAPCSC, WTQ1.InvQty, WTQ1.OpenInvQty, WTQ1.EnSetCost, WTQ1.RetCost, WTQ1.Incoterms, WTQ1.TransMod, WTQ1.LineVendor, WTQ1.DistribIS, WTQ1.ISDistrb, WTQ1.ISDistrbFC, WTQ1.ISDistrbSC, WTQ1.IsByPrdct, WTQ1.ItemType, WTQ1.PriceEdit, WTQ1.PrntLnNum, WTQ1.LinePoPrss, WTQ1.FreeChrgBP, WTQ1.TaxRelev, WTQ1.LegalText, WTQ1.ThirdParty, WTQ1.LicTradNum, WTQ1.InvQtyOnly, WTQ1.UnencReasn, WTQ1.ShipFromCo, WTQ1.ShipFromDe, WTQ1.FisrtBin, WTQ1.AllocBinC, WTQ1.ExpType, WTQ1.ExpUUID, WTQ1.ExpOpType, WTQ1.DIOTNat, WTQ1.MYFtype, WTQ1.GPBefDisc, WTQ1.ReturnRsn, WTQ1.ReturnAct, WTQ1.StgSeqNum, WTQ1.StgEntry, WTQ1.StgDesc, WTQ1.ItmTaxType, WTQ1.SacEntry, WTQ1.NCMCode, WTQ1.HsnEntry, WTQ1.OriBAbsEnt, WTQ1.OriBLinNum, WTQ1.OriBDocTyp, WTQ1.IsPrscGood, WTQ1.IsCstmAct, WTQ1.EncryptIV, WTQ1.ExtTaxRate, WTQ1.ExtTaxSum, WTQ1.TaxAmtSrc, WTQ1.ExtTaxSumF, WTQ1.ExtTaxSumS, WTQ1.StdItemId, WTQ1.CommClass, WTQ1.VatExEntry, WTQ1.VatExLN, WTQ1.NatOfTrans, WTQ1.ISDtCryImp, WTQ1.ISDtRgnImp, WTQ1.ISOrCryExp, WTQ1.ISOrRgnExp, WTQ1.NVECode, WTQ1.PoNum, WTQ1.PoItmNum, WTQ1.IndEscala, WTQ1.CESTCode, WTQ1.CtrSealQty, WTQ1.CNJPMan, WTQ1.UFFiscBene, WTQ1.U_BLD_LyID, WTQ1.U_BLD_NCps, WTQ1.U_O01FlagU, WTQ1.U_O01ProAg, WTQ1.U_O01ProCA, WTQ1.U_O01ProCZ, WTQ1.U_O01ProDI, WTQ1.U_O01PrzGr, WTQ1.U_O01ScoIm, WTQ1.U_BnTrian, WTQ1.U_Note, WTQ1.U_TrasMgEM, WTQ1.U_Totval, WTQ1.U_BNIncTrm, WTQ1.U_BNTrnMod, WTQ1.U_TestoDOC, WTQ1.U_QtySup, WTQ1.U_PRG_AZS_OpDocEntry, WTQ1.U_PRG_AZS_OpLineNum, WTQ1.U_TpForn, WTQ1.U_PRG_AZS_DescrAlt, WTQ1.U_PRG_AZS_PrevMPS, WTQ1.U_PRG_AZS_StatoComm, WTQ1.U_Colli, WTQ1.U_PRG_AZS_OcDocEntry, WTQ1.U_PRG_AZS_OcDocNum, WTQ1.U_PRG_AZS_OcLineNum, WTQ1.U_PRG_AZS_OaDocEntry, WTQ1.U_PRG_AZS_OaDocNum, WTQ1.U_PRG_AZS_OaLineNum, WTQ1.U_Datitecncompl, WTQ1.U_UTdatainiz, WTQ1.U_UTfineprog, WTQ1.U_inizioassel, WTQ1.U_Fineassel, WTQ1.U_inizassmecc, WTQ1.U_fineassmecc, WTQ1.U_PRG_AZS_OpDocNum, WTQ1.U_PRG_AZS_Commessa, WTQ1.U_PRG_AZS_NumAtCard, WTQ1.U_PRG_AZS_DataRic, WTQ1.U_PRG_AZS_DataCon, WTQ1.U_PRG_AZS_PrzProForma, WTQ1.U_PRG_CLV_PrzPia, WTQ1.U_PRG_CLV_PrzLav, WTQ1.U_PRG_CVM_DocAssoc, WTQ1.U_B1SYS_Discount, WTQ1.U_B1SYS_Discount_FC, WTQ1.U_B1SYS_Discount_SC, WTQ1.U_B1SYS_DiscountVat, WTQ1.U_B1SYS_DiscountVtFC, WTQ1.U_B1SYS_DiscountVtSC, WTQ1.U_Inizcol, WTQ1.U_Finecol, WTQ1.U_Fineapp, WTQ1.U_mod_macchina, WTQ1.U_Fine_app_MU, WTQ1.U_Inizio_ass_EL, WTQ1.U_Fine_ass_EL, WTQ1.U_Inizioapprovvigionamento, WTQ1.U_DataKOM, WTQ1.U_PListinoAcqu, WTQ1.U_Ultimoprezzodeterminato, WTQ1.U_Migliorprezzo, WTQ1.U_Migliorfornitore, WTQ1.U_Trasferito, WTQ1.U_Datrasferire, WTQ1.U_Almag01, WTQ1.U_AlmagCDS, WTQ1.U_Opportunita, WTQ1.U_Ubicazione, WTQ1.U_O01Sc1, WTQ1.U_O01Sc2, WTQ1.U_O01Sc3, WTQ1.U_O01Sc4, WTQ1.U_O01Sc5, WTQ1.U_O01Sc6, WTQ1.U_Ricarico, WTQ1.U_Prezzoarolbranch, WTQ1.U_Commissione_agente, WTQ1.U_Costo, WTQ1.U_Data_scheda_tecnica, WTQ1.U_Data_clean_order, WTQ1.U_Disegno, WTQ1.U_Produttore, WTQ1.U_Revisione, WTQ1.U_PRG_AZS_UbiDest, WTQ1.U_PRG_AZS_PrjFather, WTQ1.U_PRG_AZS_QtaEvasa, WTQ1.U_PRG_WIP_QtaRichMagAuto, WTQ1.U_PRG_QLT_QCDlnQty, WTQ1.U_PRG_QLT_QCCntQty, WTQ1.U_PRG_QLT_QCNCResE, WTQ1.U_PRG_QLT_QCNCResM, WTQ1.U_PRG_QLT_HasTC, WTQ1.U_PRG_WMS_Exp, WTQ1.U_PRG_WMS_ExpDate, WTQ1.U_PRG_WMS_MdMovQty, WTQ1.U_Coefficiente_vendita, WTQ1.U_Gestito_Ferretto, WTQ1.U_Mag_ferretto)

VALUES (" & par_quantità_trasferibile & "," & docentry_RT & "," & i & ",'-1','','','-1','','0','O','" & codice_sap & "','" & Descrizione_riga & "',GETDATE()," & par_quantità_trasferibile & ",'" & par_prezzo_listino_acquisto & "','EUR','0','0'," & par_quantità_trasferibile & "*" & par_prezzo_listino_acquisto & ",'0'," & par_quantità_trasferibile & "*" & par_prezzo_listino_acquisto & ",'0','','','" & magazzino_destinazione & "','-1','0','N','','','0'," & par_prezzo_listino_acquisto & ",GETDATE()," & par_quantità_trasferibile & ",'N','',''," & par_quantità_trasferibile & "*" & par_prezzo_listino_acquisto & "," & par_quantità_trasferibile & "*" & par_prezzo_listino_acquisto & ",'O','','','','0','','0','0','','0','','0','','0','','0','','0','','0','','0','','0','','1','1','1','1','0','Y','','','','0','0','0','" & Magazzino.trova_absentry() & "','1250000001','','','0','0','0','N','0','0','0','0','0','0','0','" & par_prezzo_listino_acquisto & "','','','N','','','','Y','','','','N','0','','17','0','0','0','4','4','0','N','N','0','0','0','0','0','0','0','PZ','1','S','0','0','0','0','0','','0','0','R','','','','0','','0','0','0','0','0','0','0','0','0','0','0','0','0','','','0','0','0','E','0','0','0','Y','N','N','','0','0','0','','','','N','N','0','0','0','','-1','','','','','','','','0','0','0','Y','','0','','','','','','','','','','','0','0','pz','1','N','','','','','N','0','','-1','0','N','N','N','','','','0','','','','0','0','0','-1','-1','Manuale','Manuale','" & magazzino_partenza & "','N','N','0','0','0','0','" & par_quantità_trasferibile & "','" & par_quantità_trasferibile & "','N','0','0','0','','N','0','0','0','N',4,'N','','N','N','Y','','N','','N','','','','','0','','','','','','0','-1','-1','','','','','','-1','','','','','N','N','','0','0','S','0','0','','','','','47','IT','0','','','','','','N','','0','','','-1','','N','0','0','0','0','0','0','N','','N','0','','','','0','" & docentry_odp & "','" & linenum_odp & "','NO','','','O','','','','','','','','','','','','','','','" & docnum_odp & "','','','','','0','0','0','','0','0','0','0','0','0','','','','','','','','','','0','0','0','0','0','0','0','0','','','','','','','','','','','0','0','','','','','','','','0','0','0','0','X','X','N','N','','0','0','','0')"


        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()

    End Sub



    Sub TROVA_docentry_RT()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand
        Dim cmd_SAP_reader_7 As SqlDataReader
        CMD_SAP_7.Connection = cnn

        CMD_SAP_7.CommandText = "SELECT AutoKey FROM ONNM
WHERE ObjectCode='1250000001'"

        cmd_SAP_reader_7 = CMD_SAP_7.ExecuteReader
        If cmd_SAP_reader_7.Read() = True Then
            docentry_RT = cmd_SAP_reader_7("AutoKey")
        Else

        End If
        cmd_SAP_reader_7.Close()
        cnn.Close()


    End Sub

    Sub TROVA_docnum_RT()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand
        Dim cmd_SAP_reader_7 As SqlDataReader
        CMD_SAP_7.Connection = cnn

        CMD_SAP_7.CommandText = "SELECT max(nextnumber) as 'docnum' FROM nnm1
WHERE ObjectCode='1250000001'"

        cmd_SAP_reader_7 = CMD_SAP_7.ExecuteReader
        If cmd_SAP_reader_7.Read() = True Then
            numero_RT = cmd_SAP_reader_7("docnum")
        Else



        End If
        cmd_SAP_reader_7.Close()
        cnn.Close()


    End Sub

    Sub trova_serie()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand
        Dim cmd_SAP_reader_7 As SqlDataReader
        CMD_SAP_7.Connection = cnn

        CMD_SAP_7.CommandText = "select series
from owtq 
where docentry=" & docentry_RT & "-1 "

        cmd_SAP_reader_7 = CMD_SAP_7.ExecuteReader
        If cmd_SAP_reader_7.Read() = True Then
            series_rt = cmd_SAP_reader_7("series")
        Else



        End If
        cmd_SAP_reader_7.Close()
        cnn.Close()


    End Sub

    Sub AGGIUSTA_NUMERATORE()


        Dim Cnn5 As New SqlConnection
        Cnn5.ConnectionString = homepage.sap_tirelli
        cnn5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = cnn5
        CMD_SAP_5.CommandText = "UPDATE ONNM SET AUTOKEY =AUTOKEY+1 WHERE OBJECTCODE='1250000001'
Update NNM1 SET NEXTNUMBER=NEXTNUMBER+1 WHERE OBJECTCODE='1250000001' And SERIES=" & series_rt & ""
        CMD_SAP_5.ExecuteNonQuery()


        cnn5.Close()

    End Sub



    Sub FERRETTO_Insert_into_lists(par_docnum As Integer, par_commessa As String, par_quantità_trasferibile As Decimal)

        listtype = 0
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli

        cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "INSERT INTO [LISTS] (
                     [recordStatus]
                    ,[recordWritingDate]           
                    ,[plantId]
                    ,[command]
                    ,[listType]
                    ,[listNumber]
                    ,[listDescription]
                    ,[initialStatus]
                    ,[lineNumber]
                    ,[item]
                    ,[requestedQty]
,Hauxtext02
,HAuxInt01
					)
	VALUES (        
                    1
                    ,GETDATE()
                    ,1
                    ,1
                    ," & listtype & "
                    ,concat('RT-','" & docentry_RT & "','-P'," & i + 1 & ")
                    ,concat('OD ','" & par_docnum & "',' ','" & par_commessa & "')
                    ,1
                    ," & i & "
                    ,'" & codice_sap & "'
                    ,'" & par_quantità_trasferibile & "'
                    ,'WIP'
                    ,7494
                    )



--Insert in tabella storico
			INSERT INTO [@PRG_WMS_STORICO]
					   ([Code]
					   ,[Name]
					   ,[U_Time]
					   ,[U_Tipoop]
					   ,[U_ItemCode]
					   ,[U_Qtar]			   
					   ,[U_Causale]
					   ,[U_Ordine]
					   ,[U_OrdDes]
					   ,[U_DocEntry]
					   ,[U_ObjType]
					   ,[U_LineNum]
					   ,[U_WhsCode]
					   ,[U_Dir])
				 SELECT
						CONCAT(FORMAT ( GETDATE(), 'yyyyMMddHHmmssfff') ,right('000' + cast(" & i & " as varchar(3)), 3))
					   ,CONCAT(FORMAT ( GETDATE(), 'yyyyMMddHHmmssfff') ,right('000' + cast(" & i & " as varchar(3)), 3))
					   ,GETDATE()
					   ,'P'
					   ,'" & codice_sap & "'		 			   
					   ,'" & par_quantità_trasferibile & "'		   
					   ,'RT'
					   ,concat('RT-','" & docentry_RT & "','-P'," & i + 1 & ")
					   ,concat('OD ', '" & par_docnum & "',' ','" & par_commessa & "')
							
					   ,'" & docentry_RT & "'
					   ,'1250000001'
					   ," & i & "
					   ,'Ferretto'
					   ,'EXP'
				
		
"
        CMD_SAP.ExecuteNonQuery()

        cnn.Close()
    End Sub

    Sub aggiorna_quantità_rich_mag_auto(par_quantità_trasferibile As Decimal)


        magazzino_destinazione = "WIP"
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli

        cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = cnn
        Cmd_SAP.CommandText = "update wor1 set wor1.u_prg_wip_qtarichmagauto= wor1.u_prg_wip_qtarichmagauto+" & par_quantità_trasferibile & " where wor1.docentry=" & docentry_odp & " and wor1.linenum= " & linenum_odp & ""


        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()

    End Sub

    Sub check_RT(par_itemcode As String, par_magazzino As String, par_docentry_op As Integer, par_linenum_op As Integer, par_quantità_trasferibile As Decimal)

        controllo_rt = 0
        stringa_controllo_rt = ""
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand
        Dim cmd_SAP_reader_7 As SqlDataReader
        CMD_SAP_7.Connection = cnn

        CMD_SAP_7.CommandText = "select t0.itemcode ,t1.onhand,
case when t2.U_PRG_WIP_QtaDaTrasf is null then 0 else t2.U_PRG_WIP_QtaDaTrasf end as 'U_PRG_WIP_QtaDaTrasf' ,
case when t2.U_PRG_WIP_QtaRichMagAuto is null then 0 else t2.U_PRG_WIP_QtaRichMagAuto end as 'U_PRG_WIP_QtaRichMagAuto'  ,
case when t3.docnum is null then '' else t3.docnum end as 'docnum_odp'
from 
oitm t0 inner join oitw t1 on t0.itemcode=t1.itemcode
left join wor1 t2 on t2.docentry=" & par_docentry_op & " and t2.LineNum=" & par_linenum_op & "
left join owor t3 on t3.docentry=t2.docentry

where t0.itemcode='" & par_itemcode & "' and t1.whscode='" & par_magazzino & "'"

        cmd_SAP_reader_7 = CMD_SAP_7.ExecuteReader

        If cmd_SAP_reader_7.Read() = True Then
            If cmd_SAP_reader_7("U_PRG_WIP_QtaDaTrasf") <= 0 Then
                controllo_rt = controllo_rt + 1
                stringa_controllo_rt = "Per codice " & cmd_SAP_reader_7("itemcode") & " ODP " & cmd_SAP_reader_7("docnum_odp") & " Da_trasferire<=0 "

            ElseIf cmd_SAP_reader_7("U_PRG_WIP_QtaRichMagAuto") > cmd_SAP_reader_7("U_PRG_WIP_QtaDaTrasf") Then
                controllo_rt = controllo_rt + 1
                stringa_controllo_rt = "Per codice " & cmd_SAP_reader_7("itemcode") & " ODP " & cmd_SAP_reader_7("docnum_odp") & " QtaRichMagAuto>da trasferire"
            ElseIf par_quantità_trasferibile > cmd_SAP_reader_7("U_PRG_WIP_QtaDaTrasf") - cmd_SAP_reader_7("U_PRG_WIP_QtaRichMagAuto") Then
                controllo_rt = controllo_rt + 1
                stringa_controllo_rt = "Per codice " & cmd_SAP_reader_7("itemcode") & " ODP " & cmd_SAP_reader_7("docnum_odp") & " quantità trasferimento >da trasferire -QtaRichMagAuto"

            ElseIf par_quantità_trasferibile > cmd_SAP_reader_7("onhand") Then
                controllo_rt = controllo_rt + 1
                stringa_controllo_rt = "Per codice " & cmd_SAP_reader_7("itemcode") & " ODP " & cmd_SAP_reader_7("docnum_odp") & " quantità trasferimento >Giacenza"
            End If


        End If
        cmd_SAP_reader_7.Close()
        cnn.Close()


    End Sub


End Class