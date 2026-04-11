Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib

Public Class Sales
    Private Excel As Excel.Application
    Private riga_distinta As Integer
    Private contatore As Integer
    Public percorso As String
    Public Foglio As String
    Public Elenco_dipendenti(1000) As String
    Public Tirelli_owner As String
    Public Tirelli_Salesman As String
    Public quota_collaudo As Integer = 0
    Public LISTINO As String
    Public filtro_tipo As String = ""
    Public listino_ As Boolean = False
    Public n_opportunità As Integer = 0
    Public Filtro_tecnologia_etichettatura As String = ""


    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Cambio_BP.Show()
        Me.Hide()
        Cambio_BP.Owner = Me
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Close()
    End Sub

    Sub apri_excel()
        Excel = CreateObject("Excel.application")
        Excel.Workbooks.Open(percorso)
        Excel.Visible = True
    End Sub


    Sub Scrivi()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT  t3.itemcode, t3.itemname, substring(t3.u_disegno,1,4) as 'Super listino', T1.[Quantity] as 'Quantità', case when t3.u_superlistino='Costificato' then sum(T5.[Quantity]*T6.[Price]) else t3.u_superlistino_vecchio end as 'Costo', t3.frozenfor, CASE WHEN T3.VALIDCOMM IS NULL THEN '' ELSE T3.VALIDCOMM END AS 'VALIDCOMM'
From [TIRELLISRLDB].[dbo].oitt t0 
left join [TIRELLISRLDB].[dbo].itt1 t1 on t0.code=t1.father
left join [TIRELLISRLDB].[dbo].oitm t2 on t2.itemcode=t0.code
left join [TIRELLISRLDB].[dbo].oitm t3 on t3.itemcode=t1.code
left join [TIRELLISRLDB].[dbo].itm1 t4 on t4.itemcode=t3.itemcode and t4.pricelist=2 
LEfT JOIN [TIRELLISRLDB].[dbo].itt1 t5 on t5.father =t4.itemcode
left join [TIRELLISRLDB].[dbo].itm1 t6 on t6.itemcode=t5.code and t6.pricelist=2
 WHERE t0.code ='" & DataGridView_commesse.Rows(riga_distinta).Cells(0).Value & "' and (substring( t1.code ,1,1)='S' or substring( t1.code ,1,1)='B' or substring( t1.code ,1,1)='R') 
group by t3.itemcode,T3.VALIDCOMM, t3.itemname, substring(t3.u_disegno,1,4), T1.[Quantity],t3.u_superlistino,t3.u_superlistino_vecchio,t2.itemcode, t1.visorder, t3.frozenfor
order by t2.itemcode, t1.visorder"
        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read() And Cnn.State = 1

            If Excel.Sheets(Foglio).Cells(contatore, 1).value = Nothing Then

                If cmd_SAP_reader("frozenfor") = "Y" Then
                    Excel.Sheets(Foglio).Cells(contatore, 1).value = cmd_SAP_reader("itemcode")
                    Excel.Sheets(Foglio).Cells(contatore, 2).value = "CODICE ANNULLATO. SOSTITUIRE"

                    Excel.Sheets(Foglio).Cells(contatore, 4).value = cmd_SAP_reader("VALIDCOMM")

                Else
                    Excel.Sheets(Foglio).Cells(contatore, 1).value = cmd_SAP_reader("itemcode")
                    Excel.Sheets(Foglio).Cells(contatore, 2).value = cmd_SAP_reader("itemname")
                    Excel.Sheets(Foglio).Cells(contatore, 3).value = cmd_SAP_reader("Quantità")
                    Excel.Sheets(Foglio).Cells(contatore, 4).value = cmd_SAP_reader("Costo")



                    'Excel.Sheets(Foglio).Cells(contatore, 5).formula = "=C" & contatore & "*D" & contatore
                    'Excel.Sheets(Foglio).Cells(contatore, 6).value = TextBox7.Text
                    'Excel.Sheets(Foglio).Cells(contatore, 7).value = quota_collaudo
                    'Excel.Sheets(Foglio).Cells(contatore, 8).value = TextBox10.Text

                    'Excel.Sheets(Foglio).Cells(contatore, 9).formula = "=C" & contatore & "*D" & contatore & "*F" & contatore & "*(100+G" & contatore & ")/100*H" & contatore

                    Excel.Sheets(Foglio).Cells(contatore, 5).value = quota_collaudo
                    Excel.Sheets(Foglio).Cells(contatore, 6).value = TextBox10.Text



                    Excel.Sheets(Foglio).Cells(contatore, 7).formula = "=C" & contatore & "*D" & contatore & "*F" & contatore & "*(100+E" & contatore & ")/100"
                    Excel.Sheets(Foglio).Cells(contatore, 8).value = LISTINO


                    Excel.Sheets(Foglio).Cells(contatore, 9).formula = "=C" & contatore & "*D" & contatore & "*F" & contatore & "*(100+E" & contatore & ")/100*H" & contatore

                End If
                contatore = contatore + 1



            End If
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()
        TextBox1.Text = contatore + 2

    End Sub


    Sub Scrivi_OPTIONAL()
        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT t0.father, t1.itemname, t0.itemcode, t2.itemname as 'Desc', t0.quantity, case when t3.u_superlistino='Costificato' then t4.price else t3.u_superlistino_vecchio end as 'Costo', t0.comments, t0.category
from [Tirelli_40].[dbo].[SENG_OPTIONAL] t0 
inner join [TIRELLISRLDB].[dbo].oitm t1 on t0.father=t1.itemcode
inner join [TIRELLISRLDB].[dbo].oitm t2 on t2.itemcode=t0.itemcode
left join [TIRELLISRLDB].[dbo].oitm t3 on t3.itemcode=t0.itemcode
left join [TIRELLISRLDB].[dbo].itm1 t4 on t4.itemcode=t3.itemcode AND t4.pricelist=2
where t0.father = '" & DataGridView_commesse.Rows(riga_distinta).Cells(0).Value & "' 

order by t0.father,t0.category"
        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read() And Cnn.State = 1

            If Excel.Sheets(Foglio).Cells(contatore, 1).value = Nothing Then
                Excel.Sheets(Foglio).Cells(contatore, 1).value = cmd_SAP_reader("itemcode")
                Excel.Sheets(Foglio).Cells(contatore, 2).value = cmd_SAP_reader("Desc")
                Excel.Sheets(Foglio).Cells(contatore, 3).value = cmd_SAP_reader("Quantity")
                Excel.Sheets(Foglio).Cells(contatore, 4).value = cmd_SAP_reader("Costo")
                Excel.Sheets(Foglio).Cells(contatore, 5).value = quota_collaudo
                Excel.Sheets(Foglio).Cells(contatore, 6).value = TextBox10.Text
                Excel.Sheets(Foglio).Cells(contatore, 7).formula = "=C" & contatore & "*D" & contatore & "*F" & contatore & "*(100+E" & contatore & ")/100"
                Excel.Sheets(Foglio).Cells(contatore, 8).value = TextBox7.Text
                Excel.Sheets(Foglio).Cells(contatore, 9).formula = "=C" & contatore & "*D" & contatore & "*F" & contatore & "*(100+E" & contatore & ")/100*H" & contatore

                'Excel.Sheets(Foglio).Cells(contatore, 5).formula = "=C" & contatore & "*D" & contatore
                'Excel.Sheets(Foglio).Cells(contatore, 6).value = TextBox7.Text
                'Excel.Sheets(Foglio).Cells(contatore, 7).value = quota_collaudo
                'Excel.Sheets(Foglio).Cells(contatore, 8).value = TextBox10.Text
                'Excel.Sheets(Foglio).Cells(contatore, 9).formula = "=C" & contatore & "*D" & contatore & "*F" & contatore & "*(100+G" & contatore & ")/100*H" & contatore
                'Excel.Sheets(Foglio).Cells(contatore, 10).value = cmd_SAP_reader("comments")
                'Excel.Sheets(Foglio).Cells(contatore, 11).value = cmd_SAP_reader("category")


                Excel.Sheets(Foglio).Cells(contatore, 10).value = cmd_SAP_reader("comments")
                Excel.Sheets(Foglio).Cells(contatore, 11).value = cmd_SAP_reader("category")


                contatore = contatore + 1

            End If
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()
        TextBox1.Text = contatore + 2

    End Sub

    Sub Scrivi_superlistino()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "Select  t20.[Codice SAP], t20.[Nome], T20.[Costo componenti], T20.[Assemblaggio], T20.[Magazzino], T20.[Ufficio tecnico], T20.[Assemblaggio elettrico], T20.[Altra manodopera], T20.[Manodopera totale], CASE WHEN T20.[Costificato]='OLD' THEN T20.[Costo OLD] ELSE T20.[Costo totale] END AS 'Costo totale', t20.[Costificato]
from
(
Select*
from (
Select t0.[Superlistino], t0.[Codice SAP], t0.[Nome], sum(t0.[Costo componenti]) as 'Costo componenti', sum(t0.[Assemblaggio]) as 'Assemblaggio', sum(t0.[Magazzino]) as 'Magazzino', sum(t0.[Ufficio tecnico]) as 'Ufficio tecnico', sum(t0.[Assemblaggio elettrico]) as 'Assemblaggio elettrico', sum(t0.[Altra manodopera]) as 'Altra manodopera', sum(t0.[Manodopera totale]) as 'Manodopera totale', sum(t0.[Costo totale]) as 'Costo totale', t0.[Costificato], t0.[Costo OLD]
from
(
SELECT T2.[ItemCode] as'Codice SAP',T2.[ItemName] as 'Nome',  T2.U_disegno as 'Superlistino', T1.[Code], T1.[Quantity], T3.[Price],  case when (substring( t1.code ,1,1)='C' or substring( t1.code ,1,1)='d' or substring(t1.code ,1,1)='0' or substring(t1.code ,1,1)='F') then T1.[Quantity]*T3.[Price] else 0 end as 'Costo componenti', case when t1.code ='R00525' then T1.[Quantity]*T3.[Price] else 0 end as 'Assemblaggio', case when t1.code ='R00542' then T1.[Quantity]*T3.[Price] else 0 end as 'Magazzino', case when t1.code ='R00529' then T1.[Quantity]*T3.[Price] else 0 end as 'Ufficio tecnico', case when t1.code ='R00530' then T1.[Quantity]*T3.[Price] else 0 end as 'Assemblaggio elettrico', 
case when (substring(t1.code,1,1)='R' and t1.code <>'R00525' and t1.code <>'R00542' and t1.code <>'R00529' and t1.code <>'R00530') then T1.[Quantity]*T3.[Price] else 0 end as 'Altra manodopera', case when substring (t1.code,1,1)='R' then T1.[Quantity]*T3.[Price] else 0 end as 'Manodopera totale' , T1.[Quantity]*T3.[Price] as 'Costo TOTALE',  t2.u_superlistino as 'Costificato', t2.u_superlistino_vecchio as 'Costo OLD'


FROM [TIRELLISRLDB].[dbo].OITT T0  
INNER JOIN [TIRELLISRLDB].[dbo].ITT1 T1 ON T0.[Code] = T1.[Father]
INNER JOIN [TIRELLISRLDB].[dbo].OITM T2 ON T0.[Code] = T2.[ItemCode]
 INNER JOIN [TIRELLISRLDB].[dbo].ITM1 T3 ON T1.[Code] = T3.[ItemCode]
 WHERE substring( t0.code ,1,1)='S' and t2.itemcode>'S0A000' and t3.pricelist='2'
)
as t0

group by
t0.[Superlistino], t0.[Codice SAP], t0.[Nome], t0.[Costificato], t0.[Costo OLD]
)
as t10

union all

Select*
from (
Select t0.[Superlistino], t0.[Codice SAP], t0.[Nome], 0 as 'Costo componenti', 0 as 'Assemblaggio', 0 as 'Magazzino', 0 as 'Ufficio tecnico', 0 as 'Assemblaggio elettrico', 0 as 'Altra manodopera', 0 as 'Manodopera totale', 0 as 'Costo totale', t0.[Costificato], t0.[Costo OLD]
from
(
SELECT T2.[ItemCode] as'Codice SAP',T2.[ItemName] as 'Nome',  T2.[U_disegno] as 'Superlistino',t2.u_superlistino as 'Costificato', t2.u_superlistino_vecchio as 'Costo OLD'



FROM [TIRELLISRLDB].[dbo].OITM T2 
LEFT OUTER JOIN [TIRELLISRLDB].[dbo].OITT T0 ON T0.[Code] = T2.[ItemCode]
WHERE substring(t2.ITEMcode,1,1)='S' and t2.itemcode>'S0A000'  AND T0.code is null
)
as t0

group by
t0.[Superlistino], t0.[Codice SAP], t0.[Nome], t0.[Costificato], t0.[Costo OLD]
)
as t10
)
as t20
LEFT JOIN [TIRELLISRLDB].[dbo].ITM1 T21 ON T21.ITEMCODE=T20.[CODICE SAP]
WHERE T21.PRICELIST='2' and t20.[Codice SAP]='" & DataGridView1.Rows(riga_distinta).Cells(0).Value & "'
order by t20.[Codice SAP]"
        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read() And Cnn.State = 1
            ' If Excel.Sheets(Foglio).Cells(contatore, 1).value = Nothing Then
            Excel.Sheets(Foglio).Cells(contatore, 1).value = cmd_SAP_reader("Codice SAP")
            Excel.Sheets(Foglio).Cells(contatore, 2).value = cmd_SAP_reader("Nome")
            Excel.Sheets(Foglio).Cells(contatore, 3).value = "1"
            Excel.Sheets(Foglio).Cells(contatore, 4).value = cmd_SAP_reader("Costo totale")

            Excel.Sheets(Foglio).Cells(contatore, 5).value = quota_collaudo
            Excel.Sheets(Foglio).Cells(contatore, 6).value = TextBox10.Text
            Excel.Sheets(Foglio).Cells(contatore, 7).formula = "=C" & contatore & "*D" & contatore & "*F" & contatore & "*(100+E" & contatore & ")/100"
            Excel.Sheets(Foglio).Cells(contatore, 8).value = TextBox7.Text
            Excel.Sheets(Foglio).Cells(contatore, 9).formula = "=C" & contatore & "*D" & contatore & "*F" & contatore & "*(100+E" & contatore & ")/100*H" & contatore



            'Excel.Sheets(Foglio).Cells(contatore, 5).formula = "=C" & contatore & "*D" & contatore
            '    Excel.Sheets(Foglio).Cells(contatore, 6).value = TextBox7.Text
            'Excel.Sheets(Foglio).Cells(contatore, 7).value = quota_collaudo
            'Excel.Sheets(Foglio).Cells(contatore, 8).value = TextBox10.Text

            '    Excel.Sheets(Foglio).Cells(contatore, 9).formula = "=C" & contatore & "*D" & contatore & "*F" & contatore & "*(100+G" & contatore & ")/100*H" & contatore

            Excel.Sheets(Foglio).range("A" & contatore & ":L" & contatore).font.color = RGB(255, 100, 10)
            TextBox1.Text = contatore + 1

            ' End If
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()


    End Sub

    Sub inserisci_commesse()

        DataGridView_commesse.Rows.Clear()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "select *
from
(
select '' as 'Codice dipendenti', '' as 'Nome', '' as 'Nome 2'
union all
SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code where t0.active='Y'
)
as t0
order by t0.nome"
        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Lavorazioni_MES.Elenco_dipendenti_MES(Indice) = cmd_SAP_reader("Codice dipendenti")
            Lavorazioni_MES.ComboBox_dipendente.Items.Add(cmd_SAP_reader("Nome"))
            Inventario.ComboBox_DIPENDENTE.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub


    Sub Trova_opportunità()

        If Homepage.ERP_provenienza = "SAP" Then


            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli

            Cnn.Open()

            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader


            CMD_SAP.Connection = Cnn
            CMD_SAP.CommandText = "SELECT T1.[OpprId] as 'OPP Nr°',  T1.[CardName] as 'Nome BP', case when t1.U_clientefinale is null then '' else t1.U_clientefinale end  as 'Cliente finale', case when T1.U_DESTINAZIONE is null then T5.NAME else T1.U_DESTINAZIONE end as 'Destinazione', T1.[U_Descrizioneprogetto] AS 'Identificativo progetto',    t2.firstname+' '+t2.lastname as 'Compilatore', T3.[Slpname] as 'Venditore'
FROM  [TIRELLISRLDB].[dbo].OOPR T1 INNER JOIN OPR1 T0 ON T1.OPPRID=T0.OPPRID
left join [TIRELLI_40].[dbo].OHEM T2 ON t2.empid=t0.owner
left join  [TIRELLISRLDB].[dbo].OSLP T3 ON T3.slpcode =t0.slpcode
LEFT JOIN [TIRELLISRLDB].[dbo].OCRD T4 ON T4.CARDCODE=T1.CARDCODE
LEFT JOIN [TIRELLISRLDB].[dbo].OCRY T5 ON T5.CODE=T4.COUNTRY

WHERE  T1.[OpprId]='" & TextBox4.Text & "'"

            cmd_SAP_reader = CMD_SAP.ExecuteReader
            Dim Indice As Integer
            Indice = 0
            If cmd_SAP_reader.Read() = True Then
                Label1.Text = cmd_SAP_reader("Nome BP")
                Label2.Text = cmd_SAP_reader("Cliente finale")
                Label3.Text = cmd_SAP_reader("Destinazione")
                MsgBox(cmd_SAP_reader("Destinazione"))


            End If
            cmd_SAP_reader.Close()
            Cnn.Close()
        End If
    End Sub

    Sub lISTA_MACCHINE(par_datagridview As DataGridView)

        par_datagridview.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[ITEMCode] as 'Codice articolo', CASE WHEN T0.[ItemName] is null then '' else t0.itemname end as 'Nome articolo' , case when T0.FRGNNAME is null then '' else T0.FRGNNAME end as 'FRGNNAME',
CASE WHEN T0.[PrcrmntMtd] ='B' THEN 'FALSE' ELSE 'TRUE' END AS 'APP'
FROM [TIRELLISRLDB].[dbo].OItM T0 

 WHERE SUBSTRING(T0.ITEMCODE,1,1)='b' and t0.itemcode   Like '%%" & TextBox2.Text & "%%' and t0.itemname  Like '%%" & TextBox3.Text & "%%' 
order by t0.itemname"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            par_datagridview.Rows.Add(cmd_SAP_reader("Codice articolo"), cmd_SAP_reader("Nome articolo"), cmd_SAP_reader("APP"), cmd_SAP_reader("FRGNNAME"))

        Loop


        cmd_SAP_reader.Close()
        Cnn.Close()

        par_datagridview.ClearSelection()
    End Sub

    Sub lISTA_optional()

        DataGridView_optional.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT t0.father, t1.itemname, t0.itemcode, t2.itemname as 'Desc', t0.quantity, t0.comments, t0.category
from [Tirelli_40].[dbo].[SENG_OPTIONAL] t0 
inner join [TIRELLISRLDB].[dbo].oitm t1 on t0.father=t1.itemcode
inner join [TIRELLISRLDB].[dbo].oitm t2 on t2.itemcode=t0.itemcode


order by t0.father,t0.category"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            DataGridView_optional.Rows.Add(cmd_SAP_reader("father"), cmd_SAP_reader("itemname"), cmd_SAP_reader("itemcode"), cmd_SAP_reader("desc"), cmd_SAP_reader("quantity"), cmd_SAP_reader("comments"), cmd_SAP_reader("category"))

        Loop


        cmd_SAP_reader.Close()
        Cnn.Close()
        DataGridView_optional.ClearSelection()
    End Sub

    Sub lISTA_superlistino()

        DataGridView1.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "Select  t20.[Codice SAP], t20.[Nome], T20.[Costo componenti], T20.[Assemblaggio], T20.[Magazzino], T20.[Ufficio tecnico], T20.[Assemblaggio elettrico], T20.[Altra manodopera], T20.[Manodopera totale], CASE WHEN T20.[Costificato]='OLD' THEN T20.[Costo OLD] ELSE T20.[Costo totale] END AS 'Costo totale', t20.[Costificato],t20.FrgnName
from
(
Select*
from (
Select t0.[Superlistino], t0.[Codice SAP], t0.[Nome], sum(t0.[Costo componenti]) as 'Costo componenti', sum(t0.[Assemblaggio]) as 'Assemblaggio', sum(t0.[Magazzino]) as 'Magazzino', sum(t0.[Ufficio tecnico]) as 'Ufficio tecnico', sum(t0.[Assemblaggio elettrico]) as 'Assemblaggio elettrico', sum(t0.[Altra manodopera]) as 'Altra manodopera', sum(t0.[Manodopera totale]) as 'Manodopera totale', sum(t0.[Costo totale]) as 'Costo totale', t0.[Costificato], t0.[Costo OLD],t0.FrgnName
from
(
SELECT T2.[ItemCode] as'Codice SAP',T2.[ItemName] as 'Nome',  T2.U_disegno as 'Superlistino', T1.[Code], T1.[Quantity], T3.[Price],  case when (substring( t1.code ,1,1)='C' or substring( t1.code ,1,1)='d' or substring(t1.code ,1,1)='0' or substring(t1.code ,1,1)='F') then T1.[Quantity]*T3.[Price] else 0 end as 'Costo componenti', case when t1.code ='R00525' then T1.[Quantity]*T3.[Price] else 0 end as 'Assemblaggio', case when t1.code ='R00542' then T1.[Quantity]*T3.[Price] else 0 end as 'Magazzino', case when t1.code ='R00529' then T1.[Quantity]*T3.[Price] else 0 end as 'Ufficio tecnico', case when t1.code ='R00530' then T1.[Quantity]*T3.[Price] else 0 end as 'Assemblaggio elettrico', 
case when (substring(t1.code,1,1)='R' and t1.code <>'R00525' and t1.code <>'R00542' and t1.code <>'R00529' and t1.code <>'R00530') then T1.[Quantity]*T3.[Price] else 0 end as 'Altra manodopera', case when substring (t1.code,1,1)='R' then T1.[Quantity]*T3.[Price] else 0 end as 'Manodopera totale' , T1.[Quantity]*T3.[Price] as 'Costo TOTALE',  t2.u_superlistino as 'Costificato', t2.u_superlistino_vecchio as 'Costo OLD', t2.FrgnName


FROM [TIRELLISRLDB].[dbo].OITT T0  INNER JOIN [TIRELLISRLDB].[dbo].ITT1 T1 ON T0.[Code] = T1.[Father]
INNER JOIN [TIRELLISRLDB].[dbo].OITM T2 ON T0.[Code] = T2.[ItemCode]
INNER JOIN [TIRELLISRLDB].[dbo].ITM1 T3 ON T1.[Code] = T3.[ItemCode]
 WHERE substring( t0.code ,1,1)='S' and t2.itemcode>'S0A000' and t3.pricelist='2' and t2.frozenfor<>'Y' and t2.[itemcode]   Like '%%" & TextBox15.Text & "%%'  and t2.[itemname]   Like '%%" & TextBox14.Text & "%%' 
)
as t0

group by
t0.[Superlistino], t0.[Codice SAP], t0.[Nome], t0.[Costificato], t0.[Costo OLD],t0.FrgnName
)
as t10

union all

Select*
from (
Select t0.[Superlistino], t0.[Codice SAP], t0.[Nome], 0 as 'Costo componenti', 0 as 'Assemblaggio', 0 as 'Magazzino', 0 as 'Ufficio tecnico', 0 as 'Assemblaggio elettrico', 0 as 'Altra manodopera', 0 as 'Manodopera totale', 0 as 'Costo totale', t0.[Costificato], t0.[Costo OLD],t0.FrgnName
from
(
SELECT T2.[ItemCode] as'Codice SAP',T2.[ItemName] as 'Nome',  T2.[U_disegno] as 'Superlistino',t2.u_superlistino as 'Costificato', t2.u_superlistino_vecchio as 'Costo OLD',t2.FrgnName



FROM [TIRELLISRLDB].[dbo].OITM T2 LEFT OUTER JOIN [TIRELLISRLDB].[dbo].OITT T0 ON T0.[Code] = T2.[ItemCode]
WHERE substring(t2.ITEMcode,1,1)='S' and t2.itemcode>'S0A000'  AND T0.code is null and t2.frozenfor<>'Y' and  t2.[itemcode]   Like '%%" & TextBox15.Text & "%%'  and t2.[itemname]   Like '%%" & TextBox14.Text & "%%'
)
as t0

group by
t0.[Superlistino], t0.[Codice SAP], t0.[Nome], t0.[Costificato], t0.[Costo OLD],t0.FrgnName
)
as t10
)
as t20
LEFT JOIN [TIRELLISRLDB].[dbo].ITM1 T21 ON T21.ITEMCODE=T20.[CODICE SAP]
WHERE T21.PRICELIST='2'
order by t20.[Codice SAP]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            DataGridView1.Rows.Add(cmd_SAP_reader("Codice SAP"), cmd_SAP_reader("Nome"), cmd_SAP_reader("Costo totale"), cmd_SAP_reader("FrgnName"))

        Loop


        cmd_SAP_reader.Close()
        Cnn.Close()
        DataGridView1.ClearSelection()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)
        Trova_opportunità()
    End Sub

    Sub filtra()
        Dim i = 0
        Dim parola1 As String
        Dim parola2 As String

        Do While i < DataGridView_commesse.RowCount


            parola1 = DataGridView_commesse.Rows(i).Cells(0).Value
            parola2 = DataGridView_commesse.Rows(i).Cells(1).Value


            If parola1.Contains(UCase(TextBox2.Text)) Then
                DataGridView_commesse.Rows(i).Visible = True
                If parola2.Contains(UCase(TextBox3.Text)) Then
                    DataGridView_commesse.Rows(i).Visible = True


                Else
                    DataGridView_commesse.Rows(i).Visible = False

                End If


            Else
                DataGridView_commesse.Rows(i).Visible = False

            End If


            i = i + 1
        Loop
    End Sub

    Sub filtra_superlistino()
        Dim i = 0
        Dim parola0 As String
        Dim parola1 As String

        Do While i < DataGridView1.RowCount


            parola0 = UCase(DataGridView1.Rows(i).Cells(0).Value)
            parola1 = UCase(DataGridView1.Rows(i).Cells(1).Value)


            If parola0.Contains(UCase(TextBox15.Text)) Then
                DataGridView1.Rows(i).Visible = True
                If parola1.Contains(UCase(TextBox14.Text)) Then
                    DataGridView1.Rows(i).Visible = True


                Else
                    DataGridView1.Rows(i).Visible = False

                End If


            Else
                DataGridView1.Rows(i).Visible = False

            End If


            i = i + 1
        Loop
    End Sub


    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        lISTA_MACCHINE(DataGridView_commesse)
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        lISTA_MACCHINE(DataGridView_commesse)
    End Sub



    Private Sub DataGridView_commesse_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_commesse.CellClick
        Dim inizio As Integer
        If e.RowIndex >= 0 Then

            If DataGridView_commesse.Rows(e.RowIndex).Cells(columnName:="Macchina_prodotta").Value = True Then
                LISTINO = TextBox7.Text
            Else
                LISTINO = TextBox12.Text
            End If

            If percorso = Nothing Or Foglio = Nothing Then
                MsgBox("Selezionare un file Excel")
            Else
                If TextBox10.Text = Nothing Then
                    MsgBox("Selezionare una valuta")

                Else

                    If TextBox7.Text = Nothing Then
                        MsgBox("Selezionare un listino Tirelli")
                    Else

                        If TextBox12.Text = Nothing Then
                            MsgBox("Selezionare un listino macchine acquistate")
                        Else
                            If ComboBox3.SelectedIndex < 0 Then
                                MsgBox("Selezionare un grado di rischio")
                            Else


                                riga_distinta = e.RowIndex

                                If e.RowIndex >= 0 Then

                                    If e.ColumnIndex = 0 Then
                                        If Excel.Sheets(Foglio).Cells(contatore, 1).value = Nothing Then

                                            contatore = TextBox1.Text
                                            Excel.Sheets(Foglio).Cells(contatore, 1).value = DataGridView_commesse.Rows(e.RowIndex).Cells(0).Value
                                            Excel.Sheets(Foglio).Cells(contatore, 1).font.bold = True
                                            Excel.Sheets(Foglio).Cells(contatore, 1).font.color = RGB(0, 0, 255)

                                            Excel.Sheets(Foglio).Cells(contatore, 2).value = DataGridView_commesse.Rows(e.RowIndex).Cells(1).Value
                                            Excel.Sheets(Foglio).Cells(contatore, 2).font.bold = True
                                            Excel.Sheets(Foglio).Cells(contatore, 2).font.color = RGB(0, 0, 255)

                                            Excel.Sheets(Foglio).Cells(contatore, 8).value = "N°"
                                            Excel.Sheets(Foglio).Cells(contatore, 8).font.bold = True
                                            Excel.Sheets(Foglio).Cells(contatore, 8).font.color = RGB(0, 0, 255)

                                            Excel.Sheets(Foglio).Cells(contatore, 9).value = "1"
                                            Excel.Sheets(Foglio).Cells(contatore, 9).font.bold = True
                                            Excel.Sheets(Foglio).Cells(contatore, 9).font.color = RGB(0, 0, 255)


                                            contatore = contatore + 1
                                            'Excel.Sheets(Foglio).Cells(contatore, 1).value = "CODICE"
                                            'Excel.Sheets(Foglio).Cells(contatore, 2).value = "DESC"
                                            'Excel.Sheets(Foglio).Cells(contatore, 3).value = "Q"
                                            'Excel.Sheets(Foglio).Cells(contatore, 4).value = "COSTO"
                                            'Excel.Sheets(Foglio).Cells(contatore, 5).value = "TOT COSTO"
                                            'Excel.Sheets(Foglio).Cells(contatore, 6).value = "LISTINO"
                                            'Excel.Sheets(Foglio).Cells(contatore, 7).value = "Q COLLAUDO"
                                            'Excel.Sheets(Foglio).Cells(contatore, 8).value = "CAMBIO"
                                            'Excel.Sheets(Foglio).Cells(contatore, 9).value = "PREZZO"


                                            Excel.Sheets(Foglio).Cells(contatore, 1).value = "CODICE"
                                            Excel.Sheets(Foglio).Cells(contatore, 2).value = "DESC"
                                            Excel.Sheets(Foglio).Cells(contatore, 3).value = "Q"
                                            Excel.Sheets(Foglio).Cells(contatore, 4).value = "COSTO"
                                            Excel.Sheets(Foglio).Cells(contatore, 5).value = "Q COLLAUDO"
                                            Excel.Sheets(Foglio).Cells(contatore, 6).value = "CAMBIO"
                                            Excel.Sheets(Foglio).Cells(contatore, 7).value = "TOT COSTO"
                                            Excel.Sheets(Foglio).Cells(contatore, 8).value = "LISTINO"
                                            Excel.Sheets(Foglio).Cells(contatore, 9).value = "PREZZO"

                                            Excel.Sheets(Foglio).ROWS(contatore).font.bold = True
                                            contatore = contatore + 1
                                            inizio = contatore
                                            Scrivi()
                                            Excel.Sheets(Foglio).Cells(contatore, 12).FORMULA = "=roundup(SUM(G" & inizio & ":G" & contatore & ")*I" & inizio - 2 & ",0)/H4"
                                            Excel.Sheets(Foglio).Cells(contatore, 13).FORMULA = "=roundup(SUM(I" & inizio & ":I" & contatore & "),0)"
                                            Excel.Sheets(Foglio).Cells(contatore, 14).FORMULA = "=roundup(SUM(I" & inizio & ":I" & contatore & ")*I" & inizio - 2 & ",0)"

                                            Excel.Sheets(Foglio).Cells(contatore, 21).FORMULA = "=roundup(SUM(G" & inizio & ":G" & contatore & ")*I" & inizio - 2 & ",0)*2.65/H4"
                                            contatore = contatore + 1
                                            Excel.Sheets(Foglio).Cells(contatore, 1).value = "OPTIONAL"
                                            Excel.Sheets(Foglio).Cells(contatore, 8).value = "N°"
                                            Excel.Sheets(Foglio).Cells(contatore, 8).font.bold = True


                                            Excel.Sheets(Foglio).Cells(contatore, 9).value = "1"
                                            Excel.Sheets(Foglio).Cells(contatore, 9).font.bold = True

                                            Excel.Sheets(Foglio).ROWS(contatore).font.color = RGB(4, 130, 21)
                                            contatore = contatore + 1
                                            'Excel.Sheets(Foglio).Cells(contatore, 1).value = "CODICE"
                                            'Excel.Sheets(Foglio).Cells(contatore, 2).value = "DESC"
                                            'Excel.Sheets(Foglio).Cells(contatore, 3).value = "Q"
                                            'Excel.Sheets(Foglio).Cells(contatore, 4).value = "COSTO"
                                            'Excel.Sheets(Foglio).Cells(contatore, 5).value = "TOT COSTO"
                                            'Excel.Sheets(Foglio).Cells(contatore, 6).value = "LISTINO"
                                            'Excel.Sheets(Foglio).Cells(contatore, 7).value = "Q COLLAUDO"
                                            'Excel.Sheets(Foglio).Cells(contatore, 8).value = "CAMBIO"
                                            'Excel.Sheets(Foglio).Cells(contatore, 9).value = "PREZZO"

                                            Excel.Sheets(Foglio).Cells(contatore, 1).value = "CODICE"
                                            Excel.Sheets(Foglio).Cells(contatore, 2).value = "DESC"
                                            Excel.Sheets(Foglio).Cells(contatore, 3).value = "Q"
                                            Excel.Sheets(Foglio).Cells(contatore, 4).value = "COSTO"
                                            Excel.Sheets(Foglio).Cells(contatore, 5).value = "Q COLLAUDO"
                                            Excel.Sheets(Foglio).Cells(contatore, 6).value = "CAMBIO"
                                            Excel.Sheets(Foglio).Cells(contatore, 7).value = "TOT COSTO"
                                            Excel.Sheets(Foglio).Cells(contatore, 8).value = "LISTINO"
                                            Excel.Sheets(Foglio).Cells(contatore, 9).value = "PREZZO"

                                            Excel.Sheets(Foglio).ROWS(contatore).font.bold = True
                                            contatore = contatore + 1
                                            inizio = contatore
                                            Scrivi_OPTIONAL()
                                            Try
                                                Excel.Sheets(Foglio).Cells(contatore, 13).FORMULA = "=roundup(SUM(I" & inizio & ":I" & contatore & "),0"

                                            Catch ex As Exception

                                            End Try

                                            Try
                                                Excel.Sheets(Foglio).Cells(contatore, 14).FORMULA = "=roundup(SUM(I" & inizio & ":I" & contatore & ")*I" & inizio - 2 & ",0)"
                                            Catch ex As Exception

                                            End Try

                                            Excel.Sheets(Foglio).range("A" & inizio & ":L" & contatore).font.color = RGB(4, 130, 21)

                                        Else
                                            MsgBox("La prima casella risulta già compilata, controllare la casella RIGA")
                                        End If

                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            percorso = OpenFileDialog1.FileName
            TextBox9.Text = percorso

            apri_excel()
            Foglio = InputBox("Inserire nome foglio")
            TextBox11.Text = Foglio
        End If



    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If ComboBox2.Text = Nothing Then
            MsgBox("Identificare il costificatore")
        Else
            'If TextBox4.Text = Nothing Then
            '    MsgBox("Selezionare un'opportunità")
            ' Else
            If ComboBox1.Text = Nothing Then
                    MsgBox("Selezionare una valuta")
                Else

                    percorso = Homepage.percorso_costificatore_seng & "SuperlistinoVB.xlsm"
                    Foglio = "Offerta"
                    TextBox9.Text = percorso
                    TextBox11.Text = Foglio
                    apri_excel()
                    Excel.Sheets(Foglio).Cells(1, 2).value = TextBox4.Text
                    Excel.Sheets(Foglio).Cells(2, 2).value = Label1.Text
                    Excel.Sheets(Foglio).Cells(3, 2).value = Label2.Text
                    Excel.Sheets(Foglio).Cells(1, 8).value = Label3.Text

                    Excel.Sheets(Foglio).Cells(4, 8).value = TextBox10.Text

                    Excel.Sheets(Foglio).Cells(1, 5).value = Tirelli_owner
                    Excel.Sheets(Foglio).Cells(2, 5).value = Tirelli_Salesman
                    Excel.Sheets(Foglio).Cells(3, 5).value = ComboBox2.Text
                    Excel.Sheets(Foglio).Cells(4, 7).value = ComboBox1.Text

                End If
            End If
        ' End If

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        percorso = TextBox9.Text
        If TextBox9.Text = Nothing Then
            TextBox9.BackColor = Color.OrangeRed
            Button5.BackColor = Color.OrangeRed
            Button2.BackColor = Color.OrangeRed
            Button6.BackColor = Color.OrangeRed
        Else
            TextBox9.BackColor = Color.Lime
            Button5.BackColor = Color.Lime
            Button2.BackColor = Color.Lime
            Button6.BackColor = Color.Lime
        End If
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        Foglio = TextBox11.Text
        If TextBox11.Text = Nothing Then
            TextBox1.BackColor = Color.OrangeRed

        Else
            TextBox11.BackColor = Color.Lime

        End If
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Offerta_budget.Owner = Me
        Offerta_budget.Inserimento_tipo_macchina()
        Offerta_budget.Inserimento_prodotto()
        Offerta_budget.Inserimento_velocita()
        Offerta_budget.lISTA_MACCHINE()
        Offerta_budget.Show()
        Me.Hide()
    End Sub

    Sub Inserimento_dipendenti()
        ComboBox2.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 
left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
where t0.active='Y' and t1.name='SENG' order by T0.[lastName]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            ComboBox2.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        Elenco_dipendenti(Indice) = 37
        ComboBox2.Items.Add("Marsiletti Matteo")
        Indice = Indice + 1
        ComboBox2.Items.Add("Tirelli Giacomo")
        Indice = Indice + 1


        Elenco_dipendenti(Indice) = 1
        ComboBox2.Items.Add("Tirelli Roberto")
        Indice = Indice + 1

        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub 'Inserisco le risorse nella combo box


    Sub Inserimento_sottogruppi_eti(par_combobox As ComboBox)
        par_combobox.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "select sottogruppo from 
[Tirelli_40].[dbo].[Superlistino_codici]
group by sottogruppo order by sottogruppo"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        par_combobox.Items.Add("")
        Do While cmd_SAP_reader.Read()

            par_combobox.Items.Add(cmd_SAP_reader("sottogruppo"))

        Loop






        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter
        lISTA_optional()
    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Enter
        lISTA_superlistino()
    End Sub





    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If percorso = Nothing Or Foglio = Nothing Then
            MsgBox("Selezionare un file Excel")
        Else

            'Me.Hide()
            riga_distinta = e.RowIndex
            Scrivi_superlistino()
            'Me.Show()

        End If
    End Sub



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            contatore = TextBox1.Text
            filtra_superlistino()
        Catch ex As Exception

        End Try

    End Sub

    Sub opportunità()
        If Homepage.ERP_provenienza = "SAP" Then

            Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T1.[OpprId] as 'OPP NR°',   T1.[CardCode] as 'BP CODE', T1.[CardName] as 'BP Name', case when T1.u_CLIENTEFINALE is null then '' else t1.u_clientefinale end AS 'End user',CASE WHEN T1.U_DESTINAZIONE IS NULL THEN T6.NAME ELSE T1.U_DESTINAZIONE END AS 'COUNTRY',  CONCAT (t1.U_clientefinale,case when t1.u_clientefinale  is null then '' else ': ' end,T1.[U_Descrizioneprogetto]) AS 'Description', t0.opendate as 'Insert date', T0.[U_PRIORITA] as 'Priorità', t0.U_layout as 'Layout',   T0.[U_Informazioni] as 'Informations', T0.[Line],  t2.firstname+' '+t2.lastname as 'Tirelli Owner', T3.[Slpname] as 'Tirelli Salesman', t0.U_prg_azs_notelivopp AS 'Percorso cartella',T0.[step_id],  t0.status as 'Status', T0.[MaxSumLoc] as 'Amount', T0.[DocNumber] as 'Document N°' 

FROM OPR1 T0 inner JOIN OOPR T1 ON T0.[OpprId] = T1.[OpprId]
left join [TIRELLI_40].[dbo].OHEM T2 ON t2.EMPID=t1.owner
left join  OSLP T3 ON T3.slpcode =t0.slpcode
left join oost t4 on t0.step_id=t4.stepid
left join OCRD T5 ON T5.CARDCODE=T1.CARDCODE
LEFT JOIN OCRY T6 ON T6.CODE=T5.COUNTRY


WHERE  T0.ObjType<>'22' and t0.step_id<>'11'and t0.step_id<>'10'and t0.step_id<>'9'and t0.step_id<>'6'and t0.step_id<>'1' and t1.opprid='" & TextBox4.Text & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader



        If cmd_SAP_reader.Read() Then
            Label1.Text = cmd_SAP_reader("BP Name")
            Label2.Text = cmd_SAP_reader("End user")
            Label3.Text = cmd_SAP_reader("Country")
            Tirelli_owner = cmd_SAP_reader("Tirelli Owner")
            Tirelli_Salesman = cmd_SAP_reader("Tirelli Salesman")
            TextBox4.BackColor = Color.Lime

        Else
            TextBox4.BackColor = Color.OrangeRed
        End If
        cmd_SAP_reader.Close()
        Cnn.Close()



        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        opportunità()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = Nothing Then
            ComboBox1.BackColor = Color.OrangeRed
        Else
            ComboBox1.BackColor = Color.Lime
        End If
        If ComboBox1.Text = "€" Then
            TextBox10.Text = 1
        Else
            TextBox10.Text = "1.25"
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.Text = Nothing Then
            TextBox7.BackColor = Color.OrangeRed
        Else
            TextBox7.BackColor = Color.Lime

        End If
    End Sub



    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = Nothing Then
            ComboBox2.BackColor = Color.OrangeRed
        Else
            ComboBox2.BackColor = Color.Lime
        End If
    End Sub



    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub



    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        lISTA_superlistino()
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        lISTA_superlistino()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.SelectedIndex = -1 Then
            ComboBox3.BackColor = Color.OrangeRed
        Else
            ComboBox3.BackColor = Color.Lime
        End If


        If ComboBox3.Text = 1 Then
            quota_collaudo = 0
        ElseIf ComboBox3.Text = 2 Then
            quota_collaudo = 5
        ElseIf ComboBox3.Text = 3 Then
            quota_collaudo = 10
        End If
    End Sub



    Private Sub TextBox12_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        If TextBox12.Text = Nothing Then
            TextBox12.BackColor = Color.OrangeRed
        Else
            TextBox12.BackColor = Color.Lime

        End If
    End Sub

    Private Sub DataGridView_commesse_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_commesse.CellContentClick

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        trova_dato_da_excel_pEr_importazionE_macchine(Homepage.percorso_costificatore_seng & "SuperlistinoVB.xlsm", "Modelli macchine", 2)

    End Sub

    Sub trova_dato_da_excel_pEr_importazionE_macchine(par_percorso_file As String, par_nome_foglio As String, par_riga_inizio As Integer)

        Dim colonna1 As String
        Dim colonna2 As String
        Dim colonna3 As String
        Dim colonna4 As String
        Dim colonna5 As String
        Dim colonna6 As String
        Dim colonna7 As String
        Dim colonna8 As String
        Dim colonna9 As String
        Dim colonna10 As String
        Dim colonna11 As String


        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True

        delete_modelli_macchina()

        Dim contatore As Integer = 1
        Do While Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value <> ""


            If Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value <> Nothing Then
                colonna1 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value
                colonna2 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value
                colonna3 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 3).value
                colonna4 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value
                colonna5 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 5).value
                colonna6 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 6).value
                colonna7 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 7).value
                colonna8 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 8).value
                colonna9 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 9).value
                colonna10 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 10).value
                colonna11 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 11).value


                insert_into_modelli_macchina(contatore, colonna1, colonna2, colonna3, colonna4, colonna5, colonna6, colonna7, colonna8, colonna9, colonna10, colonna11)

            End If
            contatore += 1
            par_riga_inizio = par_riga_inizio + 1
        Loop
        Beep()
        MsgBox("Importazione effettuata con successo")


    End Sub

    Sub trova_dato_da_excel_pEr_importazionE_codici(par_percorso_file As String, par_nome_foglio As String, par_riga_inizio As Integer)

        Dim colonna1 As String
        Dim colonna2 As String
        Dim colonna3 As String
        Dim colonna4 As String
        Dim colonna5 As String
        Dim colonna6 As String
        Dim colonna7 As String
        Dim colonna8 As String
        Dim colonna9 As String
        Dim colonna10 As String
        Dim colonna11 As String
        Dim colonna12 As String
        Dim colonna13 As String
        Dim colonna14 As String
        Dim colonna15 As String
        Dim colonna18 As String



        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True

        ' delete_modelli_macchina()
        '  delete_codici_superlistino()

        Dim contatore As Integer = 1
        Do While Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value <> ""


            If Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value <> Nothing Then
                colonna1 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value
                colonna2 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value
                colonna3 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 3).value
                colonna4 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value
                colonna5 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 5).value
                colonna6 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 6).value
                colonna7 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 7).value
                colonna8 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 8).value
                colonna9 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 9).value
                colonna10 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 10).value
                colonna11 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 11).value
                colonna12 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 12).value
                colonna13 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 13).value
                colonna14 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 14).value
                colonna15 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 15).value


                colonna18 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 18).value


                'insert_into_modelli_macchina(colonna1, colonna2, colonna3, colonna4, colonna5, colonna6, colonna7, colonna8, colonna9, colonna10, colonna11)
                insert_into_Superlistino_codici(contatore, colonna1, colonna2, colonna3, colonna4, colonna5, colonna6, colonna7, colonna8, colonna9, colonna10, colonna11, colonna12, colonna13, colonna14, colonna15, colonna18)
                ' Update(colonna1, colonna2, colonna3, colonna4, colonna5)
            End If
            contatore += 1
            par_riga_inizio = par_riga_inizio + 1
        Loop
        Beep()
        MsgBox("Importazione effettuata con successo")


    End Sub

    Sub trova_dato_da_excel_pEr_importazionE_distinte(par_percorso_file As String, par_nome_foglio As String, par_riga_inizio As Integer)

        Dim colonna1 As String
        Dim colonna2 As String
        Dim colonna3 As String
        Dim colonna4 As String
        Dim colonna5 As String
        Dim colonna6 As String
        Dim colonna7 As String
        Dim colonna8 As String
        Dim colonna9 As String
        Dim colonna10 As String
        Dim colonna11 As String


        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True


        delete_distinte_superlistino()

        Do While Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value <> ""


            If Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value <> Nothing Then
                colonna1 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value
                colonna2 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value
                colonna3 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 3).value
                colonna4 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value
                colonna5 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 5).value
                colonna6 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 6).value
                colonna7 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 7).value
                colonna8 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 8).value
                colonna9 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 9).value
                colonna10 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 10).value
                colonna11 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 11).value


                'insert_into_modelli_macchina(colonna1, colonna2, colonna3, colonna4, colonna5, colonna6, colonna7, colonna8, colonna9, colonna10, colonna11)
                insert_into_Superlistino_distinte(colonna1, colonna2, colonna3, colonna4, colonna5, colonna6, colonna7, colonna8, colonna9, colonna10, colonna11)
                ' Update(colonna1, colonna2, colonna3, colonna4, colonna5)
            End If
            par_riga_inizio = par_riga_inizio + 1
        Loop
        Beep()
        MsgBox("Importazione effettuata con successo")


    End Sub

    Sub delete_modelli_macchina()



        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1
        Cmd_SAP.CommandText = "

DELETE [Tirelli_40].[dbo].[Modelli_macchine]
         
          "

        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()

    End Sub

    Sub delete_codici_superlistino()



        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1
        Cmd_SAP.CommandText = "

DELETE [Tirelli_40].[dbo].[Superlistino_codici]
          "

        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()

    End Sub

    Sub delete_distinte_superlistino()



        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1
        Cmd_SAP.CommandText = "

DELETE [Tirelli_40].[dbo].[Superlistino_Distinte]
          "

        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()

    End Sub

    Sub insert_into_modelli_macchina(par_Contatore As Integer, par_Colonna_1 As String, par_Colonna_2 As String, par_Colonna_3 As String, par_Colonna_4 As String, par_colonna_5 As String, par_colonna_6 As String, par_colonna_7 As String, par_colonna_8 As String, par_colonna_9 As String, par_colonna_10 As String, par_colonna_11 As String)

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
        Cmd_SAP.CommandText = "



INSERT INTO [Tirelli_40].[dbo].[Modelli_macchine]
           (id,
[Codice]
,[Tipo_macchina]
,[Tipo]
           ,[Nome]
           ,[Serie]
           ,[Diametro_primitivo]
           ,[N_piattelli]
           ,[N_baie])

     VALUES
           (" & par_Contatore & ",'" & par_Colonna_2 & "'
           ,'" & par_Colonna_1 & "'
,'" & par_Colonna_3 & "'
           ,'" & par_Colonna_4 & "'
           ,'" & par_colonna_7 & "'
           ," & par_colonna_8 & "
           ," & par_colonna_9 & "
           ," & par_colonna_10 & ")

          "

        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()

    End Sub

    Sub insert_into_Superlistino_codici(par_contatore As Integer, par_Colonna_1 As String, par_Colonna_2 As String, par_Colonna_3 As String, par_Colonna_4 As String, par_colonna_5 As String, par_colonna_6 As String, par_colonna_7 As String, par_colonna_8 As String, par_colonna_9 As String, par_colonna_10 As String, par_colonna_11 As String, par_colonna_12 As String, par_colonna_13 As String, par_colonna_14 As String, par_colonna_15 As String, par_colonna_18 As String)

        par_Colonna_1 = Replace(par_Colonna_1, "'", " ")
        par_Colonna_2 = Replace(par_Colonna_2, "'", " ")
        par_Colonna_3 = Replace(par_Colonna_3, "'", " ")
        par_Colonna_4 = Replace(par_Colonna_4, "'", " ")
        par_colonna_5 = Replace(par_colonna_5, "'", " ")
        par_colonna_6 = Replace(par_colonna_6, ",", ".")
        par_colonna_7 = Replace(par_colonna_7, ",", ".")
        par_colonna_7 = Replace(par_colonna_7, "'", " ")
        par_colonna_8 = Replace(par_colonna_8, ",", ".")
        par_colonna_9 = Replace(par_colonna_9, ",", ".")
        par_colonna_10 = Replace(par_colonna_10, ",", ".")
        par_colonna_11 = Replace(par_colonna_11, ",", ".")
        par_colonna_18 = Replace(par_colonna_18, ",", ".")



        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1
        Cmd_SAP.CommandText = "



INSERT INTO [Tirelli_40].[dbo].[Superlistino_codici]
           (id,
Codice,
Tipo_macchina,
Descrizione,
[Costo_materiale],
	[Costo],
costificato,
	[ultima_revisione],
[Note],
	[Active],
[sottogruppo],
[ADE],
[STATIC],
[HOT],
[FLEX],
[EU],
[usa])

     VALUES
           (" & par_contatore & ",
           '" & par_Colonna_1 & "'
,'" & par_Colonna_2 & "'
,'" & par_Colonna_3 & "'
 ,'" & Replace(par_Colonna_4, ",", ".") & "'
           ,'" & Replace(par_colonna_5, ",", ".") & "'
           ,'" & par_colonna_18 & "'
,'" & Format(CDate(par_colonna_6), "yyyy-MM-dd") & "'
,'" & par_colonna_7 & "'
,'" & par_colonna_8 & "'
,'" & par_colonna_9 & "'
,'" & par_colonna_10 & "'
,'" & par_colonna_11 & "'
,'" & par_colonna_12 & "'
,'" & par_colonna_13 & "'
,'" & par_colonna_14 & "'
,'" & par_colonna_15 & "')

          "

        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()

    End Sub

    Sub insert_into_Superlistino_distinte(par_Colonna_1 As String, par_Colonna_2 As String, par_Colonna_3 As String, par_Colonna_4 As String, par_colonna_5 As String, par_colonna_6 As String, par_colonna_7 As String, par_colonna_8 As String, par_colonna_9 As String, par_colonna_10 As String, par_colonna_11 As String)

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
        Cmd_SAP.CommandText = "



INSERT INTO [Tirelli_40].[dbo].[Superlistino_Distinte]
           ([Padre]
           ,[Figlio]
,optional
           ,[N_figlio]
,[q]
           ,[Note]
           ,[Ultima_revisione])

     VALUES
           ('" & par_Colonna_1 & "'
,'" & par_Colonna_3 & "'
,'" & par_colonna_5 & "'
           ," & par_Colonna_4 & "
           ," & par_colonna_7 & "
,'" & par_colonna_11 & "'
,'" & Format(CDate(par_colonna_10), "yyyy-MM-dd") & "')

          "

        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        trova_dato_da_excel_pEr_importazionE_codici(Homepage.percorso_costificatore_seng & "SuperlistinoVB.xlsm", "Foglio2", 2)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        trova_dato_da_excel_pEr_importazionE_distinte(Homepage.percorso_costificatore_seng & "SuperlistinoVB.xlsm", "Distinte DEF", 2)
    End Sub

    Private Sub Sales_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'lISTA_MACCHINE(DataGridView_commesse)
        'Me.BackColor = Homepage.colore_sfondo

        '' Se il campo "ADE" è presente nella CheckedListBox1, lo seleziona
        'Dim index As Integer = CheckedListBox1.Items.IndexOf("ADE")
        'If index <> -1 Then
        '    CheckedListBox1.SetItemChecked(index, True)
        'End If
    End Sub


    Private Sub tabpage4_Click(sender As Object, e As EventArgs) Handles TabPage4.Enter
        filtra_macchine()
    End Sub

    Private Sub tabpage5_Click(sender As Object, e As EventArgs) Handles TabPage5.Enter
        filtra_datagridview()
        Inserimento_sottogruppi_eti(ComboBox5)

    End Sub

    Sub filtra_datagridview()
        riempi_codici_etichettatrici(DataGridView2, TextBox24.Text, TextBox23.Text, ComboBox5.Text, ComboBox4.Text)
    End Sub
    Sub filtra_macchine()
        riempi_datagridivew_macchine_new(DataGridView4, TextBox13.Text, TextBox16.Text, TextBox20.Text, TextBox17.Text, TextBox18.Text, TextBox19.Text, filtro_tipo, RadioButton1.Checked, RadioButton2.Checked, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked)
    End Sub



    Sub riempi_datagridivew_macchine_new(par_datagridview As DataGridView, par_codice As String, par_descrizione As String, par_produttore As String, par_diametro As String, par_piattelli As String, par_baie As String, par_filtro_tipo As String, par_eu As Boolean, par_usa As Boolean, par_ade As Boolean, par_static As Boolean, par_hot As Boolean)

        Dim par_filtro_codice As String
        Dim par_filtro_descrizione As String
        Dim par_filtro_produttore As String
        Dim par_filtro_diametro As String
        Dim par_filtro_piattelli As String
        Dim par_filtro_baie As String
        Dim par_filtro_ade As String
        Dim par_filtro_static As String
        Dim par_filtro_hot As String


        If par_codice = "" Then
            par_filtro_codice = ""
        Else
            par_filtro_codice = " and t0.codice  Like '%%" & par_codice & "%%' "

        End If
        If par_descrizione = "" Then
            par_filtro_descrizione = ""
        Else
            par_filtro_descrizione = " and t0.nome  Like '%%" & par_descrizione & "%%' "

        End If
        If par_diametro = "" Then
            par_filtro_diametro = ""
        Else
            par_filtro_diametro = " and t0.diametro_primitivo  = '" & par_diametro & "' "

        End If
        If par_piattelli = "" Then
            par_filtro_piattelli = ""
        Else
            par_filtro_piattelli = " and t0.n_piattelli   = '" & par_piattelli & "' "

        End If
        If par_baie = "" Then
            par_filtro_baie = ""
        Else
            par_filtro_baie = " and t0.n_baie  = '" & par_baie & "' "

        End If

        If par_ade = True Then
            par_filtro_ade = " and t0.ade='Y'"
        Else
            par_filtro_ade = ""
        End If

        If par_static = True Then
            par_filtro_static = " and t0.static='Y'"
        Else
            par_filtro_static = ""
        End If

        If par_hot = True Then
            par_filtro_hot = " and t0.hot='Y'"
        Else
            par_filtro_hot = ""
        End If

        If par_produttore = "" Then
            par_filtro_produttore = ""
        Else
            par_filtro_produttore = " and coalesce(t0.made_in,'') Like '%%" & par_produttore & "%%'"
        End If



        Dim Cnn1 As New SqlConnection
        par_datagridview.Rows.Clear()
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader



        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "
SELECT  t0.[ID]
      ,t0.[Codice]
      ,t0.[Tipo]
      ,t0.[Nome]
      ,t0.[Serie]
      ,t0.[Diametro_primitivo]
      ,t0.[N_piattelli]
      ,t0.[N_baie]
,coalesce(t0.made_in,'') as 'Produttore'
  
	  
  FROM [Tirelli_40].[dbo].[Modelli_macchine] t0 
where 0=0 " & par_filtro_codice & par_filtro_descrizione & par_filtro_baie & par_filtro_diametro & par_filtro_piattelli & par_filtro_ade & par_filtro_static & par_filtro_hot & par_filtro_produttore & "

  order by t0.id, t0.tipo_macchina, t0.tipo, t0.nome
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()
            par_datagridview.Rows.Add(
        cmd_SAP_reader_2("Codice"),
        cmd_SAP_reader_2("Nome"),
        cmd_SAP_reader_2("Produttore"),
        cmd_SAP_reader_2("Diametro_primitivo"),
        cmd_SAP_reader_2("N_piattelli"),
        cmd_SAP_reader_2("N_baie"), trova_presenza_distinta(cmd_SAP_reader_2("Codice")))

        Loop



        Cnn1.Close()
        par_datagridview.ClearSelection()
    End Sub

    Sub riempi_codici_etichettatrici(par_datagridview As DataGridView, par_codice As String, par_descrizione As String, par_sottogruppo As String, par_tipo_macchina As String)

        Dim par_filtro_codice As String
        Dim par_filtro_descrizione As String
        Dim par_filtro_sottogruppo As String



        If par_codice = "" Then
            par_filtro_codice = ""
        Else
            par_filtro_codice = " and t0.codice  Like '%%" & par_codice & "%%' "

        End If
        If par_descrizione = "" Then
            par_filtro_descrizione = ""
        Else
            par_filtro_descrizione = " and t0.descrizione  Like '%%" & par_descrizione & "%%' "

        End If

        If par_sottogruppo = "" Then
            par_filtro_sottogruppo = ""
        Else
            par_filtro_sottogruppo = " and coalesce(t0.sottogruppo,'')  Like '%%" & par_sottogruppo & "%%' "

        End If




        Dim Cnn1 As New SqlConnection
        par_datagridview.Rows.Clear()
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader



        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "
SELECT t0.[ID]
      ,t0.[Codice]
      ,t0.[Tipo_macchina]
      ,t0.[Descrizione]
,coalesce(t0.Sottogruppo,'') as 'Sottogruppo'
      ,t0.[Costo]
      ,t0.[ultima_revisione]
,coalesce(t0.immagine,'') as 'Immagine'
      ,t0.[Note]
      ,t0.[Active]
,t0.costificato
,coalesce(count(t1.figlio),0) as 'Imp'
  FROM [Tirelli_40].[dbo].[Superlistino_codici] t0
left join [Tirelli_40].[dbo].[Superlistino_Distinte] t1 on t0.codice =t1.figlio

where t0.[Active]='Y' " & par_filtro_codice & par_filtro_descrizione & par_filtro_sottogruppo & "
group by t0.[ID]
      ,t0.[Codice]
      ,t0.[Tipo_macchina]
      ,t0.[Descrizione]
,coalesce(t0.Sottogruppo,'') 
      ,t0.[Costo]
      ,t0.[ultima_revisione]
,coalesce(t0.immagine,'')
      ,t0.[Note]
      ,t0.[Active]
,t0.costificato

  order by coalesce(t0.Sottogruppo,''), t0.id
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Dim percorso_immagine As String
        Do While cmd_SAP_reader_2.Read()
            percorso_immagine = Homepage.Percorso_Immagini_TICKETS & "Bianco.jpg"

            If File.Exists(Homepage.Percorso_Immagini_TICKETS & cmd_SAP_reader_2("Immagine")) Then
                percorso_immagine = Homepage.Percorso_Immagini_TICKETS & cmd_SAP_reader_2("Immagine")

            End If
            ' Load the image from file path
            Dim image As Image = Image.FromFile(percorso_immagine)





            ' Imposta l'altezza massima desiderata
            Dim maxHeight As Integer = 60

            ' Calcola la nuova larghezza mantenendo le proporzioni
            Dim scaleFactor As Double = maxHeight / image.Height
            Dim newWidth As Integer = CInt(image.Width * scaleFactor)
            Dim newSize As New Size(newWidth, maxHeight)

            ' Crea l'immagine ridimensionata mantenendo le proporzioni
            Dim smallImage As New Bitmap(image, newSize)




            par_datagridview.Rows.Add(
        cmd_SAP_reader_2("Codice"),
        cmd_SAP_reader_2("Descrizione"),
        cmd_SAP_reader_2("Sottogruppo"),
        cmd_SAP_reader_2("Costo"),
        cmd_SAP_reader_2("Costificato"),
        cmd_SAP_reader_2("Imp"),
        smallImage,
cmd_SAP_reader_2("Note"))


        Loop



        Cnn1.Close()

    End Sub

    Public Function trova_presenza_distinta(par_codice As String)

        Dim presenza As Boolean = False


        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader



        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT TOP 1 [Padre]
      ,[Figlio]
      ,[N_figlio]
      ,[Note]
      ,[Ultima_revisione]
  FROM [Tirelli_40].[dbo].[Superlistino_Distinte]
where padre='" & par_codice & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() Then

            presenza = True
        Else
            presenza = False
        End If



        Cnn1.Close()
        Return presenza
    End Function

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        filtra_macchine()
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        filtra_macchine()
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
        filtra_macchine()
    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged
        filtra_macchine()
    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged
        filtra_macchine()
    End Sub


    Private DPrimColori As New Dictionary(Of String, Color)
    Private Sub DataGridView4_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView4.CellFormatting
        Dim par_datagridview As DataGridView = DataGridView4

        Try
            ' Controllo per la colonna "distinta"
            If par_datagridview.Columns(e.ColumnIndex).Name.Equals("D_prim", StringComparison.OrdinalIgnoreCase) AndAlso e.Value IsNot Nothing Then
                Dim valore As String = e.Value.ToString().Trim()

                ' Assegna un colore se il valore non è già presente
                If Not DPrimColori.ContainsKey(valore) Then
                    DPrimColori(valore) = GeneraColoreCasuale()
                End If

                ' Applica il colore alla cella
                e.CellStyle.BackColor = DPrimColori(valore)
            End If

            ' Se la colonna "distinta" è True, colora "Codice_" e "Descrizione_"
            If par_datagridview.Rows(e.RowIndex).Cells("distinta").Value = True Then
                par_datagridview.Rows(e.RowIndex).Cells("Codice_").Style.BackColor = Color.Lime
                par_datagridview.Rows(e.RowIndex).Cells("Descrizione_").Style.BackColor = Color.Lime
            End If

        Catch ex As Exception
            ' Ignora eventuali errori per celle vuote o fuori intervallo
        End Try
    End Sub

    Sub distinda_etichettatrice_su_excel(par_codice As String, par_riga As Integer, par_optional As String, par_filtro_tipo As String, par_eu As Boolean, par_usa As Boolean, par_Ade As Boolean, par_static As Boolean, par_hot As Boolean)

        Dim filtro_eu As String = ""
        Dim filtro_usa As String = ""
        Dim filtro_ade As String = ""
        Dim filtro_static As String = ""
        Dim filtro_hot As String = ""
        If par_eu = True Then
            filtro_eu = " and t1.eu ='Y'"
        Else
            filtro_eu = ""
        End If

        If par_usa = True Then
            filtro_usa = " and t1.usa ='Y'"
        Else
            filtro_usa = ""
        End If

        If par_Ade = True Then
            filtro_ade = " and t1.ade ='Y'"
        Else
            filtro_ade = ""
        End If

        If par_static = True Then
            filtro_static = " and t1.static ='Y'"
        Else
            filtro_static = ""
        End If

        If par_hot = True Then
            filtro_hot = " and t1.hot ='Y'"
        Else
            filtro_hot = ""
        End If
        If par_filtro_tipo = " AND " Then
            par_filtro_tipo = ""
        End If

        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader



        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT TOP (1000) t0.[Padre]
      ,t0.[Figlio]
      ,t0.[N_figlio]
	  ,T0.OPTIONAL
	  , t1.Descrizione
	  ,t0.q
	  ,t1.Costo
	  , T0.Q*T1.COSTO AS 'Costo_TOT'
	  ,coalesce(t1.Note,'') as 'Note_codice'
      ,t0.[Note]
      ,t0.[Ultima_revisione]
,coalesce(t1.sottogruppo,'') as 'Sottogruppo'
  FROM [Tirelli_40].[dbo].[Superlistino_Distinte] t0
  LEFT JOIN [Tirelli_40].[dbo].[Superlistino_codici] T1 ON T0.FIGLIO=T1.Codice
  where t0.padre='" & par_codice & "' and t0.optional='" & par_optional & "' " & par_filtro_tipo & filtro_eu & filtro_usa & filtro_ade & filtro_static & filtro_hot & "
  order by case when t0.optional='N' then cast(t0.[N_figlio] as varchar) else coalesce(t1.sottogruppo,'') end, t0.[N_figlio] "

        cmd_SAP_reader = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader.Read()


            Excel.Sheets(Foglio).Cells(par_riga, 1).value = cmd_SAP_reader("Figlio")
            Excel.Sheets(Foglio).Cells(par_riga, 2).value = cmd_SAP_reader("Descrizione")
            Excel.Sheets(Foglio).Cells(par_riga, 3).value = cmd_SAP_reader("q")
            Excel.Sheets(Foglio).Cells(par_riga, 4).value = cmd_SAP_reader("Costo")


            Excel.Sheets(Foglio).Cells(par_riga, 5).value = quota_collaudo
            Excel.Sheets(Foglio).Cells(par_riga, 6).value = TextBox10.Text

            Excel.Sheets(Foglio).Cells(par_riga, 7).formula = "=C" & par_riga & "*D" & par_riga & "*F" & par_riga & "*(100+E" & par_riga & ")/100"
            Excel.Sheets(Foglio).Cells(par_riga, 8).value = LISTINO


            Excel.Sheets(Foglio).Cells(par_riga, 9).formula = "=C" & par_riga & "*D" & par_riga & "*F" & par_riga & "*(100+E" & par_riga & ")/100*H" & par_riga
            Excel.Sheets(Foglio).Cells(par_riga, 10).value = cmd_SAP_reader("Sottogruppo")
            Excel.Sheets(Foglio).Cells(par_riga, 11).value = cmd_SAP_reader("Note_codice")

            par_riga = par_riga + 1




        Loop
        cmd_SAP_reader.Close()
        Cnn1.Close()
        TextBox1.Text = par_riga + 2


    End Sub

    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        If e.RowIndex < 0 Then Exit Sub ' Evita header click
        Dim par_datagridview As DataGridView = DataGridView4
        If e.ColumnIndex <> par_datagridview.Columns.IndexOf(Codice_) Then Exit Sub
        If RadioButton1.Checked = False And RadioButton2.Checked = False Then
            MsgBox("Selezionare se la macchina è configurata europea o USA")
            Return
        End If

        Dim inizio As Integer
        LISTINO = TextBox7.Text
        If TextBox4.Text = "" Then
            n_opportunità = 0
        Else
            n_opportunità = TextBox4.Text
        End If

        Dim codice_macchina As String = DataGridView4.Rows(e.RowIndex).Cells("codice_").Value.ToString()
        Dim PAR_CONFIGURATORE As Boolean = True


        If PAR_CONFIGURATORE = False Then

            If Ultima_revisione_costo_macchina(n_opportunità) > 0 Then

                Dim risposta As MsgBoxResult
                risposta = MsgBox("Di questa opportunità esiste già un costo macchina. (REV " & Ultima_revisione_costo_macchina(n_opportunità) & "). Vuoi caricare quel costo macchina?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Conferma")

                If risposta = MsgBoxResult.Yes Then
                    Form_configuratore_vendita.Show()
                    Form_configuratore_vendita.ricrea_form_da_database_new(n_opportunità, Ultima_revisione_costo_macchina(n_opportunità))

                    Return
                End If

            End If

            If ComboBox1.SelectedIndex < 0 Then
                MsgBox("Selezionare una valuta")
                Return
            End If

            If LISTINO = "" Then
                MsgBox("impostare un listino")
                Return
            End If



            If Form_distinta_vendita.trova_tipo_macchina(codice_macchina) = "ETICHETTATRICE BRB" Then


                If CheckBox1.Checked = False And CheckBox2.Checked = False And CheckBox3.Checked = False Then
                    MsgBox("Selezionare la tecnologia di etichettaggio")
                    Return
                End If
            End If
            ' Mostra il form solo se non è già visibile
            If Form_configuratore_vendita.Visible = False Then
                Dim screens = Screen.AllScreens

                ' Se ci sono almeno 2 schermi
                If screens.Length > 1 Then
                    ' Posizioniamo il form sul secondo schermo
                    Form_configuratore_vendita.StartPosition = FormStartPosition.Manual
                    Form_configuratore_vendita.Location = screens(1).WorkingArea.Location
                End If

                Form_configuratore_vendita.Show()
                If ComboBox1.Text = "€" Then
                    Form_configuratore_vendita.RadioButton1.Checked = True
                    Form_configuratore_vendita.RadioButton2.Checked = False

                Else
                    Form_configuratore_vendita.RadioButton2.Checked = True
                    Form_configuratore_vendita.RadioButton1.Checked = False

                End If

            End If
            Form_configuratore_vendita.Txt_DocNum.Text = n_opportunità
            Form_configuratore_vendita.TextBox1.Text = 1
            If Form_configuratore_vendita.TextBox1.Text = "" Then
                Form_configuratore_vendita.informazioni_testata_nuova(n_opportunità)
            End If

            Form_configuratore_vendita.inizializzazione_modulo(codice_macchina, "N", True, RadioButton1.Checked, RadioButton2.Checked, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, CheckBox1.Checked, Replace(TextBox7.Text, ".", ","))

        Else

            If Form_distinta_vendita.trova_tipo_macchina(codice_macchina) = "ETICHETTATRICE BRB" Then


                If CheckBox1.Checked = False And CheckBox2.Checked = False And CheckBox3.Checked = False Then
                    MsgBox("Selezionare la tecnologia di etichettaggio")
                    Return
                End If
            End If

            Try
                Label5.Text = par_datagridview.Rows(e.RowIndex).Cells(columnName:="codice_").Value
            Catch ex As Exception

            End Try
            If e.RowIndex < 0 OrElse e.ColumnIndex <> par_datagridview.Columns.IndexOf(Codice_) Then Exit Sub
            Label5.Text = par_datagridview.Rows(e.RowIndex).Cells(columnName:="codice_").Value
            ' Controlli preliminari per evitare errori
            If String.IsNullOrEmpty(percorso) OrElse String.IsNullOrEmpty(Foglio) Then
                MsgBox("Selezionare un file Excel")
                Exit Sub
            End If
            If String.IsNullOrEmpty(TextBox10.Text) Then
                MsgBox("Selezionare una valuta")
                Exit Sub
            End If
            If String.IsNullOrEmpty(TextBox7.Text) Then
                MsgBox("Selezionare un listino Tirelli")
                Exit Sub
            End If
            If ComboBox3.SelectedIndex < 0 Then
                MsgBox("Selezionare un grado di rischio")
                Exit Sub
            End If

            ' Controllo se la prima cella è già compilata
            If Not String.IsNullOrEmpty(Excel.Sheets(Foglio).Cells(contatore, 1).Value) Then
                MsgBox("La prima casella risulta già compilata, controllare la casella RIGA")
                Exit Sub
            End If

            ' Incremento contatore e riempimento Excel
            contatore = CInt(TextBox1.Text)
            Dim codice As String = par_datagridview.Rows(e.RowIndex).Cells("Codice_").Value
            Dim descrizione As String = par_datagridview.Rows(e.RowIndex).Cells("Descrizione_").Value
            Dim RIGA_1 As Integer

            RIGA_1 = contatore
            ScriviCella(contatore, 1, codice, True, RGB(0, 0, 255))
            ScriviCella(contatore, 2, descrizione, True, RGB(0, 0, 255))
            ScriviCella(contatore, 8, "N°", True, RGB(0, 0, 255))
            ScriviCella(contatore, 9, "1", True, RGB(0, 0, 255))



            contatore += 1
            ScriviIntestazioni(contatore)
            contatore += 1
            inizio = contatore
            distinda_etichettatrice_su_excel(codice, inizio, "N", filtro_tipo, RadioButton1.Checked, RadioButton2.Checked, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked)

            ' Inserimento formule Excel
            ScriviFormula(contatore, 12, "=roundup(SUM(G" & inizio & ":G" & contatore & ")*I" & (inizio - 2) & ",0)/H4")

            'temp
            If listino_ = True Then


                ScriviFormula(RIGA_1, 12, "=roundup(SUM(G" & inizio & ":G" & contatore & ")*I" & (inizio - 2) & ",0)/H4")
                ScriviFormula(RIGA_1, 13, "=roundup(SUM(I" & inizio & ":I" & contatore & "),0)")
                ScriviFormula(RIGA_1, 14, "=roundup(SUM(I" & inizio & ":I" & contatore & ")*I" & (inizio - 2) & ",0)")
            End If

            ScriviFormula(contatore, 13, "=roundup(SUM(I" & inizio & ":I" & contatore & "),0)")
            ScriviFormula(contatore, 14, "=roundup(SUM(I" & inizio & ":I" & contatore & ")*I" & (inizio - 2) & ",0)")

            ScriviFormula(contatore, 21, "=roundup(SUM(G" & inizio & ":G" & contatore & ")*I" & (inizio - 2) & ",0)*2.65/H4")

            contatore += 1
            ScriviCella(contatore, 1, "OPTIONAL", True)
            ScriviCella(contatore, 8, "N°", True, RGB(4, 130, 21))
            ScriviCella(contatore, 9, "1", True, RGB(4, 130, 21))

            contatore += 1
            ScriviIntestazioni(contatore)
            contatore += 1
            inizio = contatore
            contatore += 1

            'optional
            distinda_etichettatrice_su_excel(codice, inizio, "Y", filtro_tipo, RadioButton1.Checked, RadioButton2.Checked, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked)

            ' Inserimento formule Excel con gestione errori
            ScriviFormula(contatore, 13, "=roundup(SUM(I" & inizio & ":I" & contatore & "),0)")
            ScriviFormula(contatore, 14, "=roundup(SUM(I" & inizio & ":I" & contatore & ")*I" & (inizio - 2) & ",0)")

            ' Colorazione range di celle
            Excel.Sheets(Foglio).Range("A" & inizio & ":L" & contatore).Font.Color = RGB(4, 130, 21)
            TextBox1.Text += 1
        End If

    End Sub

    Public Function Ultima_revisione_costo_macchina(par_opportunità As Integer)
        If Homepage.ERP_provenienza = "SAP" Then


            Dim rev As Integer = 0


            Using Cnn As New SqlConnection(Homepage.sap_tirelli)
                Cnn.Open()

                Dim query As String = "SELECT 
      
      t0.[Opportunità]

      ,max(t0.[REV]) as 'REV'
      
  FROM [Tirelli_40].[dbo].[Superlistino_log_costificazioni_testata] t0
left join oopr t1 on t1.opprid=t0.[Opportunità]
left join [TIRELLI_40].[dbo].ohem t2 on t2.empid=t0.utente

        WHERE t0.[Opportunità] = @N_opp 
group by t0.[Opportunità]
        
        "

                Dim cmd As New SqlCommand(query, Cnn)
                cmd.Parameters.AddWithValue("@N_opp", par_opportunità)
                '  cmd.Parameters.AddWithValue("@N_rev", par_n_rev)

                Dim reader As SqlDataReader = cmd.ExecuteReader()

                If reader.Read() Then
                    rev = reader("REV")

                End If
                reader.Close()


                Return rev

            End Using

        End If
    End Function

    ' Funzione per scrivere nelle celle con opzioni di formattazione
    Private Sub ScriviCella(riga As Integer, col As Integer, valore As Object, Optional bold As Boolean = False, Optional color As Integer = -1)
        With Excel.Sheets(Foglio).Cells(riga, col)
            .Value = valore
            .Font.Bold = bold
            If color <> -1 Then .Font.Color = color
        End With



    End Sub

    ' Funzione per scrivere le intestazioni delle colonne
    Private Sub ScriviIntestazioni(riga As Integer)
        Dim intestazioni As String() = {"CODICE", "DESC", "Q", "COSTO", "Q COLLAUDO", "CAMBIO", "TOT COSTO", "LISTINO", "PREZZO", "GRUPPO", "NOTE"}
        For i As Integer = 0 To intestazioni.Length - 1
            ScriviCella(riga, i + 1, intestazioni(i), True)
        Next
    End Sub

    ' Funzione per scrivere formule Excel con gestione errori
    Private Sub ScriviFormula(riga As Integer, col As Integer, formula As String)
        Try
            Excel.Sheets(Foglio).Cells(riga, col).Formula = formula
        Catch ex As Exception
            ' Logga errore se necessario
        End Try
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim par_datagridview As DataGridView = DataGridView2

        Dim inizio As Integer
        LISTINO = TextBox7.Text

        If e.RowIndex >= 0 Then
            Label4.Text = par_datagridview.Rows(e.RowIndex).Cells(columnName:="codice___").Value
            If e.ColumnIndex = par_datagridview.Columns.IndexOf(Codice___) Then

                If percorso = Nothing Or Foglio = Nothing Then
                    MsgBox("Selezionare un file Excel")
                    Return
                End If


                If percorso = Nothing Or Foglio = Nothing Then
                    MsgBox("Selezionare un file Excel")
                    Return
                End If

                If TextBox10.Text = Nothing Then
                    MsgBox("Selezionare una valuta")
                    Return
                End If

                If TextBox7.Text = Nothing Then
                    MsgBox("Selezionare un listino Tirelli")
                    Return
                End If

                If ComboBox3.SelectedIndex < 0 Then
                    MsgBox("Selezionare un grado di rischio")
                    Return
                End If



                'If Excel.Sheets(Foglio).Cells(contatore, 1).value <> Nothing Then
                '    MsgBox("La prima casella risulta già compilata, controllare la casella RIGA")
                '    Return

                'End If

                aggiungi_elemento_singolo_su_excel(par_datagridview.Rows(e.RowIndex).Cells(columnName:="codice___").Value, contatore, ComboBox4.Text)

            End If
        End If
    End Sub

    Sub aggiungi_elemento_singolo_su_excel(par_codice As String, par_riga As Integer, par_tipo_macchina As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT  top 1 t0.[ID]
      ,t0.[Codice]
      ,t0.[Tipo_macchina]
      ,t0.[Descrizione]
      ,t0.[Costo]
      ,t0.[ultima_revisione]
      ,t0.[Note]
      ,t0.[Active]
,coalesce(t0.sottogruppo,'') as 'Sottogruppo'

  FROM [Tirelli_40].[dbo].[Superlistino_codici] t0
where t0.codice='" & par_codice & "' "
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        'and t0.[Tipo_macchina]='" & par_tipo_macchina & "'
        If cmd_SAP_reader.Read() Then
            ' If Excel.Sheets(Foglio).Cells(contatore, 1).value = Nothing Then
            Excel.Sheets(Foglio).Cells(par_riga, 1).value = cmd_SAP_reader("Codice")
            Excel.Sheets(Foglio).Cells(par_riga, 2).value = cmd_SAP_reader("Descrizione")
            Excel.Sheets(Foglio).Cells(par_riga, 3).value = "1"
            Excel.Sheets(Foglio).Cells(par_riga, 4).value = cmd_SAP_reader("Costo")

            Excel.Sheets(Foglio).Cells(par_riga, 5).value = quota_collaudo
            Excel.Sheets(Foglio).Cells(par_riga, 6).value = TextBox10.Text
            Excel.Sheets(Foglio).Cells(par_riga, 7).formula = "=C" & par_riga & "*D" & par_riga & "*F" & par_riga & "*(100+E" & par_riga & ")/100"
            Excel.Sheets(Foglio).Cells(par_riga, 8).value = TextBox7.Text
            Excel.Sheets(Foglio).Cells(par_riga, 9).formula = "=C" & par_riga & "*D" & par_riga & "*F" & par_riga & "*(100+E" & par_riga & ")/100*H" & par_riga
            Excel.Sheets(Foglio).Cells(par_riga, 10).value = cmd_SAP_reader("Sottogruppo")
            Excel.Sheets(Foglio).Cells(par_riga, 11).value = cmd_SAP_reader("Note")

            Excel.Sheets(Foglio).range("A" & par_riga & ":L" & par_riga).font.color = RGB(255, 100, 10)



        End If
        TextBox1.Text = par_riga + 1
        cmd_SAP_reader.Close()
        Cnn.Close()


    End Sub

    Private Sub DataGridView2_CellContextMenuStripChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContextMenuStripChanged

    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged
        filtra_datagridview()
    End Sub

    Private Sub TextBox23_TextChanged(sender As Object, e As EventArgs) Handles TextBox23.TextChanged
        filtra_datagridview()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        filtra_datagridview()
    End Sub










    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        filtra_datagridview()
    End Sub

    ' Dizionario per tenere traccia dei colori assegnati ai diversi valori di "sottogruppo"

    ' Dizionario per assegnare colori univoci ai valori di "sottogruppo"
    Private SottogruppoColori As New Dictionary(Of String, Color)
    Private Shared rnd As New Random() ' Random statico per evitare duplicati

    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        ' Verifica che la colonna sia quella desiderata
        If DataGridView2.Columns(e.ColumnIndex).Name.Equals("Sottogruppo", StringComparison.OrdinalIgnoreCase) AndAlso e.Value IsNot Nothing Then
            Dim valore As String = e.Value.ToString().Trim()

            ' Assegna un colore se il valore non è già presente
            If Not SottogruppoColori.ContainsKey(valore) Then
                SottogruppoColori(valore) = GeneraColoreCasuale()
            End If

            ' Applica il colore alla cella
            e.CellStyle.BackColor = SottogruppoColori(valore)
        End If

        If DataGridView2.Columns(e.ColumnIndex).Name = "Costificato" Then
            If e.Value IsNot Nothing AndAlso Not IsDBNull(e.Value) Then
                Select Case e.Value.ToString().ToUpper()
                    Case "Y"
                        e.CellStyle.BackColor = Color.LightGreen
                    Case "S"
                        e.CellStyle.BackColor = Color.Yellow
                    Case Else
                        e.CellStyle.BackColor = Color.LightCoral
                End Select
            Else
                e.CellStyle.BackColor = Color.LightCoral
            End If
        End If
    End Sub

    ' Funzione per generare un colore casuale evitando colori troppo scuri
    Private Function GeneraColoreCasuale() As Color
        Return Color.FromArgb(255, rnd.Next(120, 255), rnd.Next(120, 255), rnd.Next(120, 255))
    End Function

    ' Rinfresca i colori dopo il caricamento dei dati
    Private Sub DataGridView2_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles DataGridView2.DataBindingComplete
        DataGridView2.Invalidate() ' Forza il refresh per assicurarsi che i colori vengano applicati correttamente
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)
        Form_Codici_vendita.Show()
    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs)




        Form_Codici_vendita.Show()
        Form_Codici_vendita.inizializza_form(Label4.Text)
    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs)
        Form_distinta_vendita.Show()
        Form_distinta_vendita.inizializza_form(Label5.Text)
    End Sub

    Private Sub GestisciMacchinaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GestisciMacchinaToolStripMenuItem.Click
        Dim codiceSelezionato As String = DataGridView4.CurrentRow.Cells("Codice_").Value.ToString()

        ' Crea una nuova istanza del form
        Dim nuovaFinestra As New Form_distinta_vendita()

        ' Inizializza e mostra
        nuovaFinestra.inizializza_form(codiceSelezionato)
        nuovaFinestra.Show()
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub GestisciCodiceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GestisciCodiceToolStripMenuItem.Click
        Dim nuovaFinestra As New Form_Codici_vendita()
        nuovaFinestra.TopMost = True
        nuovaFinestra.Show()
        Dim codiceSelezionato As String = DataGridView2.CurrentRow.Cells("Codice___").Value.ToString()
        nuovaFinestra.inizializza_form(codiceSelezionato)

    End Sub

    Private Sub Button9_Click_2(sender As Object, e As EventArgs) Handles Button9.Click
        listino_ = True
        Dim par_datagridview As DataGridView = DataGridView4
        ' Ciclo per ogni riga della DataGridView
        For Each row As DataGridViewRow In par_datagridview.Rows
            If Not row.IsNewRow Then
                ' Ottieni l'indice della riga e della colonna "codice_"
                Dim rowIndex As Integer = row.Index
                Dim colIndex As Integer = par_datagridview.Columns("codice_").Index

                ' Imposta la cella corrente (opzionale ma utile)
                par_datagridview.CurrentCell = par_datagridview.Rows(rowIndex).Cells(colIndex)

                ' Richiama manualmente l'evento CellClick
                Call DataGridView4_CellClick(par_datagridview, New DataGridViewCellEventArgs(colIndex, rowIndex))
            End If
        Next
    End Sub



    Private Sub Button10_Click_2(sender As Object, e As EventArgs) Handles Button10.Click
        Dim par_datagridview As DataGridView = DataGridView2
        ' Ciclo per ogni riga della DataGridView
        For Each row As DataGridViewRow In par_datagridview.Rows
            If Not row.IsNewRow Then
                ' Ottieni l'indice della riga e della colonna "codice_"
                Dim rowIndex As Integer = row.Index
                Dim colIndex As Integer = par_datagridview.Columns("codice___").Index

                ' Imposta la cella corrente (opzionale ma utile)
                par_datagridview.CurrentCell = par_datagridview.Rows(rowIndex).Cells(colIndex)

                ' Richiama manualmente l'evento CellClick
                Call DataGridView2_CellClick(par_datagridview, New DataGridViewCellEventArgs(colIndex, rowIndex))
            End If
        Next
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs)
        Form_configuratore_vendita.Show()
    End Sub

    Private Sub DataGridView_commesse_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_commesse.CellEndEdit

    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView4_CellContextMenuStripChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContextMenuStripChanged

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        filtra_macchine()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        filtra_macchine()
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        filtra_macchine()
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged
        filtra_macchine()
    End Sub
End Class