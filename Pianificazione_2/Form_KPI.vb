Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib
Imports System.Security.Cryptography
Imports Microsoft.Office.Interop.Word
Imports System.Windows.Media.Media3D
Imports System.Security.Policy
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView
Imports System.Diagnostics
Imports System.Net.Http
Imports Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.Runtime.InteropServices

Public Class Form_KPI
    Private stopCycle As Boolean = False


    Sub avvia_KPI(par_nome_file As String, par_tempo_slide As Integer)
        ' Creare una nuova istanza di Excel
        Dim excelApp As New Excel.Application
        ' Disabilitare la visibilità per permettere di modificarlo
        excelApp.Visible = True

        ' Aprire il file Excel dal Desktop
        Dim filePath As String = "\\Tirfs01\tirelli\00-Tirelli 4.0\KPI TV\" & par_nome_file & ""
        Dim workbook As Excel.Workbook = excelApp.Workbooks.Open(filePath)

        ' Disabilitare le barre degli strumenti per ottenere un'esperienza a schermo intero
        excelApp.DisplayFullScreen = True
        ' excelApp.WindowState = Excel.XlWindowState.xlMaximized

        ' Inizializzare la variabile di controllo
        stopCycle = False

        ' Ciclo che continua finché ci sono schede visibili
        Dim i As Integer = 1
        Do
            ' Controllo se l'utente ha premuto il tasto Esc
            If GetAsyncKeyState(Keys.Escape) <> 0 Then
                stopCycle = True
            End If

            ' Verifica se è stato richiesto di fermare il ciclo
            If stopCycle Then Exit Do

            ' Verifica se il foglio è visibile
            If workbook.Sheets(i).Visible = Excel.XlSheetVisibility.xlSheetVisible Then
                ' Attivare il foglio
                workbook.Sheets(i).Activate()

                ' Attendere 2 secondi prima di passare al foglio successivo
                System.Threading.Thread.Sleep(par_tempo_slide)
            End If

            ' Incrementa il contatore
            i += 1

            ' Se si è arrivati all'ultimo foglio, ripartire dal primo
            If i > workbook.Sheets.Count Then
                i = 1
            End If

            ' Permette l'elaborazione degli eventi di Excel
            System.Windows.Forms.Application.DoEvents()

        Loop While True

        ' Chiudere il file e l'applicazione Excel se necessario
        workbook.Close(SaveChanges:=False)
        excelApp.Quit()
    End Sub

    ' Importare la funzione GetAsyncKeyState per rilevare la pressione del tasto Esc
    <DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Percorso della cartella da controllare
        Dim folderPath As String = "\\Tirfs01\tirelli\00-Tirelli 4.0\KPI TV"
        ' Nome del file da cercare
        Dim fileName As String = "Chiudi.txt"
        ' Controlla se il file esiste
        If File.Exists(Path.Combine(folderPath, fileName)) Then
            ' Chiudi il programma
            ' MessageBox.Show("File 'chiudi.txt' trovato. Il programma verrà chiuso.", "Chiusura", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Environment.Exit(0)
        End If
    End Sub

    Private Sub Form_KPI_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class