Imports Npgsql

Public Class Modulo_dip
    Sub inizializza_modulo(par_risorsa)
        dati_risorsa(par_risorsa)
    End Sub

    Sub dati_risorsa(par_risorsa As Integer)

        Dim connString As String = Homepage.JPM_TIRELLI
        Using conn As New NpgsqlConnection(connString)
            conn.Open()
            ' Esegui le query qui
            Dim STRINGA_QUERY As String = "SELECT 
  res.uid AS resuid, 
  res.rescod AS codice_risorsa,
  res.resdsc AS descrizione_risorsa,
  res.dteval,
  grp.grpcod AS codice_gruppo, 
  gl.grpdsc AS descr_gruppo, 
substring(gl.grpdsc,1,3) as Prime_tre_rep,
  lvl.lvlcod AS codice_org, 
  lvl.lvldsc AS descr_org
  


FROM angres res
LEFT JOIN angresgrp rg ON (res.uid = rg.resuid AND rg.prjgrppri = -1)
LEFT JOIN anggrp grp ON (grp.uid = rg.grpuid)
LEFT JOIN anggrplng gl ON (grp.uid = gl.recuid AND gl.lnguid = 1)
LEFT JOIN orglvlres ol ON (ol.resuid = res.uid)
LEFT JOIN orglvl lvl ON (lvl.uid = ol.lvluid)
WHERE 
  res.logdel = 0 and res.rescod = '" & par_risorsa & "'  ;"

            Dim cmd As New NpgsqlCommand(STRINGA_QUERY, conn)
            Dim reader As NpgsqlDataReader = cmd.ExecuteReader()


            If reader.Read() Then
                Dim nomeCompleto As String = reader("descrizione_risorsa").ToString().Trim()
                Dim iniziali As String = String.Join("", nomeCompleto.Split(" "c).Where(Function(p) p.Length > 0).Select(Function(p) p(0).ToString().ToUpper()))
                Label1.Text = iniziali
                Label2.Text = reader("prime_tre_rep").ToString()

            End If




            reader.Close()
            conn.Close()
        End Using
    End Sub
End Class
