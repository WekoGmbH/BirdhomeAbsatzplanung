Imports System.Collections.Generic

Imports Sagede.OfficeLine.Engine
Imports Sagede.Core.Tools
Imports Sagede.Core.Tools.DateTimeHelper
Imports Sagede.OfficeLine.Data
Imports System.IO
Imports System.Collections
Imports Sagede.OfficeLine.Wawi.BelegEngine
Imports Sagede.OfficeLine.Wawi.Services
Imports Sagede.OfficeLine.Wawi.Tools
Imports Sagede.OfficeLine.Wawi.BelegProxyEngine

Public Class clsImport
    Private oPlanung As clsPlanung

    Private nx As Int32
    Private J As Int32
    Private PlanJahr As List(Of clsPlanung)
    Private mandant As Mandant




    Public Sub New(man As Mandant)
        mandant = man
    End Sub


    Public Sub mReadPlanung(sFile As String, sPlannummer As Int32)
        Dim oEx As Object
        Dim oWS As Object
        Dim xlBook As Object
        Try

            oEx = CreateObject("Excel.Application")
            '   xlBook = oEx.Workbooks.Add
            '  oWS = xlBook.Worksheet

            oEx.Workbooks.Open(sFile)
            oWS = oEx.Worksheets.Item("Prognose")
            PlanJahr = New List(Of clsPlanung)
            nx = 2
            J = 1
            'For nx = 2 To oWS.Rows.Count
            While ConversionHelper.ToString(oWS.Cells(nx, J).Value) <> ""
                '    'Kennzahlen
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 13).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 13).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 14).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 14).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 15).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 15).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 16).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 16).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 17).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 17).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 18).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 18).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 19).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 19).value)
                End With
                PlanJahr.Add(oPlanung)

                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 20).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 20).value)
                End With
                PlanJahr.Add(oPlanung)

                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 21).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 21).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 22).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 22).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 23).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 23).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 24).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 24).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 25).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 25).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 26).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 26).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 27).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 27).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 28).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 28).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 29).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 29).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 30).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 30).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 31).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 31).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 32).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 32).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 33).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 33).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 34).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 34).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 35).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 35).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 36).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 36).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 37).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 37).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 38).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 38).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 39).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 39).value)
                End With
                PlanJahr.Add(oPlanung)
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 8).value)
                    .Plannummer = ConversionHelper.ToInt32(sPlannummer)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 40).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(1, 40).value)
                End With
                PlanJahr.Add(oPlanung)

                nx = nx + 1
                J = 1

            End While
            oWS = Nothing
            oEx.Quit()

            ' oEx.Application.Quit()
            oEx = Nothing
            TraceLog.LogVerbose("Absatzplanung gelesen")
            mCreatePlanung()
            TraceLog.LogVerbose("Absatzplanung in Tabellen geschrieben")
        Catch ex As Exception
            Throw ex
            'TraceLog.LogException(ex)
            'MsgBox(ex.Message)
        End Try





    End Sub
    Public Sub mReadLieferschein(sFile As String, sPlannummer As Int32)

        Dim oEx As Object
        Dim oWS As Object
        Dim xlBook As Object

        Try
            oEx = CreateObject("Excel.Application")
            'xlBook = oEx.Workbooks.Add
            'oWS = xlBook.Worksheet(1)

            '   oEx.Workbooks.Open(sFile)
            '  oWS = oEx.Worksheets.Item("Prognose")

            TraceLog.LogVerbose("WEKO Excel auslesen")
            oEx.Workbooks.Open(sFile)
            'oWS = CType(oEx.Worksheets.Item("Worksheet 1"), Worksheet)
            oWS = oEx.Worksheets.Item("Worksheet 1")
            PlanJahr = New List(Of clsPlanung)
            nx = 2
            J = 1
            'For nx = 2 To oWS.Rows.Count
            While ConversionHelper.ToString(oWS.Cells(nx, J).Value) <> ""
                '    'Kennzahlen
                oPlanung = New clsPlanung
                With oPlanung
                    .Artikelnummer = ConversionHelper.ToString(oWS.Cells(nx, 2).value)
                    .Menge = ConversionHelper.ToDecimal(oWS.Cells(nx, 3).value)
                    .Datum = ConversionHelper.ToDateTime(oWS.Cells(nx, 1).value)
                    '.Datum = ConversionHelper.ToDateTime(Mid(oWS.Cells(nx, 1).value, 1, 5) & Mid(oWS.Cells(nx, 1).value, 9, 2) & Mid(oWS.Cells(nx, 1).value, 5, 3))
                End With
                PlanJahr.Add(oPlanung)
                'oPlanung = New clsPlanung
                'PlanJahr.Add(oPlanung)

                nx = nx + 1
                J = 1

            End While
            oWS = Nothing
            oEx.Quit()

            ' oEx.Application.Quit()
            oEx = Nothing
            TraceLog.LogVerbose("Lieferschein gelesen")
            mCreateLieferschein()
            TraceLog.LogVerbose("Lieferschein angelegt")
        Catch ex As Exception
            Throw ex
            TraceLog.LogException(ex)
            'MsgBox(ex.Message)
        End Try





    End Sub
    Public Sub mCreateLieferschein()

        Try
            ' Dim objReader As New StreamReader(sFile)
            'Dim sLine As String = ""
            'Dim arrText As New String
            'Dim arrText As String()
            Dim sQry As String
            Dim oBeleg As Beleg
            Dim oPos As BelegPosition
            Dim Positionen As IGenericReader
            Dim Lieferdatum As Date

            '  Lieferdatum = DateTime.Today

            TraceLog.LogVerbose("Lieferschein tmp  tWEKOLieferscheinImport schreiben")

            sQry = "delete from  tWEKOLieferscheinImport where Benutzer =  " & "'" & mandant.Benutzer.Name & "'" & " AND Mandant=" & mandant.Id
            mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            For Each oPlanung In PlanJahr
                '           If oPlanung.Artikelnummer = "" Or IsDBNull(oPlanung.Artikelnummer) Then
                '           Else
                sQry = "INSERT INTO tWEKOLieferscheinImport (EAN, Menge, Benutzer, Mandant, Datum) VALUES (" & "'" & oPlanung.Artikelnummer & "'" & "," & Str((oPlanung.Menge)) & "," & "'" & mandant.Benutzer.Name & "'" & "," & mandant.Id & "," & "'" & (oPlanung.Datum) & "'" & ")"
                mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)
                '          End If
            Next
            TraceLog.LogVerbose("Lieferschein tmp in  tWEKOLieferscheinImport geschrieben")




            sQry = "delete from  tWEKOLieferscheinImport where Menge = 0 AND Benutzer =  " & "'" & mandant.Benutzer.Name & "'" & " AND Mandant=" & mandant.Id
            mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)


            sQry = "UPDATE       tWEKOLieferscheinImport SET  Lieferdatum = (CONVERT(DATETIME, SUBSTRING(CAST(Datum AS varchar(20)), 7, 4) + '-' + SUBSTRING(RTRIM(CAST(Datum AS varchar(20))), 1, 2) + '-' + SUBSTRING(RTRIM(CAST(Datum AS varchar(20))), 4, 2) , 104))  WHERE        (Mandant = " & mandant.Id & ") AND (Benutzer = " & "'" & mandant.Benutzer.Name & "'" & ")"
            mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            sQry = "UPDATE tWEKOLieferscheinImport SET        tWEKOLieferscheinImport.Artikelnummer= KHKArtikelVarianten.Artikelnummer From tWEKOLieferscheinImport INNER Join " &
                   "      KHKArtikelVarianten On tWEKOLieferscheinImport.Mandant = KHKArtikelVarianten.Mandant And tWEKOLieferscheinImport.EAN = KHKArtikelVarianten.EANNummer " &
                    " Where (tWEKOLieferscheinImport.Benutzer = " & "'" & mandant.Benutzer.Name & "'" & ") AND (tWEKOLieferscheinImport.Mandant = " & mandant.Id & ")"
            mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            sQry = "UPDATE tWEKOLieferscheinImport SET  tWEKOLieferscheinImport.BelID = KHKVKBelege.BelID FROM  KHKVKBelege INNER JOIN " &
                   "      KHKVKBelegePositionen ON KHKVKBelege.Mandant = KHKVKBelegePositionen.Mandant AND KHKVKBelege.BelID = KHKVKBelegePositionen.BelID INNER JOIN " &
                  "       tWEKOLieferscheinImport ON KHKVKBelegePositionen.Artikelnummer = tWEKOLieferscheinImport.Artikelnummer AND KHKVKBelegePositionen.Liefertermin = tWEKOLieferscheinImport.Lieferdatum " &
                    " WHERE        (KHKVKBelege.Belegkennzeichen = 'VLL') AND (tWEKOLieferscheinImport.Benutzer = " & "'" & mandant.Benutzer.Name & "'" & ") AND (tWEKOLieferscheinImport.Mandant = " & mandant.Id & ")"
            mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            sQry = "UPDATE  tWEKOLieferscheinImport SET  BelID = 99999 " &
                    " Where  (Lieferdatum < CONVERT(DATETIME, '2019-03-27 00:00:00', 102)) AND (tWEKOLieferscheinImport.Benutzer = " & "'" & mandant.Benutzer.Name & "'" & ") AND (tWEKOLieferscheinImport.Mandant = " & mandant.Id & ")"
            mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            TraceLog.LogVerbose("Lieferschein Daten in  tWEKOLieferscheinImport aktualisiert")

            'If mandant.MainDevice.Lookup.RowExists("Belegnummer", "KHKVKBelege", "Belegdatum = " & "'" & Lieferdatum & "'" & " AND Belegkennzeichen = 'VLL' AND A0Empfaenger='" & "100030" & "' and Mandant=" & mandant.Id & "") = True Then
            '        MsgBox("Es gibt bereits einen Lieferschein, der Vorgang wird abgebrochen!", MsgBoxStyle.Critical)

            '        GoTo Ende
            '    End If


            sQry = "SELECT Artikelnummer, Menge, Benutzer, Mandant FROM  tWEKOLieferscheinImport WHERE   (BelID IS NULL) AND     (Benutzer = " & "'" & mandant.Benutzer.Name & "'" & ") AND (Mandant = " & mandant.Id & ") AND (NOT (Artikelnummer IS NULL)) ORDER BY Artikelnummer"
            Positionen = mandant.MainDevice.GenericConnection.ExecuteReader(sQry)
            If Positionen.Read = False Then
                Throw New Exception("Es gibt bereits einen Lieferschein, die Positionen wurden eingelesen, der Vorgang wird abgebrochen!")
                'MsgBox("Es gibt bereits einen Lieferschein, die Positionen wurden eingelesen, der Vorgang wird abgebrochen!", MsgBoxStyle.Critical)

                GoTo Ende
            End If


            oBeleg = New Beleg(mandant, Sagede.OfficeLine.Wawi.Tools.Erfassungsart.Verkauf)
            With oBeleg
                .Initialize("VLL", DateTime.Today, mandant.PeriodenManager.Perioden.Date2Periode(DateTime.Today).Jahr)
                .Bearbeiter = mandant.Benutzer.Name
                .SetKonto("100020", False)
                '.SetKonto("D100000", False)
                .Belegdatum = DateTime.Today
                TraceLog.LogVerbose("Beleg wurde initialisiert.")
                'sQry = "SELECT Artikelnummer, Format(Lieferdatum,'d','de-de') as Lieferdatum, Menge, Benutzer, Mandant FROM  tWEKOLieferscheinImport WHERE   (BelID IS NULL) AND     (Benutzer = " & "'" & mandant.Benutzer.Name & "'" & ") AND (Mandant = " & mandant.Id & ") AND (NOT (Artikelnummer IS NULL)) ORDER BY Artikelnummer"
                'sQry = "SELECT Artikelnummer, replace(convert(NVARCHAR, Lieferdatum, 104), ' ', '/') as Lieferdatum, Menge, Benutzer, Mandant FROM  tWEKOLieferscheinImport WHERE   (BelID IS NULL) AND     (Benutzer = " & "'" & mandant.Benutzer.Name & "'" & ") AND (Mandant = " & mandant.Id & ") AND (NOT (Artikelnummer IS NULL)) ORDER BY Artikelnummer"
                sQry = "SELECT Artikelnummer, Convert(DateTime, SUBSTRING(CAST(Datum As varchar(20)), 7, 4) + '-' + SUBSTRING(RTRIM(CAST(Datum AS varchar(20))), 1, 2) + '-' + SUBSTRING(RTRIM(CAST(Datum AS varchar(20))), 4, 2), 104) as Lieferdatum, Menge, Benutzer, Mandant FROM  tWEKOLieferscheinImport WHERE   (BelID IS NULL) AND     (Benutzer = " & "'" & mandant.Benutzer.Name & "'" & ") AND (Mandant = " & mandant.Id & ") AND (NOT (Artikelnummer IS NULL)) ORDER BY Artikelnummer"
                Positionen = mandant.MainDevice.GenericConnection.ExecuteReader(sQry)
                Do While Positionen.Read

                    oPos = New BelegPosition(oBeleg)
                    oPos.SetArtikel(Positionen.GetString("Artikelnummer"), 0)
                    oPos.Liefertermin = (Positionen.GetValue("Lieferdatum"))
                    ' oPos.Liefertermin = DateTime.Today

                    oPos.Menge = Positionen.GetDecimal("Menge")
                    oPos.Calculate()
                    oBeleg.Positionen.Add(oPos)

                Loop
                Positionen.Close()
                .Renumber()

                .Calculate(True)

                If .Save(True) = False Then
                    For X = 0 To oBeleg.Errors.Count - 1
                        Throw New Exception("Fehler beim Beleg erstellen: " & oBeleg.Errors.Item(X).Description)
                        ' MsgBox(oBeleg.Errors.Item(X).Description)
                    Next

                Else

                End If
            End With


            TraceLog.LogVerbose("Lieferschein als Beleg angelegt")
Ende:

            sQry = "SELECT * FROM  tWEKOLieferscheinImport WHERE   (BelID IS NULL) AND     (Benutzer = " & "'" & mandant.Benutzer.Name & "'" & ") AND (Mandant = " & mandant.Id & ") AND (NOT (Artikelnummer IS NULL)) ORDER BY Artikelnummer"
            Positionen = mandant.MainDevice.GenericConnection.ExecuteReader(sQry)
            Do While Positionen.Read

                sQry = "INSERT INTO WEKO_LieferscheinImportLog ( Artikelnummer, Menge, EAN, Benutzer, Lieferdatum, Mandant, BelID, ImportBenutzer, ImportDatum) " &
                        " SELECT      " & "'" & Positionen.GetString("Artikelnummer") & "'" & ", " & Str(Positionen.GetDecimal("Menge")) & "," & "'" & Positionen.GetString("EAN") & "'" & ", " & "'" & mandant.Benutzer.Name & "'" & ", " & "'" & DateValue(Positionen.GetValue("Lieferdatum")) & "'" & ", " & mandant.Id & ", " & Positionen.GetInt32("BelID") & ", " & "'" & mandant.Benutzer.Name & "'" & "," & "'" & DateValue(Now) & "'"
                mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            Loop
            Positionen.Close()

            TraceLog.LogVerbose("Lieferschein beendet")
        Catch ex As Exception
            Throw ex
            'TraceLog.LogException(ex)
            'MsgBox(ex.Message)
        End Try





    End Sub


    Private Sub mCreatePlanung()

        Dim sQry As String
        Dim PlanNr As Int32
        Dim PlanPosID As Int32



        Try


            sQry = "delete from  tWEKOArtikelPlanung where Benutzer =  " & "'" & mandant.Benutzer.Name & "'" & " AND Mandant=" & mandant.Id
            mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            PlanNr = oPlanung.Plannummer

            For Each oPlanung In PlanJahr
                PlanPosID = mandant.GetTan("KHKArtikelPlanungPositionen")
                sQry = "INSERT INTO tWEKOArtikelPlanung (PlanId, Artikelnummer,  Datum, Menge, Mandant, Benutzer,PlanPosID) VALUES (" & oPlanung.Plannummer & ", " & "'" & oPlanung.Artikelnummer & "'" & "," & Sagede.Core.Tools.DateTimeHelper.DateToSqlServer(oPlanung.Datum) & "," & Str(oPlanung.Menge) & "," & mandant.Id & "," & "'" & mandant.Benutzer.Name & "'" & "," & PlanPosID & ")"
                mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)
            Next

            sQry = "DELETE FROM tWEKOArtikelPlanung WHERE  (Artikelnummer = '') AND Benutzer =  " & "'" & mandant.Benutzer.Name & "'" & " AND Mandant=" & mandant.Id
            mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            'sQry = "UPDATE tWEKOArtikelPlanung SET  tWEKOArtikelPlanung.PlanID = KHKArtikelPlanung.PlanID From KHKArtikelPlanung INNER Join " &
            '        " tWEKOArtikelPlanung On KHKArtikelPlanung.Mandant = tWEKOArtikelPlanung.Mandant And KHKArtikelPlanung.Plannummer = tWEKOArtikelPlanung.Plannummer " &
            '        " Where (tWEKOArtikelPlanung.Benutzer = " & "'" & mandant.Benutzer.Name & "'" & ") AND (tWEKOArtikelPlanung.Mandant = " & mandant.Id & ")"
            'mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            sQry = "INSERT INTO WEKO_KHKArtikelPlanungPositionenLog (PlanPosID, PlanID, Mandant, Artikelnummer, AuspraegungID, Datum, MengeBasis0, MengeBasis1, MengeBasisAuswertung, DatumErstellt, BenutzerErstellt, DatumGeaendert, BenutzerGeaendert, BenutzerPlan, DatumPlan) " &
                    " SELECT KHKArtikelPlanungPositionen.PlanPosID, KHKArtikelPlanungPositionen.PlanID, KHKArtikelPlanungPositionen.Mandant, KHKArtikelPlanungPositionen.Artikelnummer, KHKArtikelPlanungPositionen.AuspraegungID, " &
                    "     KHKArtikelPlanungPositionen.Datum, KHKArtikelPlanungPositionen.MengeBasis0, KHKArtikelPlanungPositionen.MengeBasis1, KHKArtikelPlanungPositionen.MengeBasisAuswertung, " &
                     "    KHKArtikelPlanungPositionen.DatumErstellt, KHKArtikelPlanungPositionen.BenutzerErstellt, KHKArtikelPlanungPositionen.DatumGeaendert, KHKArtikelPlanungPositionen.BenutzerGeaendert, " & "'" & mandant.Benutzer.Name & "'" & " AS Expr1, " &
                    "      GETDATE() AS Expr2 FROM            KHKArtikelPlanungPositionen INNER JOIN " &
                    "     KHKArtikelPlanung ON KHKArtikelPlanungPositionen.PlanID = KHKArtikelPlanung.PlanID AND KHKArtikelPlanungPositionen.Mandant = KHKArtikelPlanung.Mandant " &
                    " WHERE        (KHKArtikelPlanungPositionen.Mandant = " & mandant.Id & ") AND (KHKArtikelPlanung.PlanId = " & PlanNr & ")"
            mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            sQry = "UPDATE  KHKArtikelPlanungPositionen SET       KHKArtikelPlanungPositionen.BenutzerGeaendert = 'Leeren' FROM            KHKArtikelPlanungPositionen INNER JOIN " &
                   "      tWEKOArtikelPlanung ON KHKArtikelPlanungPositionen.Mandant = tWEKOArtikelPlanung.Mandant AND KHKArtikelPlanungPositionen.Artikelnummer = tWEKOArtikelPlanung.Artikelnummer " &
                    "WHERE        (KHKArtikelPlanungPositionen.Mandant = " & mandant.Id & ") AND (KHKArtikelPlanungPositionen.PlanID = " & PlanNr & ")"
            mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            sQry = "DELETE  FROM   KHKArtikelPlanungPositionen WHERE   BenutzerGeaendert = 'Leeren' AND  (PlanID = " & PlanNr & ") AND (Mandant = " & mandant.Id & ")"
            mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)


            sQry = "INSERT INTO KHKArtikelPlanungPositionen (PlanID, Mandant, Artikelnummer, Datum, MengeBasis0, MengeBasis1, MengeBasisAuswertung, DatumErstellt, BenutzerErstellt, DatumGeaendert, BenutzerGeaendert, AuspraegungID, PlanPosID) " &
                "SELECT  tWEKOArtikelPlanung.PlanID, tWEKOArtikelPlanung.Mandant, tWEKOArtikelPlanung.Artikelnummer, tWEKOArtikelPlanung.Datum, tWEKOArtikelPlanung.Menge AS MengeBasis0, 0 AS MengeBasis1, " &
                "         tWEKOArtikelPlanung.Menge AS MengeBasisAuswertung, GETDATE() AS DatumErstellt, tWEKOArtikelPlanung.Benutzer AS BenutzerErstellt, GETDATE() AS DatumGeaendert, " &
                "         tWEKOArtikelPlanung.Benutzer AS BenutzerGeaendert, KHKArtikelVarianten.AuspraegungID, tWEKOArtikelPlanung.PlanPosID " &
                " FROM            tWEKOArtikelPlanung INNER JOIN KHKArtikelVarianten ON tWEKOArtikelPlanung.Artikelnummer = KHKArtikelVarianten.Artikelnummer AND tWEKOArtikelPlanung.Mandant = KHKArtikelVarianten.Mandant " &
                " WHERE        (tWEKOArtikelPlanung.Mandant = " & mandant.Id & ") AND (tWEKOArtikelPlanung.Benutzer = " & "'" & mandant.Benutzer.Name & "'" & ") AND (tWEKOArtikelPlanung.PlanID = " & PlanNr & ")"
            mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)
        Catch ex As Exception

        End Try
    End Sub

End Class
