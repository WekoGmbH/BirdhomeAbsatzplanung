Imports Sagede.OfficeLine.Shared.RealTimeData.MacroProcess
' required for Extension of MacroProcessBase
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports Sagede.Core.Data
' Referenz benötigt für Handling Datacontainer / Datacontainerset
Imports Sagede.Shared.RealTimeData.Common
Imports Sagede.OfficeLine.Engine
Imports Sagede.OfficeLine.Data


Public Class MacroProcess
    Inherits MacroProcessBase
    ''' <summary>
    ''' Name des Macroprozesses
    ''' </summary>
    Protected Overrides ReadOnly Property Name() As String
        'SAMPLE: get { return "Mein Macroprozess"; }
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    ''' <summary>
    ''' Execute Methode wird aus dem Makro heraus über "AufrufenDLL" angesprochen
    ''' MakroParameter 1: Vollständiger DLL Name inklusive .dll, z.B. PSDEV.OfficeLine.Examples.RealTimeData.dll
    ''' MakroParameter 2: Vollständiger Name der anzusprechenden Klasse, z.B. PSDEV.OfficeLine.Examples.RealTimeData.SampleMacroProcess
    ''' MakroParameter 3: Zu übergebende Strukturfelder in Makro Syntax inkl. [ ], z.B. [Kto];[VorID] - mit Semikolon getrennt
    ''' </summary>
    ''' <param name="parameters">Collection von Datenstrukturfeldern die in die Macroprozessimplementierung hineingereicht wird</param>
    ''' <param name="cancel">Rückgabe ob Macrofunktionalität abgebrochen wurde true / false</param>
    ''' <param name="cancelMessage">Fehlermeldung falls Macrofunktionalität abgebrochen wurde</param>
    ''' <returns></returns>
    '''
    Protected Overrides Function Execute(parameters As NamedParameters, ByRef cancel As Boolean, ByRef cancelMessage As String) As NamedParameters
        Try
            ' TODO: Logik implementieren
            Dim oPlanung As New clsImport(Mandant)

            Dim blobPath = parameters.Item(0)
            Dim importDatei = parameters.Item(1).Value
            Dim blobPathFile = parameters.Item(2).Value
            Dim Plannummer = parameters.Item(3).Value
            Dim Auswahl = parameters.Item(4).Value
            Dim blobProvider = New BlobProvider(Mandant.Credential.Name, Mandant.Credential.Password)
            Dim nDatei As String

            TraceLog.LogVerbose("WEKO Blob Pfad ")

            Dim pfad = blobProvider.GetBlobAsFile(blobPathFile, importDatei)
            TraceLog.LogVerbose("WEKO Blob Pfad " & pfad)
            nDatei = "C:\ProgramData\Sage\BlobStorage\data\Containers\Data\BirdHomeAutomation\8\Batchfiles\" & importDatei
            TraceLog.LogVerbose("WEKO Blob DAtei " & nDatei)
            'nDatei = "C:\ProgramData\Sage\BlobStorage\data\Containers\Data\BirdHomeAutomation\8\BirdHomeAutomation\8\BatchFiles\" & importDatei


            If Auswahl = "Artikel" Then
                TraceLog.LogVerbose("Absatzplanung start")
                ' pfad = "C:\ProgramData\Sage\BlobStorage\openports\data\Containers\Data\" & blobPathFile
                oPlanung.mReadPlanung(nDatei, Plannummer)
                'oPlanung.mReadPlanung(importDatei, Plannummer)
            End If
            If Auswahl = "Lieferschein" Then
                TraceLog.LogVerbose("WEKO Lieferschein start")
                '  pfad = "C:\ProgramData\Sage\BlobStorage\openports\data\Containers\Data\" & blobPathFile
                oPlanung.mReadLieferschein(nDatei, Plannummer)

            End If

            ' TODO: Rückgabewerte setzen
            cancel = False
            cancelMessage = "Import durchgeführt" + [String].Empty
            Return parameters
        Catch ex As Exception
            TraceLog.LogException(ex)
            cancel = True
            cancelMessage = "Fehler in Erweiterung: " + ex.Message
            'cancelMessage = ex.Message
            Return parameters
        End Try
    End Function

    ''' <summary>
    ''' Vorbereitung der Ausführung
    ''' </summary>
    Protected Overrides Sub Prepare()
        ' Throw New NotImplementedException()
    End Sub
End Class
