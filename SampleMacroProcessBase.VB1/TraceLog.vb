Imports Sagede.Core.Logging

''' <summary>
''' Stellt eine, für diese Assembly,
''' exclusive Instanz des Loggers zur Verfügung.
''' </summary>
Public NotInheritable Class TraceLog

    Private Sub New()
    End Sub

    Private Shared _logger As ILogger
    Private Shared ReadOnly LockObject As New Object()

    ''' <summary>
    ''' Liefert die Instanz des Loggers.
    ''' </summary>
    Friend Shared ReadOnly Property Logger() As ILogger
        Get
            SyncLock LockObject
                Return If(_logger, (InlineAssignHelper(_logger, LogManager.GetLogger("Sage100", "WEKO ImportAbsatzplanung"))))
            End SyncLock
        End Get
    End Property

    ''' <summary>
    ''' Loggt Debug-Informationen im TraceLog-Manager
    ''' </summary>
    ''' <param name="message"></param>
    ''' <param name="args"></param>
    Public Shared Sub LogVerbose(message As String, ParamArray args As Object())
        'Logger.LogVerboseFormat(message, args);
        Logger.Verbose(String.Format(message, args))
    End Sub

    ''' <summary>
    ''' Loggt Exceptions im Tracelog-Manager
    ''' </summary>
    ''' <param name="ex"></param>
    Public Shared Sub LogException(ex As Exception)
        Logger.[Error]([String].Format("{0},{1},{2}", ex.Message, Environment.NewLine, ex.StackTrace))
    End Sub

    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function

End Class