Imports Sagede.Core.ServerEnvironment
Imports Sagede.Settings.OfficeLine
Imports Sagede.[Shared].BlobStorage
Imports System
Imports System.IO
Imports System.Security


Public Class BlobProvider
    Private ReadOnly _user As String
    Private ReadOnly _password As SecureString
    Private _container As IContainer

    Public Sub New(ByVal user As String, ByVal password As SecureString)
        _user = user
        _password = password
    End Sub

    Private Function GetContainer() As IContainer
        Try
            If _container IsNot Nothing Then Return _container
            Dim configurationAccess = TryCast(ConfigurationService.ConfigurationAccess, ConfigurationAccess)

            If configurationAccess IsNot Nothing Then
                Dim serviceRegistryStorage = New ServiceRegistryStorageRemoteRegistry(configurationAccess.RemoteRegistryServerName, configurationAccess.RemoteRegistryDefaultKey)
                Dim storage = StorageProvider.GetBlobStorage(serviceRegistryStorage, _user, _password)

                If storage IsNot Nothing Then
                    _container = storage.GetContainerReference("Data")

                    If _container IsNot Nothing Then
                        _container.CreateIfNotExists()
                    Else
                        TraceLog.Logger.Warning("GetContainer() _container==null")
                    End If
                Else
                    TraceLog.Logger.Warning("GetContainer() storage==null")
                End If
            Else
                TraceLog.Logger.Warning("GetContainer() configurationAccess==null")
            End If

        Catch ex As Exception
            TraceLog.LogException(ex)
            Return Nothing
        End Try

        Return _container
    End Function

    Public Sub SaveBlob(ByVal blobName As String, ByVal path As String, ByVal content As Byte())
        Try
            If Not String.IsNullOrEmpty(path) Then blobName = String.Join("/", path, blobName)
            Dim container = GetContainer()
            If container Is Nothing Then Throw New Exception("ErrorBlobStorageUnavailable")
            container.GetBlobReference(blobName).DeleteIfExists()

            Using stream = New MemoryStream(content)
                container.GetBlobReference(blobName).UploadFromStream(stream)
                stream.Close()
                stream.Dispose()
            End Using

            Dim blob = container.GetBlobReference(blobName)
            Dim meta = blob.GetMetadata()
            meta.[Public] = True
            blob.SetMetadata(meta)
        Catch ex As Exception
            TraceLog.LogException(ex)
        End Try
    End Sub

    Public Function GetBlob(ByVal blobName As String, ByVal path As String) As Byte()
        Try
            If Not BlobExists(blobName, path) Then Return Nothing
            If Not String.IsNullOrEmpty(path) Then blobName = String.Join("/", path, blobName)
            Return GetBlob(blobName)
        Catch ex As Exception
            TraceLog.LogException(ex)
            Return Nothing
        End Try
    End Function

    Public Function GetBlob(ByVal fullBlobName As String) As Byte()
        Try
            Dim container = GetContainer()
            If container Is Nothing Then Throw New Exception("BlobStorage not available")

            Using stream = New MemoryStream()
                container.GetBlobReference(fullBlobName).DownloadToStream(stream)
                Dim content = stream.ToArray()
                stream.Close()
                stream.Dispose()
                Return content
            End Using

        Catch ex As Exception
            TraceLog.LogException(ex)
            Return Nothing
        End Try
    End Function

    Public Function GetBlobAsFile(ByVal fullBlobName As String, fileNameUebergabe As String) As String
        Try
            Dim container = GetContainer()
            If container Is Nothing Then Throw New Exception("ErrorBlobStorageUnavailable")
            Dim fileName = fileNameUebergabe
            'If File.Exists(fileName) Then File.Delete(fileName)

            Using stream = New FileStream(fileName, FileMode.Create)
                container.GetBlobReference(fullBlobName).DownloadToStream(stream)
                stream.Close()
                stream.Dispose()
            End Using

            Return fileName
        Catch ex As Exception
            TraceLog.LogException(ex)
            Return Nothing
        End Try
    End Function

    Public Function BlobExists(ByVal blobName As String, ByVal path As String) As Boolean
        Try
            If Not String.IsNullOrEmpty(path) Then blobName = String.Join("/", path, blobName)
            Dim container = GetContainer()
            If container Is Nothing Then Throw New Exception("ErrorBlobStorageUnavailable")
            Return container.GetBlobReference(blobName).Exists()
        Catch ex As Exception
            TraceLog.LogException(ex)
            Return False
        End Try
    End Function

    Public Function DeleteBlob(ByVal blobName As String, ByVal path As String) As Boolean
        Try
            If Not String.IsNullOrEmpty(path) Then blobName = String.Join("/", path, blobName)
            Dim container = GetContainer()
            If container Is Nothing Then Throw New Exception("ErrorBlobStorageUnavailable")
            Return container.GetBlobReference(blobName).DeleteIfExists()
        Catch ex As Exception
            TraceLog.LogException(ex)
            Return False
        End Try
    End Function
End Class

