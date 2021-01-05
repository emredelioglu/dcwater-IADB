Imports Microsoft.VisualBasic
Imports ESRI.ArcGIS.JTX
Imports ESRI.ArcGIS.Geodatabase
Imports System.IO

Public Class mdlJTX

    Public Shared Function GetFeatureWorkspace(ByVal jtxDB As String, ByVal jtxDBWorkspace As String, ByVal jtxVersionName As String)
        Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
        Dim pJTXDatabase As IJTXDatabase2 = pJTXDBMan.GetDatabase(jtxDB)
        Dim pJTXDataWorkspaceNameSet As IJTXDataWorkspaceNameSet = pJTXDatabase.GetDataWorkspaceNames(Nothing)
        Dim pJTXDataWorkspaceName As IJTXDataWorkspaceName

        pJTXDataWorkspaceNameSet.Reset()
        pJTXDataWorkspaceName = pJTXDataWorkspaceNameSet.Next()

        Dim jtxDBID As String
        Do While Not pJTXDataWorkspaceName Is Nothing
            If pJTXDataWorkspaceName.Name = jtxDBWorkspace Then
                jtxDBID = pJTXDataWorkspaceName.DatabaseID
                Exit Do
            End If
            pJTXDataWorkspaceName = pJTXDataWorkspaceNameSet.Next()
        Loop

        If jtxDBID Is Nothing Then
            Console.WriteLine("Invalid Workspace Name : " & jtxDBWorkspace)
            Throw New Exception("Invalid Workspace Name : " & jtxDBWorkspace)
        End If

        Return pJTXDatabase.GetDataWorkspace(jtxDBID, jtxVersionName)
    End Function

    Public Shared Function CreateJob(ByVal pJTXDatabase As IJTXDatabase, ByVal jobTypeName As String, ByVal workspaceName As String, ByVal parentVersion As String, ByVal log As System.IO.StreamWriter) As IJTXJob

        Dim pJTXJobMan As IJTXJobManager = Nothing
        Dim pNewJob As IJTXJob2 = Nothing

        Try
            Dim pJTXDBMan As IJTXDatabaseManager
            pJTXDBMan = New JTXDatabaseManager
            pJTXJobMan = pJTXDatabase.JobManager

            Dim jconfigMan As JTXConfigurationManager
            jconfigMan = pJTXDatabase.ConfigurationManager

            Dim pJtxJobTypePremiseExceptions As IJTXJobType = jconfigMan.GetJobType(jobTypeName)
            pNewJob = pJTXJobMan.CreateJob(pJtxJobTypePremiseExceptions, 0, True)

            Dim jtxWorkflowUtil As IJTXWorkflowUtilities = New JTXUtility
            Dim strWorkflowXML As String = jtxWorkflowUtil.GetWorkflowXML(pJTXDatabase, pJtxJobTypePremiseExceptions)
            If strWorkflowXML <> "" Then
                jtxWorkflowUtil.CopyWorkflowXML(pJTXDatabase, pNewJob)
            End If

            Dim pActType As IJTXActivityType = jconfigMan.GetActivityType("CreateJob")
            If Not pActType Is Nothing Then
                pNewJob.LogJobAction(pActType, Nothing, "")
            End If

            pNewJob.ParentVersion = parentVersion
            pNewJob.Status = jconfigMan.GetStatus("Created")
            pNewJob.Store()

            If Not pNewJob.VersionExists Then
                pNewJob.CreateVersion(esriVersionAccess.esriVersionAccessProtected)
            End If

            Dim jtxSystemUtil As IJTXSystemUtilities = New JTXUtility
            jtxSystemUtil.SendNotification("JobCreated", pJTXDatabase, pNewJob, Nothing)

            Return pNewJob
        Catch ex As Exception
            If Not pNewJob Is Nothing Then
                pJTXJobMan.DeleteJob(pNewJob.ID, True)

                If pNewJob.VersionExists = True Then
                    pNewJob.DeleteVersion()
                End If

            Else
                log.WriteLine(Now & " Job creation failed.")
            End If

            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

            Throw ex
        End Try

    End Function

    Public Shared Sub CreateJobVersion(ByVal pNewJob As IJTXJob, ByVal log As System.IO.StreamWriter)
        Try
            pNewJob.CreateVersion(esriVersionAccess.esriVersionAccessPublic)
            log.WriteLine(Now & ": version " & pNewJob.VersionName & " has been createcd for job " & pNewJob.ID)
        Catch ex As Exception
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

            Throw ex
        End Try
    End Sub


    Public Shared Sub DeleteJob(ByVal pJTXDatabase As IJTXDatabase, ByVal jtxjob As IJTXJob, ByVal workingPath As String, ByVal log As System.IO.StreamWriter)
        Try
            pJTXDatabase.JobManager.DeleteJob(jtxjob.ID)
            log.WriteLine(Now & ": job " & jtxjob.ID & " has been removed")
            If File.exists(workingPath & "\jtxjob.txt") Then
                File.Delete(workingPath & "\jtxjob.txt")
                log.WriteLine(Now & ": jtxjob.txt has been removed")
            End If
        Catch ex As Exception
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())
        End Try
    End Sub
End Class
