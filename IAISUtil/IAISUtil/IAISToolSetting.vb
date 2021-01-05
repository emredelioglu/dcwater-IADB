Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms
Imports IAIS.Windows.UI
Imports System.IO
Imports Microsoft.VisualBasic.CompilerServices
Imports System.Reflection
Imports ESRI.ArcGIS.JTX
Imports Microsoft.Win32

Public Class IAISToolSetting
    Private Shared ParaMap As Dictionary(Of String, String)
    Private Shared ActiveForm As Form
    Private Shared ArcmapApplication As IApplication

    Public Shared Sub SetApplication(ByVal app As IApplication)
        ArcmapApplication = app
    End Sub

    Public Shared Function isInitilized() As Boolean
        Return (Not ParaMap Is Nothing)
    End Function


    Public Shared Function GetParameterValue(ByVal paraName As String, Optional ByRef pApp As IApplication = Nothing) As String
        If ParaMap Is Nothing Then
            If pApp Is Nothing Then
                If Not ArcmapApplication Is Nothing Then
                    If Not Initilize(ArcmapApplication, withJTXFlag:=False) Then
                        Throw New Exception("Can't initialize system setting table.")
                    End If
                Else
                    Throw New Exception("System setting table has not been initialized.")
                End If
            ElseIf Not Initilize(pApp) Then
                Throw New Exception("Can't initialize system setting table.")
            End If
        End If

        If ParaMap.ContainsKey(paraName) Then
            Return ParaMap.Item(paraName)
        Else
            Throw New Exception("Can't find system value for [" & paraName & "]")
        End If

    End Function

    Public Shared Sub UpdateParameter(ByVal paraName As String, ByVal paraNamenew As String, ByVal paraValue As String)
        If ParaMap Is Nothing Then
            Throw New Exception("system setting has not been initilized.")
        End If

        ParaMap.Remove(paraName)
        ParaMap.Add(paraNamenew, paraValue)
    End Sub

    Public Shared Sub AddParameter(ByVal paraName As String, ByVal paraValue As String)
        If ParaMap Is Nothing Then
            Throw New Exception("system setting has not been initilized.")
        End If

        ParaMap.Add(paraName, paraValue)
    End Sub

    Public Shared Sub DeleteParameter(ByVal paraName As String)
        If ParaMap Is Nothing Then
            Throw New Exception("system setting has not been initilized.")
        End If

        ParaMap.Remove(paraName)
    End Sub

	Public Shared Function Initilize(ByRef pApp As IApplication, Optional ByVal forceFlag As Boolean = False, Optional ByVal withJTXFlag As Boolean = True) As Boolean

		'If Not forceFlag AndAlso Not ParaMap Is Nothing Then
		'	Return True
		'End If

		'Dim logFile As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location()) + "\IAISToolSetting.log"
		Dim logFilePath As String = Environment.GetEnvironmentVariable("TEMP")
		If logFilePath Is Nothing Then
			logFilePath = Environment.GetEnvironmentVariable("TMP")
		End If
		If logFilePath Is Nothing Then
			logFilePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location())
		End If

		Dim logFile As String = logFilePath + "\IAISToolSetting.log"
		'Environment.GetEnvironmentVariable("") 
		Dim logStream As System.IO.StreamWriter = System.IO.File.AppendText(logFile)
		Try
			logStream.WriteLine(Now & ": Initilize system setting table with forceFlag=" & forceFlag & " and withJTXFlag=" & withJTXFlag)

			Dim registryKey As Microsoft.Win32.RegistryKey
			registryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\IAIS")
			Dim tnIAISSetting As String = registryKey.GetValue("SYSTEM_TABLE")

			If tnIAISSetting Is Nothing Or tnIAISSetting = "" Then
				Throw New Exception("Invalid SYSTEM_TABLE value in system registry")
			End If

			logStream.WriteLine(Now & ": System table name - " & tnIAISSetting)

			Dim pDoc As IMxDocument
			pDoc = pApp.Document

			If ParaMap Is Nothing Then
				ParaMap = New Dictionary(Of String, String)
			Else
				ParaMap.Clear()
			End If


			Dim pTable As ITable = Nothing

			If withJTXFlag Then
				Dim jtxExt As ESRI.ArcGIS.JTXExt.JTXExtension = DirectCast(pApp.FindExtensionByName("Workflow Manager"), ESRI.ArcGIS.JTXExt.JTXExtension)
				If jtxExt.Job Is Nothing Then
					logStream.WriteLine(Now & ": function called with JTX Flag=True and no Job is open")
					If jtxExt.Job.VersionExists Then
						logStream.WriteLine(Now & ": Opening table with JTX data workspace " & jtxExt.Database.Alias)
						pTable = CType(jtxExt.Database.DataWorkspace(jtxExt.Job.VersionName), IFeatureWorkspace).OpenTable(tnIAISSetting)
						logStream.WriteLine(Now & ": pTable" & pTable.OIDFieldName)
					Else
						logStream.WriteLine(Now & ": function called with JTX Flag=True and no version is available for JTX job")
						Return Initilize(pApp, forceFlag, withJTXFlag)


					End If

				End If

			End If

			If pTable Is Nothing Then
				Dim pLayer As IFeatureLayer = MapUtil.GetLayerByTableName("GISDBP.IADB.PremsInterPt", pDoc.FocusMap)
				logStream.WriteLine(Now & ": pLayer" & pLayer.Name)
				If pLayer Is Nothing Then
					pLayer = MapUtil.GetLayerByTableName("GISDBP.IADB.BaseOwnerPly", pDoc.FocusMap)
					logStream.WriteLine(Now & ": pLayer2" & pLayer.Name)
					If pLayer Is Nothing Then
						Throw New Exception("Can't connect to system table. PremsInterPt or OwnerPly is not loaded.")
					End If
				End If

				Dim pDataset As IDataset = pLayer
				logStream.WriteLine(Now & ": Opening table with workspace from layers")
				pTable = CType(pDataset.Workspace, IFeatureWorkspace).OpenTable(tnIAISSetting)
			End If


			Dim pCursor As ICursor
			Dim pRow As IRow

			logStream.WriteLine(Now & ": Reading values from system setting table")
			pCursor = pTable.Search(Nothing, False)
			pRow = pCursor.NextRow
			Do While Not pRow Is Nothing
				logStream.WriteLine(Now & ": NAME=" & pRow.Value(pRow.Fields.FindField("NAME")) &
									" VALUE=<" & pRow.Value(pRow.Fields.FindField("VALUE")) & ">")
				ParaMap.Add(pRow.Value(pRow.Fields.FindField("NAME")), pRow.Value(pRow.Fields.FindField("VALUE")))
				pRow = pCursor.NextRow
			Loop

			ArcmapApplication = pApp

			logStream.WriteLine(Now & ": System setting table initilized")

			Return True
		Catch ex As Exception
			If Not ParaMap Is Nothing Then
				ParaMap.Clear()
				ParaMap = Nothing
			End If

			logStream.WriteLine(Now & ":" & ex.Message & vbCrLf & ex.StackTrace())
			Return False
		Finally
			logStream.Close()
		End Try

	End Function
	Public Shared Sub OpenForm(ByVal pForm As Form)
        If Not ActiveForm Is Nothing AndAlso ActiveForm.Visible Then
            MsgBox("Another editing form is open. Please finish the editing first.", Title:="IAIS Tools")
            Return
        End If

        ActiveForm = pForm
        ActiveForm.Show(New ModelessDialog(ArcmapApplication.hWnd))
    End Sub

    Public Shared Sub CloseActiveForm(Optional ByVal pForm As Form = Nothing)
        If Not pForm Is Nothing Then
            pForm.Close()
        ElseIf Not ActiveForm Is Nothing Then
            ActiveForm.Close()
            ActiveForm = Nothing
        End If
    End Sub

    Public Shared Function GetActiveForm() As Form
        Return ActiveForm
    End Function

End Class
