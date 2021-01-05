Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Runtime.InteropServices

' Set 'RunInstaller' attribute to true.
<RunInstaller(True)> _
Public Class IAISInstaller
    Inherits Installer

    Public Sub New()
        MyBase.New()
    End Sub 'New

    ' Event handler for 'Committing' event.
    Private Sub MyInstaller_Committing(ByVal sender As Object, _
                                       ByVal e As InstallEventArgs)
        Console.WriteLine("")
        Console.WriteLine("Committing Event occured.")
        Console.WriteLine("")
    End Sub 'MyInstaller_Committing

    ' Event handler for 'Committed' event.
    Private Sub MyInstaller_Committed(ByVal sender As Object, _
                                      ByVal e As InstallEventArgs)
        Console.WriteLine("")
        Console.WriteLine("Committed Event occured.")
        Console.WriteLine("")
    End Sub 'MyInstaller_Committed

    ' Override the 'Install' method.
    Public Overrides Sub Install(ByVal stateSaver As System.Collections.IDictionary)
        MyBase.Install(stateSaver)
        Dim regsrv As New RegistrationServices
        regsrv.RegisterAssembly(MyBase.GetType().Assembly, AssemblyRegistrationFlags.SetCodeBase)
    End Sub

    ' Override the 'Uninstall' method.
    Public Overrides Sub Uninstall(ByVal savedState As System.Collections.IDictionary)
        MyBase.Uninstall(savedState)
        Dim regsrv As New RegistrationServices
        regsrv.UnregisterAssembly(MyBase.GetType().Assembly)
    End Sub

    ' Override the 'Commit' method.
    Public Overrides Sub Commit(ByVal savedState As IDictionary)
        MyBase.Commit(savedState)
    End Sub 'Commit

    ' Override the 'Rollback' method.
    Public Overrides Sub Rollback(ByVal savedState As IDictionary)
        MyBase.Rollback(savedState)
    End Sub 'Rollback

    Private Sub InitializeComponent()

    End Sub
End Class 'MyInstallerClass


