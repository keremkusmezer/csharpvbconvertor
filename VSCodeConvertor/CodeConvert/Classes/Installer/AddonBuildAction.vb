#Region "Imported Libraries"
Imports System.Runtime.InteropServices
Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Text
Imports System.Reflection
Imports System.Diagnostics
#End Region
Namespace AddonBuildAction
    ''' <summary>
    ''' 
    ''' </summary>
    <System.ComponentModel.RunInstaller(True)> _
    Public Class AddInstaller
        Inherits System.Configuration.Install.Installer
        Private Const AddinSubPath As String = "Application Data\Microsoft\MSEnvShared\AddIns"
        Private Const Addin2005Path As String = "Visual Studio 2005\AddIns"
        Private Const Addin2008Path As String = "Visual Studio 2008\AddIns"
        Private Const AllUsersProfile As String = "ALLUSERSPROFILE"
        ''' <summary>
        ''' Nts the query system information.
        ''' </summary>
        ''' <param name="query">The query.</param>
        ''' <param name="dataPtr">The data PTR.</param>
        ''' <param name="size">The size.</param>
        ''' <param name="returnedSize">Size of the returned.</param>
        ''' <returns></returns>
        <DllImport("ntdll.dll", CharSet:=CharSet.Auto)> _
        Private Shared Function NtQuerySystemInformation(ByVal query As Integer, ByVal dataPtr As IntPtr, ByVal size As Integer, <Out()> ByRef returnedSize As Integer) As Integer
        End Function
        ''' <summary>
        ''' Determines whether [is development environment running].
        ''' </summary>
        ''' <returns>
        ''' <c>true</c> if [is development environment running]; otherwise, <c>false</c>.
        ''' </returns>
        Private Function IsDevelopmentEnvironmentRunning() As Boolean
            'Retrieves The CommandLine Parameters
            'Context
        End Function
        ''' <summary>
        ''' Determines whether this instance is vista.
        ''' </summary>
        ''' <returns>
        ''' <c>true</c> if this instance is vista; otherwise, <c>false</c>.
        ''' </returns>
        Private Function IsVista() As Boolean
            Return System.Environment.OSVersion.Version.Major >= 6
        End Function
        ''' <summary>
        ''' When overridden in a derived class, removes an installation.
        ''' </summary>
        ''' <param name="savedState">An <see cref="T:System.Collections.IDictionary" /> that contains the state of the computer after the installation was complete.</param>
        ''' <exception cref="T:System.ArgumentException">
        ''' The saved-state <see cref="T:System.Collections.IDictionary" /> might have been corrupted.
        ''' </exception>
        ''' <exception cref="T:System.Configuration.Install.InstallException">
        ''' An exception occurred while uninstalling. This exception is ignored and the uninstall continues. However, the application might not be fully uninstalled after the uninstallation completes.
        ''' </exception>
        Public Overrides Sub Uninstall(ByVal savedState As System.Collections.IDictionary)
            MyBase.Uninstall(savedState)
            Dim allusersProfilePath As String = _
                    Environment.GetEnvironmentVariable(AllUsersProfile)
            allusersProfilePath = Path.Combine(allusersProfilePath, AddinSubPath)
            Dim assemblyPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            For Each tempFile As String In Directory.GetFiles(assemblyPath)
                Try
                    File.Delete(Path.Combine(allusersProfilePath, Path.GetFileName(tempFile)))
                Catch ex As Exception
                End Try
            Next
        End Sub
        ''' <summary>
        ''' Raises the <see cref="E:System.Configuration.Install.Installer.AfterInstall" /> event.
        ''' </summary>
        ''' <param name="savedState">An <see cref="T:System.Collections.IDictionary" /> that contains the state of the computer after all the installers contained in the <see cref="P:System.Configuration.Install.Installer.Installers" /> property have completed their installations.</param>
        Protected Overrides Sub OnAfterInstall(ByVal savedState As System.Collections.IDictionary)
            MyBase.OnAfterInstall(savedState)
            Dim allusersProfilePath As String = String.Empty
            If Not IsVista() Then
                Try
                    allusersProfilePath = _
                        Path.Combine(Environment.GetEnvironmentVariable(AllUsersProfile), AddinSubPath)
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine(ex.Message)
                End Try
            Else
                'This Handles Vista Based Install Paths
                allusersProfilePath = _
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Addin2005Path)
                CopyToTargetPath(allusersProfilePath)
                allusersProfilePath = _
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Addin2008Path)
            End If
            CopyToTargetPath(allusersProfilePath)
        End Sub
        ''' <summary>
        ''' Copies to target path.
        ''' </summary>
        ''' <param name="allusersProfilePath">The allusers profile path.</param>
        Private Sub CopyToTargetPath(ByVal allusersProfilePath As String)
            If Not (Directory.Exists(allusersProfilePath)) Then
                Directory.CreateDirectory(allusersProfilePath)
            End If
            Dim assemblyPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            For Each tempFile As String In Directory.GetFiles(assemblyPath)
                Dim targetPath As String = Path.Combine(allusersProfilePath, Path.GetFileName(tempFile))
                Try
                    File.Copy(tempFile, targetPath, True)
                Catch ex As Exception
                End Try
            Next
        End Sub
    End Class
End Namespace