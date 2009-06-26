Imports System.Runtime.InteropServices
Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Text
Imports System.Reflection
Imports System.Diagnostics
Namespace AddonBuildAction
    <System.ComponentModel.RunInstaller(True)> _
    Public Class AddInstaller
        Inherits System.Configuration.Install.Installer
        Private Const AddinSubPath As String = "Application Data\Microsoft\MSEnvShared\AddIns"
        Private Const AllUsersProfile As String = "ALLUSERSPROFILE"
        <DllImport("ntdll.dll", CharSet:=CharSet.Auto)> _
        Private Shared Function NtQuerySystemInformation(ByVal query As Integer, ByVal dataPtr As IntPtr, ByVal size As Integer, <Out()> ByRef returnedSize As Integer) As Integer
        End Function
        Private Function IsDevelopmentEnvironmentRunning() As Boolean
            'Retrieves The CommandLine Parameters
            'Context
        End Function
        Private Function IsVista() As Boolean
            Return System.Environment.OSVersion.Version.Major >= 6
        End Function
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
        Protected Overrides Sub OnAfterInstall(ByVal savedState As System.Collections.IDictionary)
            MyBase.OnAfterInstall(savedState)
            If Not IsVista() Then
                Try
                    Dim allusersProfilePath As String = _
                        Path.Combine(Environment.GetEnvironmentVariable(AllUsersProfile), AddinSubPath)
                    Dim assemblyPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                    For Each tempFile As String In Directory.GetFiles(assemblyPath)
                        Dim targetPath As String = Path.Combine(allusersProfilePath, Path.GetFileName(tempFile))
                        Try
                            File.Copy(tempFile, targetPath, True)
                        Catch ex As Exception
                        End Try
                    Next
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine(ex.Message)
                End Try
            End If
        End Sub
    End Class
End Namespace