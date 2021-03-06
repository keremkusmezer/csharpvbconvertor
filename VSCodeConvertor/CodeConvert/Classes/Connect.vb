#Region "Imported Libraries"
Imports System
Imports Microsoft.VisualStudio.CommandBars
Imports Extensibility
Imports EnvDTE
Imports EnvDTE80
#End Region
Public Class Connect
    Implements IDTExtensibility2

#Region "Private Variables"
    Private isHandlersSet As Boolean
    Private _addInInstance As AddIn
    Public WithEvents VBToCSharpBar As Microsoft.VisualStudio.CommandBars.CommandBarButton
    Public WithEvents CSharpToVBBar As Microsoft.VisualStudio.CommandBars.CommandBarButton
    Public WithEvents CodeWindow As Microsoft.VisualStudio.CommandBars.CommandBar
    Public WithEvents WindowEvents As EnvDTE.WindowEvents
    Public Shared DTE As DTE2
#End Region

#Region "Addin Methods"
    '''<summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
    Public Sub New()
    End Sub
    '''<summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
    '''<param name='application'>Root object of the host application.</param>
    '''<param name='connectMode'>Describes how the Add-in is being loaded.</param>
    '''<param name='addInInst'>Object representing this Add-in.</param>
    '''<remarks></remarks>
    Public Sub OnConnection(ByVal application As Object, ByVal connectMode As ext_ConnectMode, ByVal addInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
        DTE = CType(application, DTE2)
        _addInInstance = CType(addInInst, AddIn)
        WindowEvents = DTE.Events.WindowEvents
        CreateHandlers()
        If DTE.ActiveWindow IsNot Nothing Then
            HandleDocument(DTE.ActiveWindow)
        End If
    End Sub
    '''<summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
    '''<param name='disconnectMode'>Describes how the Add-in is being unloaded.</param>
    '''<param name='custom'>Array of parameters that are host application specific.</param>
    '''<remarks></remarks>
    Public Sub OnDisconnection(ByVal disconnectMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
        If disconnectMode = ext_DisconnectMode.ext_dm_SolutionClosed OrElse _
           disconnectMode = ext_DisconnectMode.ext_dm_HostShutdown OrElse _
           disconnectMode = ext_DisconnectMode.ext_dm_UserClosed Then
            RemoveHandlers()
        End If
    End Sub
    '''<summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification that the collection of Add-ins has changed.</summary>
    '''<param name='custom'>Array of parameters that are host application specific.</param>
    '''<remarks></remarks>
    Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
    End Sub
    '''<summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
    '''<param name='custom'>Array of parameters that are host application specific.</param>
    '''<remarks></remarks>
    Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
    End Sub
    '''<summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
    '''<param name='custom'>Array of parameters that are host application specific.</param>
    '''<remarks></remarks>
    Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
    End Sub
#End Region

#Region "Private Method Implementation"
    ''' <summary>
    ''' Retrieves the code window.
    ''' </summary>
    ''' <returns></returns>
    Private Function RetrieveCodeWindow() As Microsoft.VisualStudio.CommandBars.CommandBar
        If CodeWindow Is Nothing Then
            Dim wholeBars As Microsoft.VisualStudio.CommandBars.CommandBars = CType(DTE.CommandBars, Microsoft.VisualStudio.CommandBars.CommandBars)
            CodeWindow = wholeBars.Item("Code Window")
        End If
        Return CodeWindow
    End Function
    ''' <summary>
    ''' Removes the handlers.
    ''' </summary>
    Private Sub RemoveHandlers()
        If isHandlersSet Then
            CodeWindow = RetrieveCodeWindow()
            RemoveBarIfExists("VB To CSharp", CodeWindow)
            RemoveBarIfExists("CSharp To VB", CodeWindow)
            WindowEvents = Nothing
        End If
    End Sub
    ''' <summary>
    ''' Creates the handlers.
    ''' </summary>
    Public Sub CreateHandlers()
        If Not isHandlersSet Then
            CodeWindow = RetrieveCodeWindow()
            VBToCSharpBar = CheckBarForExistence("VB To CSharp", CodeWindow)
            CSharpToVBBar = CheckBarForExistence("CSharp To VB", CodeWindow)
            isHandlersSet = True
        End If
    End Sub
    ''' <summary>
    ''' Windows the events_ window activated.
    ''' </summary>
    ''' <param name="GotFocus">The got focus.</param>
    ''' <param name="LostFocus">The lost focus.</param>
    Private Sub WindowEvents_WindowActivated(ByVal GotFocus As EnvDTE.Window, ByVal LostFocus As EnvDTE.Window) Handles WindowEvents.WindowActivated
        HandleDocument(GotFocus)
    End Sub
    ''' <summary>
    ''' Windows the events_ window created.
    ''' </summary>
    ''' <param name="GotFocus">The got focus.</param>
    Private Sub WindowEvents_WindowCreated(ByVal GotFocus As EnvDTE.Window) Handles WindowEvents.WindowCreated
        HandleDocument(GotFocus)
    End Sub
    ''' <summary>
    ''' Handles the document.
    ''' </summary>
    ''' <param name="GotFocus">The got focus.</param>
    Private Sub HandleDocument(ByVal GotFocus As EnvDTE.Window)
        CreateHandlers()
        If GotFocus IsNot Nothing AndAlso GotFocus.Document IsNot Nothing Then
            If GotFocus.Document.Language = "Basic" Then
                CSharpToVBBar.Visible = False
                VBToCSharpBar.Visible = True
            ElseIf GotFocus.Document.Language = "CSharp" Then
                CSharpToVBBar.Visible = True
                VBToCSharpBar.Visible = False
            Else
                CSharpToVBBar.Visible = False
                VBToCSharpBar.Visible = False
            End If
        End If
    End Sub
    ''' <summary>
    ''' VBs to C sharp bar_ click.
    ''' </summary>
    ''' <param name="Ctrl">The CTRL.</param>
    ''' <param name="CancelDefault">if set to <c>true</c> [cancel default].</param>
    Private Sub VBToCSharpBar_Click(ByVal Ctrl As Microsoft.VisualStudio.CommandBars.CommandBarButton, ByRef CancelDefault As Boolean) Handles VBToCSharpBar.Click
        CodeConvertor.ConvertToCsharpAndVb()
    End Sub
    ''' <summary>
    ''' Cs the sharp to VB bar_ click.
    ''' </summary>
    ''' <param name="Ctrl">The CTRL.</param>
    ''' <param name="CancelDefault">if set to <c>true</c> [cancel default].</param>
    Private Sub CSharpToVBBar_Click(ByVal Ctrl As Microsoft.VisualStudio.CommandBars.CommandBarButton, ByRef CancelDefault As Boolean) Handles CSharpToVBBar.Click
        CodeConvertor.ConvertToCsharpAndVb()
    End Sub
#End Region
End Class