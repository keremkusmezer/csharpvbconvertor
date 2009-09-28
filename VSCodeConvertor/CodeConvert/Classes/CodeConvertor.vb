#Region "Imported Libraries"
Imports System.Reflection
Imports System
Imports EnvDTE
Imports EnvDTE80
Imports System.Diagnostics
#End Region
Public Module CodeConvertor
    ''' <summary>
    ''' Converts to csharp and vb.
    ''' </summary>
    Sub ConvertToCsharpAndVb()
        Dim targetBody As String = String.Empty
        If Connect.DTE.ActiveDocument IsNot Nothing Then            
            Dim currentSelection As TextSelection = CType(Connect.DTE.ActiveDocument.Selection, TextSelection)
            currentSelection.SelectAll()
            targetBody = currentSelection.Text
            Dim ch As New ClipboardHelper()
            If Connect.DTE.ActiveDocument.Language = "Basic" Then
                Connect.DTE.ItemOperations.NewFile("General\Visual C# Class")
                currentSelection = CType(Connect.DTE.ActiveDocument.Selection, TextSelection)
                currentSelection.SelectAll()
                currentSelection.Cut()
                ch.SetClipboardText(ConversionLoader.ConvertVBToCSharp(targetBody))
                currentSelection.Paste()
            ElseIf Connect.DTE.ActiveDocument.Language = "CSharp" Then
                Connect.DTE.ItemOperations.NewFile("General\Visual Basic Class")
                currentSelection = CType(Connect.DTE.ActiveDocument.Selection, TextSelection)
                currentSelection.SelectAll()
                currentSelection.Delete()
                ch.SetClipboardText(ConversionLoader.ConvertCSharpToVB(targetBody))
                currentSelection.Paste()
            End If
        End If
    End Sub
    ''' <summary>
    ''' Using Clipboard In Visual Studio Addons Is Problematic There This Class Does
    ''' Crete A New Single Thread Apartment Model Thread, And Runs The Clipboard Request Inside It
    ''' </summary>
    Public Class ClipboardHelper
        ''' <summary>
        ''' Sets the clipboard text.
        ''' </summary>
        ''' <param name="sourceText">The source text.</param>
        Public Sub SetClipboardText(ByVal sourceText As String)
            Dim setterThread As New Threading.Thread(AddressOf SetClipboardTextInternal)
            With setterThread
                .SetApartmentState(Threading.ApartmentState.STA)
                .IsBackground = True
                .Start(sourceText)
                .Join()
            End With
        End Sub
        ''' <summary>
        ''' Sets the clipboard text internal.
        ''' </summary>
        ''' <param name="sourceText">The source text.</param>
        Private Sub SetClipboardTextInternal(ByVal sourceText As Object)
            System.Windows.Forms.Clipboard.SetText(CType(sourceText, String))
        End Sub
    End Class
    ''' <summary>
    ''' Removes the bar if exists.
    ''' </summary>
    ''' <param name="barName">Name of the bar.</param>
    ''' <param name="sourceBar">The source bar.</param>
    Public Sub RemoveBarIfExists(ByVal barName As String, ByVal sourceBar As Microsoft.VisualStudio.CommandBars.CommandBar)
        For i As Integer = sourceBar.Controls.Count To 1 Step -1
            If sourceBar.Controls(i).Caption = barName Then
                sourceBar.Controls(i).Delete()
                Continue For
            End If
            If String.IsNullOrEmpty(sourceBar.Controls(i).Caption) Then
                sourceBar.Controls(i).Delete()
                Continue For
            End If
        Next
    End Sub
    ''' <summary>
    ''' Checks the bar for existence.
    ''' </summary>
    ''' <param name="barName">Name of the bar.</param>
    ''' <param name="sourceBar">The source bar.</param>
    ''' <returns></returns>
    Public Function CheckBarForExistence(ByVal barName As String, ByVal sourceBar As Microsoft.VisualStudio.CommandBars.CommandBar) As Microsoft.VisualStudio.CommandBars.CommandBarButton
        Dim sourceButtonItem As Microsoft.VisualStudio.CommandBars.CommandBarButton = Nothing
        Dim foundAlready As Boolean
        For i As Integer = sourceBar.Controls.Count To 1 Step -1
            If sourceBar.Controls(i).Caption = barName Then
                If Not foundAlready Then
                    foundAlready = True
                    sourceButtonItem = CType(sourceBar.Controls(i), Microsoft.VisualStudio.CommandBars.CommandBarButton)
                Else
                    sourceBar.Controls(i).Delete()
                    Continue For
                End If
            End If
            If String.IsNullOrEmpty(sourceBar.Controls(i).Caption) Then
                sourceBar.Controls(i).Delete()
                Continue For
            End If
        Next
        If sourceButtonItem Is Nothing Then
            sourceButtonItem = CType(sourceBar.Controls.Add(Microsoft.VisualStudio.CommandBars.MsoControlType.msoControlButton), Microsoft.VisualStudio.CommandBars.CommandBarButton)
            sourceButtonItem.Caption = barName
        End If
        Return sourceButtonItem
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    Public Class ConversionLoader
        ''' <summary>
        ''' Converts the C sharp to VB.
        ''' </summary>
        ''' <param name="sourceCode">The source code.</param>
        ''' <returns></returns>
        Public Shared Function ConvertCSharpToVB(ByVal sourceCode As String) As String
            Return ConversionWrapper.CodeConversionHelper.ConvertCSharpToVB(sourceCode)
        End Function
        ''' <summary>
        ''' Converts the VB to C sharp.
        ''' </summary>
        ''' <param name="sourceCode">The source code.</param>
        ''' <returns></returns>
        Public Shared Function ConvertVBToCSharp(ByVal sourceCode As String) As String
            Return ConversionWrapper.CodeConversionHelper.ConvertVBToCSharp(sourceCode)
        End Function
    End Class
End Module