#Region "Imported Libraries"
Imports System.Reflection
Imports System
Imports EnvDTE
Imports EnvDTE80
Imports System.Diagnostics
#End Region
Public Module CodeConvertor
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
    Public Class ClipboardHelper
        Public Sub SetClipboardText(ByVal sourceText As String)
            Dim setterThread As New Threading.Thread(AddressOf SetClipboardTextInternal)
            With setterThread
                .SetApartmentState(Threading.ApartmentState.STA)
                .IsBackground = True
                .Start(sourceText)
                .Join()
            End With
        End Sub
        Private Sub SetClipboardTextInternal(ByVal sourceText As Object)
            System.Windows.Forms.Clipboard.SetText(CType(sourceText, String))
        End Sub
    End Class
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
    Public Class ConversionLoader
        ''Private Shared ConvertToCsharpDelegate As ConvertLanguage
        ''Private Shared ConvertToVbDelegate As ConvertLanguage
        ''Private Delegate Function ConvertLanguage(ByVal source As String) As String
        'Shared Sub New()
        '    'Dim currentFile As System.Reflection.Assembly = Assembly.Load(Convert.FromBase64String(AssemblyStorage.RetrieveAssembly()))
        '    'Dim convertorType As Type = currentFile.GetType("ConversionWrapper.CodeConversionHelper", False, True)
        '    'If convertorType IsNot Nothing Then
        '    '    'Dim firstDelegate As [Delegate]
        '    '    Dim sourceMethod As MethodInfo = convertorType.GetMethod("ConvertCSharpToVB", BindingFlags.Static Or BindingFlags.Public)
        '    '    ConvertToVbDelegate = CType([Delegate].CreateDelegate(GetType(ConvertLanguage), sourceMethod), ConvertLanguage)
        '    '    sourceMethod = convertorType.GetMethod("ConvertVBToCSharp", BindingFlags.Static Or BindingFlags.Public)
        '    '    ConvertToCsharpDelegate = CType([Delegate].CreateDelegate(GetType(ConvertLanguage), sourceMethod), ConvertLanguage)
        '    'End If
        'End Sub
        Public Shared Function ConvertCSharpToVB(ByVal sourceCode As String) As String
            Return ConversionWrapper.CodeConversionHelper.ConvertCSharpToVB(sourceCode)
        End Function
        Public Shared Function ConvertVBToCSharp(ByVal sourceCode As String) As String
            Return ConversionWrapper.CodeConversionHelper.ConvertVBToCSharp(sourceCode)
        End Function
    End Class
End Module