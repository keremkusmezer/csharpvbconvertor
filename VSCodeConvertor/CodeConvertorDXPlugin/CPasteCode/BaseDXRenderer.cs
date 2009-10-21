using CPasterCode.Interfaces;
using DevExpress.CodeRush.StructuralParser;
using DevExpress.CodeRush.Core;
public abstract class BaseDXRenderer : IDXRenderer
{
    // Methods
    public string Render(LanguageElement RootNode)
    {
        return CodeRush.Language.GenerateElement(RootNode, this.LanguageID);
    }

    public string RenderToEditor(LanguageElement RootNode)
    {
        string str = this.Render(RootNode);
        CodeRush.Documents.ActiveTextView.Selection.Text = str;
        return str;
    }

    //// Properties
    //public abstract string CR_Paste.IDXOperable.LanguageID { get; }

    //public abstract string LanguageID { get; }

    #region IDXOperable Members

    public abstract string LanguageID
    {
        get;
    }

    #endregion
}


