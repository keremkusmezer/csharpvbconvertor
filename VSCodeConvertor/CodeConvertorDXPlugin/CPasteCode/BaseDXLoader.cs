using CPasterCode.Interfaces;
using DevExpress.CodeRush.StructuralParser;
using DevExpress.CodeRush.Core;
public abstract class BaseDXLoader : IDXLoader
{
    // Fields
    protected string mLanguageID = string.Empty;
    protected LanguageElement RootNode;

    // Methods
    private void AllocateLanguageByExample(string Text)
    {
        if (this.isTextOfCodeType(Text, "Basic"))
        {
            this.mLanguageID = "Basic";
        }
        if (this.isTextOfCodeType(Text, "CSharp"))
        {
            this.mLanguageID = "CSharp";
        }
    }

    private bool isTextOfCodeType(string Text, string LanguageID)
    {
        LanguageElement element = CodeRush.Language.GetParserFromLanguageID(LanguageID).ParseString(Text);
        return (CodeRush.Language.GenerateElement(element) == Text);
    }

    public virtual bool Load(string Text)
    {
        if (this.LanguageID == string.Empty)
        {
            this.AllocateLanguageByExample(Text);
        }
        if (this.LanguageID == string.Empty)
        {
            return false;
        }
        this.RootNode = CodeRush.Language.GetParserFromLanguageID(this.LanguageID).ParseString(Text);
        return true;
    }

    // Properties
    LanguageElement IDXLoader.TreeRoot
    {
        get
        {
            return this.RootNode;
        }
    }

    string IDXOperable.LanguageID
    {
        get
        {
            return this.mLanguageID;
        }
    }

    protected string LanguageID
    {
        get
        {
            return this.mLanguageID;
        }
    }

    public LanguageElement TreeRoot
    {
        get
        {
            return this.RootNode;
        }
    }
}