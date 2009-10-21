public class GenericDXRenderer : BaseDXRenderer
{
    // Fields
    protected string mLanguageID;

    // Methods
    public GenericDXRenderer(string LanguageID)
    {
        this.mLanguageID = LanguageID;
    }

    // Properties
    public override string LanguageID
    {
        get
        {
            return this.mLanguageID;
        }
    }
}


