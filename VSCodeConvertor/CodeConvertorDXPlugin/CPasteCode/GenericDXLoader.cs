public class GenericDXLoader : BaseDXLoader
{
    public GenericDXLoader():this("")
    {
    }
    // Methods
    public GenericDXLoader(string LanguageID)
    {
        base.mLanguageID = LanguageID;
    }
}


