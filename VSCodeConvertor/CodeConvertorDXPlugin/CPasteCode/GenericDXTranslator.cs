using CPasterCode.Interfaces;
using System.Windows.Forms;
public class GenericDXTranslator : IPaster
{
    // Fields
    protected IDXLoader Loader;
    protected IDXRenderer Renderer;
    public GenericDXTranslator(string RendererLanguageID)
        : this(RendererLanguageID, string.Empty)
    {

    }
    // Methods
    public GenericDXTranslator(string RendererLanguageID, string LoaderLanguageID)
    {
        this.Loader = new GenericDXLoader(LoaderLanguageID);
        this.Renderer = new GenericDXRenderer(RendererLanguageID);
    }

    public void Paste()
    {
        if (Clipboard.ContainsText() && this.Loader.Load(Clipboard.GetText()))
        {
            this.Renderer.RenderToEditor(this.Loader.TreeRoot);
        }
    }
}


