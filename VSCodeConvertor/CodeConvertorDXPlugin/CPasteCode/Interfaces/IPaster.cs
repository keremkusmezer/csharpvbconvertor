using DevExpress.CodeRush.StructuralParser;
namespace CPasterCode.Interfaces
{
    public interface IPaster
    {
        // Methods
        void Paste();
    }
    public interface IDXRenderer : IDXOperable
    {
        // Methods
        string Render(LanguageElement RootNode);
        string RenderToEditor(LanguageElement RootNode);
    }
    public interface IDXOperable
    {
        // Properties
        string LanguageID { get; }
    }
    public interface IDXLoader : IDXOperable
    {
        // Methods
        bool Load(string Text);

        // Properties
        LanguageElement TreeRoot { get; }
    }
}