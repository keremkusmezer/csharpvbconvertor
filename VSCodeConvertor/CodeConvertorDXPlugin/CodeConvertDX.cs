#region Imported Libraries
using System;
using System.ComponentModel;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.CodeRush.Core;
using DevExpress.CodeRush.PlugInCore;
using DevExpress.CodeRush.StructuralParser;
using System.Runtime.CompilerServices;
using System.Diagnostics;
#endregion
namespace CodeConvertorDXPlugin
{
    public partial class CodeConvertDX : StandardPlugIn
    {

        //public SmartTagProvider PasteSmartTag
        //{
        //    get
        //    {
        //        return _PasteSmartTag;
        //    }
        //    set
        //    {
        //        _PasteSmartTag = value;
        //    }
        //}

    private SmartTagProvider _PasteSmartTag;
    private IContainer components;

    // Methods
    [DebuggerNonUserCode]
    public CodeConvertDX()
    {
        this.InitializeComponent();
    }

    [DebuggerNonUserCode]
    protected override void Dispose(bool disposing)
    {
        if (disposing && (this.components != null))
        {
            this.components.Dispose();
        }
        base.Dispose(disposing);
    }

    public override void FinalizePlugIn()
    {
        base.FinalizePlugIn();
    }

    [DebuggerStepThrough]
    private void InitializeComponent()
    {
        this.components = new Container();
        this.PasteSmartTag = new SmartTagProvider(this.components);
        
        ((ISupportInitialize)this.PasteSmartTag).BeginInit();
        this.BeginInit();
        this.PasteSmartTag.Description = "Paste";// ("Paste");
        this.PasteSmartTag.DisplayName = "Paste";
        this.PasteSmartTag.MenuOrder = 0;
        this.PasteSmartTag.ProviderName = "Paste";
        this.PasteSmartTag.Register = true;
        this.PasteSmartTag.ShowInContextMenu = true;
        this.PasteSmartTag.ShowInPopupMenu = false;
        ((ISupportInitialize)this.PasteSmartTag).EndInit();
        this.PasteSmartTag.GetSmartTagItems += new GetSmartTagItemsEventHandler(PasteSmartTag_GetSmartTagItems);
        this.EndInit();
    }

    public override void InitializePlugIn()
    {
        base.InitializePlugIn();
    }

    private void PasteCSharpAsVBNet(object sender, EventArgs e)
    {
        new GenericDXTranslator("Basic", "CSharp").Paste();
    }

    private void PasteSmartTag_GetSmartTagItems(object sender, GetSmartTagItemsEventArgs ea)
    {
        SmartTagItem item2 = new SmartTagItemEx("Smart Paste", new MethodInvoker(SmartPaste));
        SmartTagItem item = new SmartTagItemEx("CSharp -> VBNet", new MethodInvoker(PasteCSharpAsVBNet));
        SmartTagItem item3 = new SmartTagItemEx("VBNet -> CSharp", new MethodInvoker(PasteVBNetAsCSharp));
        ea.Add(item2);
        ea.Add(item);
        ea.Add(item3);
    }

    /// <summary>
    /// Pastes the C sharp as VB net.
    /// </summary>
    private void PasteCSharpAsVBNet()
    {
        new GenericDXTranslator("CSharp", "Basic").Paste();
    }
    /// <summary>
    /// Pastes the VB net as C sharp.
    /// </summary>
    private void PasteVBNetAsCSharp()
    {
        new GenericDXTranslator("CSharp", "Basic").Paste();
    }

    /// <summary>
    /// Smarts the paste.
    /// </summary>
    private void SmartPaste()
    {
        new GenericDXTranslator(CodeRush.Documents.ActiveLanguage, "").Paste();
    }
    private DevExpress.CodeRush.Core.SmartTagProvider PasteSmartTag;

}
}