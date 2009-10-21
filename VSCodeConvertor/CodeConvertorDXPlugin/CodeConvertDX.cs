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
          // Fields
    [AccessedThroughProperty("PasteSmartTag")]
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
        //this.PasteSmartTag.BeginInit();
        this.BeginInit();
        this.PasteSmartTag.Description = "Paste";// ("Paste");
        this.PasteSmartTag.DisplayName = "Paste";
        this.PasteSmartTag.MenuOrder = 0;
        this.PasteSmartTag.ProviderName = "Paste";
        this.PasteSmartTag.Register = true;
        this.PasteSmartTag.ShowInContextMenu = true;
        this.PasteSmartTag.ShowInPopupMenu = false;
        //this.PasteSmartTag.EndInit();
        //this.EndInit();
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
        SmartTagItem item2 = new SmartTagItem("Smart Paste");
        SmartTagItem item = new SmartTagItem("CSharp -> VBNet");
        SmartTagItem item3 = new SmartTagItem("VBNet -> CSharp");
        ea.Add(item2);
        ea.Add(item);
        ea.Add(item3);
        item2.Execute += new EventHandler(this.PasteCSharpAsVBNet);
        //item2.add_Execute();
        item.Execute += (new EventHandler(this.PasteCSharpAsVBNet));
        item3.Execute += (new EventHandler(this.PasteVBNetAsCSharp));
    }

    private void PasteVBNetAsCSharp(object sender, EventArgs e)
    {
        new GenericDXTranslator("CSharp", "Basic").Paste();
    }

    private void SmartPaste(object sender, EventArgs e)
    {
        new GenericDXTranslator(CodeRush.Documents.ActiveLanguage, "").Paste();
    }

    // Properties
    internal virtual SmartTagProvider PasteSmartTag
    {
        get
        {
            return this._PasteSmartTag;
        }
        [MethodImpl(MethodImplOptions.Synchronized)]
        set
        {
            GetSmartTagItemsEventHandler handler = 
                new GetSmartTagItemsEventHandler(PasteSmartTag_GetSmartTagItems);
            if (this._PasteSmartTag != null)
            {
                this._PasteSmartTag.GetSmartTagItems -= handler;
                //this._PasteSmartTag.remove_GetSmartTagItems(handler);
            }
            this._PasteSmartTag = value;
            if (this._PasteSmartTag != null)
            {
                this._PasteSmartTag.GetSmartTagItems += (handler);
            }
        }
    }
}
}