using DevExpress.CodeRush.Core;
using System.Windows.Forms;
public class SmartTagItemEx : SmartTagItem
{
    private MethodInvoker mExecute;
    public SmartTagItemEx(string Caption, MethodInvoker OnExecute)
    {
        base.Caption = Caption;
        mExecute = OnExecute;
        this.Execute += new System.EventHandler(SmartTagItemEx_Execute);
    }
    private void  SmartTagItemEx_Execute(object sender, System.EventArgs e)
    {
        mExecute.Invoke();
    }
}
