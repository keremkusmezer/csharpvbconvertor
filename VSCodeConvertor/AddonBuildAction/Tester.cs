using System;
using System.Collections.Generic;
using System.Text;

namespace AddonBuildAction
{
    class Tester
    {
        public static void Main(string[] args)
        {
            IEnumerator<InfoWrapper> wrapperStorage =
                DetectOpenFiles.GetOpenFilesEnumerator(2316);
            while (wrapperStorage.MoveNext())
            {
                if (wrapperStorage.Current.WrappedObject.FullName.Contains("CSF.DAL"))
                {
                    wrapperStorage.Current.CloseHandle();
                }
            }
        }        
    }
}
