#region Imported Libraries
using System;
using System.Collections.Generic;
using System.Text;
using ICSharpCode.NRefactory.PrettyPrinter;
using ICSharpCode.NRefactory;
using System.IO;
using ICSharpCode.NRefactory.Parser;
#endregion
namespace ConversionWrapper
{
    public static class CodeConversionHelper
    {
        public static string ConvertVBToCSharp(string sourceCode)
        {
            if (sourceCode == null || sourceCode.Length == 0)
                throw new ArgumentNullException(sourceCode, "sourceCode");
            return GenerateCode(sourceCode, SupportedLanguage.VBNet);
        }
        public static string ConvertCSharpToVB(string sourceCode)
        {
            if (sourceCode == null || sourceCode.Length == 0)
                throw new ArgumentNullException(sourceCode, "sourceCode");
            return GenerateCode(sourceCode, SupportedLanguage.CSharp);
        }

        private static string GenerateCode(string sourceCode,SupportedLanguage language)
        {
            using (IParser parser = ParserFactory.CreateParser(language, new StringReader(sourceCode)))
            {
                parser.Parse();
                if (parser.Errors.Count == 0)
                {
                        IList<ISpecial> savedSpecialsList = new ISpecial[0];
                            IOutputAstVisitor targetVisitor;
                            if (language == SupportedLanguage.CSharp)
                                targetVisitor = new VBNetOutputVisitor();
                            else
                                targetVisitor = new CSharpOutputVisitor();
                        using (SpecialNodesInserter.Install(savedSpecialsList, targetVisitor))
                        {
                            parser.CompilationUnit.AcceptVisitor(targetVisitor, null);
                        }
                        return targetVisitor.Text;
                }
                StringBuilder errorBuilder = new StringBuilder();
                return parser.Errors.ErrorOutput;
            }
        }
    }
}
