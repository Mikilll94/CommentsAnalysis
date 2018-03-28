using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace RoslynPlay
{
    class MethodSyntaxWalker : CSharpSyntaxWalker
    {
        private string _fileName;
        private CommentLocationStore _commentLocationStore;

        public MethodSyntaxWalker(string fileName, CommentLocationStore commentLocationStore)
        {
            _fileName = fileName;
            _commentLocationStore = commentLocationStore;
        }

        public override void VisitMethodDeclaration(MethodDeclarationSyntax node)
        {
            int startLine = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
            int endLine = node.GetLocation().GetLineSpan().EndLinePosition.Line + 1;

            MethodStore.Methods.Add(new Method()
            {
                Name = node.Identifier.ToString(),
                FileName = _fileName,
                LineNumber = startLine,
                LineEnd = endLine,
            });

            _commentLocationStore.AddCommentLocation(startLine - 1, "method_description");
            _commentLocationStore.AddCommentLocation(startLine, "method_start");
            _commentLocationStore.AddCommentLocation(startLine + 1, endLine - 1, "method_inner");
            _commentLocationStore.AddCommentLocation(endLine, "method_end");
            base.VisitMethodDeclaration(node);
        }
    }
}
