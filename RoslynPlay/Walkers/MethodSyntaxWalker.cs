using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace RoslynPlay
{
    class MethodSyntaxWalker : CSharpSyntaxWalker
    {
        string visitedFile;

        public MethodSyntaxWalker(string fileName) : base(SyntaxWalkerDepth.StructuredTrivia)
        {
            visitedFile = fileName;
        }

        public override void VisitMethodDeclaration(MethodDeclarationSyntax node)
        {
            int startLine = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
            int endLine = node.GetLocation().GetLineSpan().EndLinePosition.Line + 1;

            MethodStore.Methods.Add(new Method()
            {
                Name = node.Identifier.ToString(),
                FileName = visitedFile,
                LineNumber = startLine,
                LineEnd = endLine,
            });

            CommentLocationStore.CommentLocations.TryAdd(startLine - 1, $"method_description {visitedFile}");
            CommentLocationStore.CommentLocations.TryAdd(startLine, $"method_start {visitedFile}");
            CommentLocationStore.AddCommentLocation(startLine + 1, endLine - 1, $"method_inner {visitedFile}");
            CommentLocationStore.CommentLocations.TryAdd(endLine, $"method_end {visitedFile}");
            base.VisitMethodDeclaration(node);
        }
    }
}
