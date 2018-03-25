using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using System;

namespace RoslynPlay
{
    public class SyntaxWalker : CSharpSyntaxWalker
    {
        string visitedFile;

        public SyntaxWalker(string fileName) : base(SyntaxWalkerDepth.StructuredTrivia)
        {
            visitedFile = fileName; 
        }

        public override void VisitTrivia(SyntaxTrivia trivia)
        {
            if (trivia.Kind() == SyntaxKind.SingleLineCommentTrivia ||
                trivia.Kind() == SyntaxKind.MultiLineCommentTrivia)
            {
                CommentsStore.Comments.Add(new Comment(trivia.ToString())
                {
                    LineNumber = trivia.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                    Type = "single_line_comment",
                    FileName = visitedFile,
                });
            }
            base.VisitTrivia(trivia);
        }

        public override void Visit(SyntaxNode node)
        {
            if (node.Kind() == SyntaxKind.SingleLineDocumentationCommentTrivia)
            {
                CommentsStore.Comments.Add(new Comment(node.ToString())
                {
                    LineNumber = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                    Type = "doc_comment",
                    FileName = visitedFile,
                });
            }
            base.Visit(node);
        }
    }
}
