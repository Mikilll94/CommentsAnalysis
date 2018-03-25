using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using System.Text.RegularExpressions;

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
            if (trivia.Kind() == SyntaxKind.SingleLineCommentTrivia)
            {
                string content = trivia.ToString();
                CommentsStore.Comments.Add(new Comment(content.Substring(content.IndexOf("//") + 2))
                {
                    LineNumber = trivia.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                    Type = "single_line_comment",
                    FileName = visitedFile,
                });
            }
            else if (trivia.Kind() == SyntaxKind.MultiLineCommentTrivia)
            {
                string content = trivia.ToString();
                content = new Regex(@"\/\*(.*)\*\/", RegexOptions.Singleline).Match(content).Groups[1].ToString();
                CommentsStore.Comments.Add(new Comment(content)
                {
                    LineNumber = trivia.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                    Type = "multi_line_comment",
                    FileName = visitedFile,
                });
            }
            base.VisitTrivia(trivia);
        }

        public override void Visit(SyntaxNode node)
        {
            if (node.Kind() == SyntaxKind.SingleLineDocumentationCommentTrivia)
            {
                string content = node.ToString();
                content = new Regex(@"(\/\/\/)").Replace(content, "");
                CommentsStore.Comments.Add(new Comment(content)
                {
                    LineNumber = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                    Type = "single_line_doc",
                    FileName = visitedFile,
                });
            }
            else if (node.Kind() == SyntaxKind.MultiLineDocumentationCommentTrivia)
            {
                string content = node.ToString();
                content = new Regex(@"(\/\/\/)").Replace(content, "");
                CommentsStore.Comments.Add(new Comment(node.ToString())
                {
                    LineNumber = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                    Type = "multi_line_doc",
                    FileName = visitedFile,
                });
            }
            base.Visit(node);
        }
    }
}
