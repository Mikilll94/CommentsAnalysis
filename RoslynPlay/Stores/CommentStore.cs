using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace RoslynPlay
{
    public class CommentStore
    {
        public List<Comment> Comments { get; } = new List<Comment>();

        public void AddCommentTrivia(SyntaxTrivia trivia,
            CommentLocationStore commentLocationstore, string fileName)
        {
            string content = trivia.ToString();
            if (trivia.Kind() == SyntaxKind.SingleLineCommentTrivia)
            {
                content = content.Substring(content.IndexOf("//") + 2);
            }
            else if (trivia.Kind() == SyntaxKind.MultiLineCommentTrivia)
            {
                content = new Regex(@"\/\*(.*)\*\/", RegexOptions.Singleline).Match(content).Groups[1].ToString();
            }
            Comments.Add(new Comment(content, trivia.GetLocation().GetLineSpan().EndLinePosition.Line + 1,
                commentLocationstore)
            {
                LineNumber = trivia.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                Type = trivia.Kind().ToString(),
                FileName = fileName,
            });
        }

        public void AddCommentNode(DocumentationCommentTriviaSyntax node,
            CommentLocationStore commentLocationstore, string fileName)
        {
            string content = node.ToString();
            content = new Regex(@"(\/\/\/)").Replace(content, "");
            Comments.Add(new Comment(content,
                node.GetLocation().GetLineSpan().EndLinePosition.Line,
                commentLocationstore)
            {
                LineNumber = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                Type = node.Kind().ToString(),
                FileName = fileName,
            });
        }
    }
}
