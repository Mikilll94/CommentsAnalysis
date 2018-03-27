﻿using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace RoslynPlay
{
    public class CommentSyntaxWalker : CSharpSyntaxWalker
    {
        string visitedFile;

        public CommentSyntaxWalker(string fileName) : base(SyntaxWalkerDepth.StructuredTrivia)
        {
            visitedFile = fileName; 
        }

        public override void VisitTrivia(SyntaxTrivia trivia)
        {
            if (trivia.Kind() == SyntaxKind.SingleLineCommentTrivia)
            {
                string content = trivia.ToString();
                CommentStore.Comments.Add(new Comment(content.Substring(content.IndexOf("//") + 2),
                    trivia.GetLocation().GetLineSpan().StartLinePosition.Line + 1)
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
                CommentStore.Comments.Add(new Comment(content, trivia.GetLocation().GetLineSpan().EndLinePosition.Line + 1)
                {
                    LineNumber = trivia.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                    Type = "multi_line_comment",
                    FileName = visitedFile,
                });
            }
            base.VisitTrivia(trivia);
        }

        public override void VisitDocumentationCommentTrivia(DocumentationCommentTriviaSyntax node)
        {
            if (node.Kind() == SyntaxKind.SingleLineDocumentationCommentTrivia)
            {
                string content = node.ToString();
                content = new Regex(@"(\/\/\/)").Replace(content, "");
                CommentStore.Comments.Add(new Comment(content,
                    node.GetLocation().GetLineSpan().EndLinePosition.Line)
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
                CommentStore.Comments.Add(new Comment(node.ToString(),
                    node.GetLocation().GetLineSpan().EndLinePosition.Line + 1)
                {
                    LineNumber = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                    Type = "multi_line_doc",
                    FileName = visitedFile,
                });
            }
            base.VisitDocumentationCommentTrivia(node);
        }
    }
}
