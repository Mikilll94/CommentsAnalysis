using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System.Collections.Generic;

namespace CommentsAnalysis
{
    public class CommentStore
    {
        public List<Comment> Comments { get; } = new List<Comment>();

        public void AddCommentTrivia(SyntaxTrivia trivia,
            LocationStore commentLocationstore, string fileName)
        {
            if (trivia.Kind() == SyntaxKind.SingleLineCommentTrivia)
            {
                Comments.Add(new SingleLineComment(trivia.ToString(),
                    trivia.GetLocation().GetLineSpan().EndLinePosition.Line + 1, commentLocationstore)
                {
                    FileName = fileName,
                });
            }
            else if (trivia.Kind() == SyntaxKind.MultiLineCommentTrivia)
            {
                Comments.Add(new MultiLineComment(trivia.ToString(),
                    trivia.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                    trivia.GetLocation().GetLineSpan().EndLinePosition.Line + 1, commentLocationstore)
                {
                    FileName = fileName,
                });
            }
        }

        public void AddCommentNode(DocumentationCommentTriviaSyntax node,
            LocationStore commentLocationstore, string fileName)
        {
            Comments.Add(new DocComment(node.ToString(),
                node.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                node.GetLocation().GetLineSpan().EndLinePosition.Line,
                commentLocationstore)
            {
                FileName = fileName,
            });
        }
    }
}
