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
            LocationStore commentLocationstore, ClassStore classStore, string fileName)
        {
            if (trivia.Kind() == SyntaxKind.SingleLineCommentTrivia)
            {
                Comment comment = new SingleLineComment(trivia.ToString(),
                    trivia.GetLocation().GetLineSpan().EndLinePosition.Line + 1)
                {
                    FileName = fileName,
                };
                comment.Initialize(commentLocationstore, classStore);
                Comments.Add(comment);
            }
            else if (trivia.Kind() == SyntaxKind.MultiLineCommentTrivia)
            {
                Comment comment = new MultiLineComment(trivia.ToString(),
                    trivia.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                    trivia.GetLocation().GetLineSpan().EndLinePosition.Line + 1)
                {
                    FileName = fileName,
                };
                comment.Initialize(commentLocationstore, classStore);
                Comments.Add(comment);
            }
        }

        public void AddCommentNode(DocumentationCommentTriviaSyntax node,
            LocationStore commentLocationstore, ClassStore classStore, string fileName)
        {
            Comment comment = new DocComment(node.ToString(),
                node.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                node.GetLocation().GetLineSpan().EndLinePosition.Line)
            {
                FileName = fileName,
            };
            comment.Initialize(commentLocationstore, classStore);
            Comments.Add(comment);
        }
    }
}
