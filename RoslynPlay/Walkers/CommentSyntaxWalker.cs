using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace RoslynPlay
{
    public class CommentSyntaxWalker : CSharpSyntaxWalker
    {
        private string _fileName;
        private CommentLocationStore _commentLocationStore;

        public CommentSyntaxWalker(string fileName, CommentLocationStore commentLocationStore)
            : base(SyntaxWalkerDepth.StructuredTrivia)
        {
            _fileName = fileName;
            _commentLocationStore = commentLocationStore;
        }

        public override void VisitTrivia(SyntaxTrivia trivia)
        {
            if (trivia.Kind() == SyntaxKind.SingleLineCommentTrivia
                || trivia.Kind() == SyntaxKind.MultiLineCommentTrivia)
            {
                CommentStore.AddCommentTrivia(trivia, _commentLocationStore, _fileName);
            }
            base.VisitTrivia(trivia);
        }

        public override void VisitDocumentationCommentTrivia(DocumentationCommentTriviaSyntax node)
        {
            CommentStore.AddCommentNode(node, _commentLocationStore, _fileName);
            base.VisitDocumentationCommentTrivia(node);
        }
    }
}
