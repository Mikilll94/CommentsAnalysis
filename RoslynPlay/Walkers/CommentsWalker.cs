using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace RoslynPlay
{
    public class CommentsWalker : CSharpSyntaxWalker
    {
        private string _fileName;
        private LocationStore _commentLocationStore;
        private CommentStore _commentStore;

        public CommentsWalker(string fileName, LocationStore commentLocationStore,
            CommentStore commentStore)
            : base(SyntaxWalkerDepth.StructuredTrivia)
        {
            _fileName = fileName;
            _commentLocationStore = commentLocationStore;
            _commentStore = commentStore;
        }

        public override void VisitTrivia(SyntaxTrivia trivia)
        {
            if (trivia.Kind() == SyntaxKind.SingleLineCommentTrivia
                || trivia.Kind() == SyntaxKind.MultiLineCommentTrivia)
            {
                _commentStore.AddCommentTrivia(trivia, _commentLocationStore, _fileName);
            }
            base.VisitTrivia(trivia);
        }

        public override void VisitDocumentationCommentTrivia(DocumentationCommentTriviaSyntax node)
        {
            _commentStore.AddCommentNode(node, _commentLocationStore, _fileName);
            base.VisitDocumentationCommentTrivia(node);
        }
    }
}
