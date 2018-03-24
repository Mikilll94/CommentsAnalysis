using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System;
using System.IO;
using System.Linq;

namespace RoslynPlay
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileContent = File.ReadAllText("GitRef.cs");
            var tree = CSharpSyntaxTree.ParseText(fileContent);
            var root = tree.GetRoot();
            var walker = new Walker();
            walker.Visit(root);
            Console.ReadKey();
        }
    }

    public class Walker : CSharpSyntaxWalker
    {
        int i = 0;

        //public override void VisitMethodDeclaration(MethodDeclarationSyntax node)
        //{
        //    Console.WriteLine($"{(++i).ToString()}.\nName: {node.Identifier}" +
        //        $"\nLocation: {node.GetLocation().GetLineSpan().StartLinePosition.Line + 1}");
        //    base.VisitMethodDeclaration(node);
        //}
        public Walker() : base(SyntaxWalkerDepth.StructuredTrivia)
        {

        }

        public override void VisitTrivia(SyntaxTrivia trivia)
        {
            if (trivia.Kind() == SyntaxKind.SingleLineCommentTrivia ||
                trivia.Kind() == SyntaxKind.MultiLineCommentTrivia)
            {
                Console.WriteLine($"{++i}. {trivia}\nKind: {trivia.Kind()}" +
                    $"\nLocation: {trivia.GetLocation().GetLineSpan().StartLinePosition.Line + 1}");
            }
            base.VisitTrivia(trivia);
        }

        public override void Visit(SyntaxNode node)
        {
            if (node.Kind() == SyntaxKind.SingleLineDocumentationCommentTrivia)
            {
                Console.WriteLine($"{++i} {node.ToString()} Kind: {node.Kind()}" +
                    $"\nLocation: {node.GetLocation().GetLineSpan().StartLinePosition.Line + 1}");
            }
            base.Visit(node);
        }
    }
}
