using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.CodeAnalysis;

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

        public override void VisitMethodDeclaration(MethodDeclarationSyntax node)
        {
            Console.WriteLine($"{(++i).ToString()}. Name: {node.Identifier} Location: {node.GetLocation().GetLineSpan().StartLinePosition.Line + 1}");
            base.VisitMethodDeclaration(node);
        }
    }
}
