using CommentsAnalysis.Models;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System;
using System.Linq;

namespace RoslynPlay
{
    class MethodsAndClassesWalker : CSharpSyntaxWalker
    {
        private string _fileName;
        private LocationStore _locationStore;
        private ClassStore _classStore;

        public MethodsAndClassesWalker(string fileName, LocationStore locationStore, ClassStore classStore)
        {
            _fileName = fileName;
            _locationStore = locationStore;
            _classStore = classStore;
        }

        public override void VisitMethodDeclaration(MethodDeclarationSyntax node)
        {
            int startLine = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
            int endLine = node.GetLocation().GetLineSpan().EndLinePosition.Line + 1;
            string methodName = node.Identifier.ToString();

            _locationStore.AddLocationRelativeToMethod(startLine - 1, LocationRelativeToMethod.MethodDescription, methodName);
            _locationStore.AddLocationRelativeToMethod(startLine, LocationRelativeToMethod.MethodStart, methodName);
            _locationStore.AddLocationRelativeToMethod(startLine + 1, endLine - 1, LocationRelativeToMethod.MethodInner, methodName);
            _locationStore.AddLocationRelativeToMethod(endLine, LocationRelativeToMethod.MethodEnd, methodName);
            base.VisitMethodDeclaration(node);
        }

        public override void VisitConstructorDeclaration(ConstructorDeclarationSyntax node)
        {
            int startLine = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
            int endLine = node.GetLocation().GetLineSpan().EndLinePosition.Line + 1;
            string methodName = node.Identifier.ToString();

            _locationStore.AddLocationRelativeToMethod(startLine - 1, LocationRelativeToMethod.MethodDescription, methodName);
            _locationStore.AddLocationRelativeToMethod(startLine, LocationRelativeToMethod.MethodStart, methodName);
            _locationStore.AddLocationRelativeToMethod(startLine + 1, endLine - 1, LocationRelativeToMethod.MethodInner, methodName);
            _locationStore.AddLocationRelativeToMethod(endLine, LocationRelativeToMethod.MethodEnd, methodName);
            base.VisitConstructorDeclaration(node);
        }

        public override void VisitClassDeclaration(ClassDeclarationSyntax node)
        {
            int startLine = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
            int endLine = node.GetLocation().GetLineSpan().EndLinePosition.Line + 1;

            string name = node.Identifier.ToString();
            NamespaceDeclarationSyntax namespaceNode = node.Ancestors().OfType<NamespaceDeclarationSyntax>().FirstOrDefault();
            string @namespace = namespaceNode != null ? namespaceNode.Name.ToString() : "";

            Func<Smell, bool> predicate = s => s.ClassName == name && s.ClassNamespace == @namespace;

            var visitedClass = new Class
            {
                FileName = _fileName,
                Namespace = @namespace,
                Name = name,
                SmellsCount = SmellsStore.Smells.Count(predicate),
                AbstractionSmellsCount = SmellsStore.Smells.Where(s => s.Type == SmellType.Abstraction).Count(predicate),
                EncapsulationSmellsCount = SmellsStore.Smells.Where(s => s.Type == SmellType.Encapsulation).Count(predicate),
                ModularizationSmellsCount = SmellsStore.Smells.Where(s => s.Type == SmellType.Modularization).Count(predicate),
                HierarchySmellsCount = SmellsStore.Smells.Where(s => s.Type == SmellType.Hierarchy).Count(predicate),
            };

            _locationStore.AddLocationRelativeToClass(startLine - 1, LocationRelativeToClass.ClassDescription, visitedClass);
            _locationStore.AddLocationRelativeToClass(startLine, LocationRelativeToClass.ClassStart, visitedClass);
            _locationStore.AddLocationRelativeToClass(startLine + 1, endLine - 1, LocationRelativeToClass.ClassInner, visitedClass);
            _locationStore.AddLocationRelativeToClass(endLine, LocationRelativeToClass.ClassEnd, visitedClass);

            // Handle partial classes
            if (!_classStore.Classes.Exists(c => c.Name == name && c.Namespace == @namespace))
            {
                _classStore.Classes.Add(visitedClass);
            }

            base.VisitClassDeclaration(node);
        }
    }
}
