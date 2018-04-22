using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
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

            string className = node.Identifier.ToString();
            var visitedClass = new Class
            {
                FileName = _fileName,
                Name = className,
                SmellsCount = SmellyClasses.All.Count(a => a == className),
                AbstractionSmellsCount = SmellyClasses.Abstraction.Count(a => a == className),
                EncapsulationSmellsCount = SmellyClasses.Encapsulation.Count(a => a == className),
                ModularizationSmellsCount = SmellyClasses.Modularization.Count(a => a == className),
                HierarchySmellsCount = SmellyClasses.Hierarchy.Count(a => a == className),
            };

            _locationStore.AddLocationRelativeToClass(startLine - 1, LocationRelativeToClass.ClassDescription, visitedClass);
            _locationStore.AddLocationRelativeToClass(startLine, LocationRelativeToClass.ClassStart, visitedClass);
            _locationStore.AddLocationRelativeToClass(startLine + 1, endLine - 1, LocationRelativeToClass.ClassInner, visitedClass);
            _locationStore.AddLocationRelativeToClass(endLine, LocationRelativeToClass.ClassEnd, visitedClass);

            _classStore.Classes.Add(visitedClass);

            base.VisitClassDeclaration(node);
        }
    }
}
