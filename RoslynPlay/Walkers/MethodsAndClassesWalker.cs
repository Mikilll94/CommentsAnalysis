using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace RoslynPlay
{
    class MethodsAndClassesWalker : CSharpSyntaxWalker
    {
        private string _fileName;
        private LocationStore _locationStore;

        public MethodsAndClassesWalker(string fileName, LocationStore locationStore)
        {
            _fileName = fileName;
            _locationStore = locationStore;
        }

        public override void VisitMethodDeclaration(MethodDeclarationSyntax node)
        {
            int startLine = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
            int endLine = node.GetLocation().GetLineSpan().EndLinePosition.Line + 1;
            string methodName = node.Identifier.ToString();

            _locationStore.AddMethodLocation(startLine - 1, "method_description", methodName);
            _locationStore.AddMethodLocation(startLine, "method_start", methodName);
            _locationStore.AddMethodLocation(startLine + 1, endLine - 1, "method_inner", methodName);
            _locationStore.AddMethodLocation(endLine, "method_end", methodName);
            base.VisitMethodDeclaration(node);
        }

        public override void VisitConstructorDeclaration(ConstructorDeclarationSyntax node)
        {
            int startLine = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
            int endLine = node.GetLocation().GetLineSpan().EndLinePosition.Line + 1;
            string methodName = node.Identifier.ToString();

            _locationStore.AddMethodLocation(startLine - 1, "method_description", methodName);
            _locationStore.AddMethodLocation(startLine, "method_start", methodName);
            _locationStore.AddMethodLocation(startLine + 1, endLine - 1, "method_inner", methodName);
            _locationStore.AddMethodLocation(endLine, "method_end", methodName);
            base.VisitConstructorDeclaration(node);
        }

        public override void VisitClassDeclaration(ClassDeclarationSyntax node)
        {
            int startLine = node.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
            int endLine = node.GetLocation().GetLineSpan().EndLinePosition.Line + 1;
            string className = node.Identifier.ToString();

            _locationStore.AddClassLocation(startLine - 1, "class_description", className);
            _locationStore.AddClassLocation(startLine, "class_start", className);
            _locationStore.AddClassLocation(startLine + 1, endLine - 1, "class_inner", className);
            _locationStore.AddClassLocation(endLine, "class_end", className);
            base.VisitClassDeclaration(node);
        }
    }
}
