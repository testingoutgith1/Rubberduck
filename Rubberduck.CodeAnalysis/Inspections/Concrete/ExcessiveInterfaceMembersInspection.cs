using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete 
{
    /// <summary>
    /// Identifies class modules that define an interface with an excessive number of public members and reminds users about Interface Segregation Principle.
    /// </summary>
    /// <why>
    /// Interfaces should not be designed to continually grow new members; we should be keeping them small, specific, and specialized.
    /// </why>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// '@Interface
    ///
    /// Public Sub DoSomething()
    /// 
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// '@Interface
    ///
    /// Public Sub DoSomething1()
    /// 
    /// End Sub
    /// 
    /// Public Sub DoSomething2()
    /// 
    /// End Sub
    /// 
    /// '...
    /// 
    /// Public Sub DoSomethingNGreaterThanMaxPublicMemberCount()
    ///  
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// 

    internal sealed class ExcessiveInterfaceMembersInspection : DeclarationInspectionBase 
    {
        private static int ExcessiveMemberCount = 10; //todo: make setting rather than constant

        public ExcessiveInterfaceMembersInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.ClassModule) 
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder) => IsInterfaceDeclaration(declaration) ? HasExcessiveMembers((ClassModuleDeclaration)declaration) : false;

        private static bool IsInterfaceDeclaration(Declaration declaration) => declaration is ClassModuleDeclaration classModule ? classModule.IsInterface : false;

        private static bool HasExcessiveMembers(ClassModuleDeclaration declaration)
        {
            var pub = declaration.Members.Where(member => { int acc = (int)member.Accessibility; return acc >= 3 && acc <= 5; }); //get rid of non-public members
            pub = pub.Where(member => !(member.DeclarationType == DeclarationType.Event)); //get rid of public member types that are not part of VBA interface
            pub = pub.Where(member => member.DeclarationType == DeclarationType.PropertyGet ? NoMatchingSetter(pub, member) : true); //get rid of PropertyGet statements for which there is a corresponding Let/Set because these constitute one unique point of access on the interface
            return pub.Count() > ExcessiveMemberCount;
        }

        private static bool NoMatchingSetter(IEnumerable<Declaration> pub, Declaration property) 
        {
            var setter = pub.Where(member => member.IdentifierName == property.IdentifierName);
            return setter.Count() == 0; //if setter count is 1 then there is a matching setter so this returns false
        }

        protected override string ResultDescription(Declaration declaration) //note: I don't have a good sense of what the result description is supposed to be
        { 
            var qualifiedName = declaration.QualifiedModuleName.ToString();
            var declarationType = Resources.RubberduckUI.ResourceManager
                .GetString("DeclarationType_" + declaration.DeclarationType)
                .Capitalize();
            var identifierName = declaration.IdentifierName;

            return string.Format(
                InspectionResults.ExcessiveInterfaceMembersInspection,
                qualifiedName,
                declarationType,
                identifierName);

        }
    }
}