using System.IO;
using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Rewriter
{
    /// <summary>
    /// A module rewriter that works off the MemberAttributes token stream and exports, overwrites and re-imports its module on <see cref="Rewrite"/>.
    /// This rewriter works off a token stream obtained from the AttributeParser, well before the code pane parse tree is acquired.
    /// </summary>
    /// <remarks>
    /// <ul>
    /// <li>DO NOT use this rewriter with any pending (not yet re-parsed) changes (e.g. refactorings, quick-fixes), or these changes will be lost.</li>
    /// <li>DO NOT use this rewriter to change any token that the VBE renders, or line number positions will be off.</li>
    /// <li>DO use this rewriter to add/remove hidden <c>Attribute</c> instructions to/from a module.</li>
    /// </ul>
    /// </remarks>
    public class AttributesRewriter : ModuleRewriterBase
    {
        private readonly ISourceCodeHandler _sourceCodeHandler;

        public AttributesRewriter(QualifiedModuleName module, ITokenStream tokenStream, IProjectsProvider projectsProvider, ISourceCodeHandler sourceCodeHandler)
            : base(module, tokenStream, projectsProvider)
        {
            _sourceCodeHandler = sourceCodeHandler;
        }

        //TODO Make this actually work in the presence of attributes.
        //Currently, this will always return true since there are always module level attributes.
        public override bool IsDirty
        {
            get
            {
                using (var codeModule = CodeModule())
                {
                    return codeModule == null || codeModule.Content() != Rewriter.GetText();
                }
            }
        }

        public override void Rewrite()
        {
            if (!IsDirty)
            {
                return;
            }
            
            if (Module.ComponentType == ComponentType.Document)
            {
                // can't re-import a document module
                return;
            }

            var component = ProjectsProvider.Component(Module);
            var file = _sourceCodeHandler.Export(component);

            var content = Rewriter.GetText();
            File.WriteAllText(file, content);
            
            _sourceCodeHandler.Import(component, file);
        }
    }
}