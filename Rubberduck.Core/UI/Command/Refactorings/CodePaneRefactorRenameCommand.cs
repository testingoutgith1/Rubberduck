using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class CodePaneRefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _messageBox;

        public CodePaneRefactorRenameCommand(IVBE vbe, RubberduckParserState state, IMessageBox messageBox, IRewritingManager rewritingManager) 
            : base (vbe)
        {
            _state = state;
            _rewritingManager = rewritingManager;
            _messageBox = messageBox;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            Declaration target;
            using (var activePane = Vbe.ActiveCodePane)
            {
                if (activePane == null || activePane.IsWrappingNullReference)
                {
                    return false;
                }
            
                target = _state.FindSelectedDeclaration(activePane);
            }

            return _state.Status == ParserState.Ready 
                && target != null 
                && target.IsUserDefined 
                && !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        protected override void OnExecute(object parameter)
        {
            Declaration target;
            using (var activePane = Vbe.ActiveCodePane)
            {
                if (activePane == null || activePane.IsWrappingNullReference)
                {
                    return;
                }

                if (parameter != null)
                {
                    target = parameter as Declaration;
                }
                else
                {
                    target = _state.FindSelectedDeclaration(activePane);
                }
            }

            if (target == null || !target.IsUserDefined)
            {
                return;
            }

            using (var view = new RenameDialog(new RenameViewModel(_state)))
            {
                var factory = new RenamePresenterFactory(Vbe, view, _state);
                var refactoring = new RenameRefactoring(Vbe, factory, _messageBox, _state, _state.ProjectsProvider, _rewritingManager);

                refactoring.Refactor(target);
            }
        }
    }
}