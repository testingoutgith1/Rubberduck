﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.Command;
using Rubberduck.UI.Refactorings;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
// ReSharper disable CanBeReplacedWithTryCastAndCheckForNull

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerViewModel : ViewModelBase
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly NewUnitTestModuleCommand _newUnitTestModuleCommand;
        private readonly Indenter _indenter;
        private readonly ICodePaneWrapperFactory _wrapperFactory;
        private readonly FindAllReferencesCommand _findAllReferences;
        private readonly FindAllImplementationsCommand _findAllImplementations;

        public CodeExplorerViewModel(VBE vbe,
            RubberduckParserState state,
            INavigateCommand navigateCommand,
            NewUnitTestModuleCommand newUnitTestModuleCommand,
            Indenter indenter,
            ICodePaneWrapperFactory wrapperFactory,
            FindAllReferencesCommand findAllReferences,
            FindAllImplementationsCommand findAllImplementations)
        {
            _vbe = vbe;
            _state = state;
            _newUnitTestModuleCommand = newUnitTestModuleCommand;
            _indenter = indenter;
            _wrapperFactory = wrapperFactory;
            _findAllReferences = findAllReferences;
            _findAllImplementations = findAllImplementations;
            _state.StateChanged += ParserState_StateChanged;
            _state.ModuleStateChanged += ParserState_ModuleStateChanged;

            _navigateCommand = navigateCommand;
            _contextMenuNavigateCommand = new DelegateCommand(ExecuteContextMenuNavigateCommand, CanExecuteContextMenuNavigateCommand);
            _refreshCommand = new DelegateCommand(ExecuteRefreshCommand, _ => CanRefresh);
            _addTestModuleCommand = new DelegateCommand(ExecuteAddTestModuleCommand);
            _addStdModuleCommand = new DelegateCommand(ExecuteAddStdModuleCommand, CanAddModule);
            _addClsModuleCommand = new DelegateCommand(ExecuteAddClsModuleCommand, CanAddModule);
            _addFormCommand = new DelegateCommand(ExecuteAddFormCommand, CanAddModule);
            _indenterCommand = new DelegateCommand(ExecuteIndenterCommand, _ => CanExecuteIndenterCommand);
            _renameCommand = new DelegateCommand(ExecuteRenameCommand, _ => CanExecuteRenameCommand);
            _findAllReferencesCommand = new DelegateCommand(ExecuteFindAllReferencesCommand, _ => CanExecuteFindAllReferencesCommand);
            _findAllImplementationsCommand = new DelegateCommand(ExecuteFindAllImplementationsCommand, _ => CanExecuteFindAllImplementationsCommand);
        }

        private readonly ICommand _refreshCommand;
        public ICommand RefreshCommand { get { return _refreshCommand; } }

        private readonly ICommand _addTestModuleCommand;
        public ICommand AddTestModuleCommand { get { return _addTestModuleCommand; } }

        private readonly ICommand _addStdModuleCommand;
        public ICommand AddStdModuleCommand { get { return _addStdModuleCommand; } }

        private readonly ICommand _addClsModuleCommand;
        public ICommand AddClsModuleCommand { get { return _addClsModuleCommand; } }

        private readonly ICommand _addFormCommand;
        public ICommand AddFormCommand { get { return _addFormCommand; } }

        private readonly ICommand _indenterCommand;
        public ICommand IndenterCommand { get { return _indenterCommand; } }

        private readonly ICommand _renameCommand;
        public ICommand RenameCommand { get { return _renameCommand; } }

        private readonly ICommand _findAllReferencesCommand;
        public ICommand FindAllReferencesCommand { get { return _findAllReferencesCommand; } }

        private readonly ICommand _findAllImplementationsCommand;
        public ICommand FindAllImplementationsCommand { get { return _findAllImplementationsCommand; } }

        private readonly INavigateCommand _navigateCommand;
        public ICommand NavigateCommand { get { return _navigateCommand; } }

        private readonly ICommand _contextMenuNavigateCommand;
        public ICommand ContextMenuNavigateCommand { get { return _contextMenuNavigateCommand; } }

        public string Description
        {
            get
            {
                if (SelectedItem is CodeExplorerProjectViewModel)
                {
                    return ((CodeExplorerProjectViewModel)SelectedItem).Declaration.DescriptionString;
                }

                if (SelectedItem is CodeExplorerComponentViewModel)
                {
                    return ((CodeExplorerComponentViewModel)SelectedItem).Declaration.DescriptionString;
                }

                if (SelectedItem is CodeExplorerMemberViewModel)
                {
                    return ((CodeExplorerMemberViewModel)SelectedItem).Declaration.DescriptionString;
                }

                if (SelectedItem is CodeExplorerCustomFolderViewModel)
                {
                    return ((CodeExplorerCustomFolderViewModel)SelectedItem).FolderAttribute;
                }

                return string.Empty;
            }
        }

        private CodeExplorerItemViewModel _selectedItem;
        public CodeExplorerItemViewModel SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value; 
                OnPropertyChanged();
                // ReSharper disable ExplicitCallerInfoArgument
                OnPropertyChanged("CanExecuteIndenterCommand");
                OnPropertyChanged("CanExecuteRenameCommand");
                OnPropertyChanged("CanExecuteFindAllReferencesCommand");
                OnPropertyChanged("Description");
                // ReSharper restore ExplicitCallerInfoArgument
            }
        }

        private bool _isBusy;

        public bool IsBusy
        {
            get { return _isBusy; }
            set
            {
                _isBusy = value; 
                OnPropertyChanged();
                CanRefresh = !_isBusy;
            }
        }

        private bool _canRefresh = true;
        public bool CanRefresh
        {
            get { return _canRefresh; }
            private set
            {
                _canRefresh = value;
                OnPropertyChanged();
            }
        }

        private bool CanAddModule(object param)
        {
            return _vbe.ActiveVBProject != null;
        }

        private bool CanExecuteContextMenuNavigateCommand(object param)
        {
            return SelectedItem != null && SelectedItem.QualifiedSelection.HasValue;
        }

        public bool CanExecuteIndenterCommand
        {
            get
            {
                return _state.Status == ParserState.Ready && !(SelectedItem is CodeExplorerCustomFolderViewModel);
            }
        }

        public bool CanExecuteRenameCommand
        {
            get
            {
                return _state.Status == ParserState.Ready && !(SelectedItem is CodeExplorerCustomFolderViewModel);
            }
        }

        public bool CanExecuteFindAllReferencesCommand
        {
            get
            {
                return _state.Status == ParserState.Ready && !(SelectedItem is CodeExplorerCustomFolderViewModel);
            }
        }

        public bool CanExecuteFindAllImplementationsCommand
        {
            get
            {
                return _state.Status == ParserState.Ready &&
                       (SelectedItem is CodeExplorerComponentViewModel ||
                       SelectedItem is CodeExplorerMemberViewModel);
            }
        }

        private ObservableCollection<CodeExplorerItemViewModel> _projects;

        public ObservableCollection<CodeExplorerItemViewModel> Projects
        {
            get { return _projects; }
            set
            {
                _projects = value; 
                OnPropertyChanged();
            }
        }

        private void ParserState_StateChanged(object sender, EventArgs e)
        {
            if (Projects == null)
            {
                Projects = new ObservableCollection<CodeExplorerItemViewModel>();
            }

            IsBusy = _state.Status == ParserState.Parsing;
            if (_state.Status != ParserState.Ready)
            {
                return;
            }

            Debug.WriteLine("Creating Code Explorer model...");
            var userDeclarations = _state.AllUserDeclarations
                .GroupBy(declaration => declaration.Project)
                .Where(grouping => grouping.Key != null)
                .ToList();

            if (userDeclarations.Any(grouping => grouping.All(declaration => declaration.DeclarationType != DeclarationType.Project)))
            {
                return;
            }

            var newProjects = new ObservableCollection<CodeExplorerItemViewModel>(userDeclarations.Select(grouping =>
                new CodeExplorerProjectViewModel(grouping.SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Project), grouping)));

            UpdateNodes(Projects, newProjects);
            Projects = newProjects;
        }

        private void UpdateNodes(IEnumerable<CodeExplorerItemViewModel> oldList,
            IEnumerable<CodeExplorerItemViewModel> newList)
        {
            foreach (var item in newList)
            {
                CodeExplorerItemViewModel oldItem;

                if (item is CodeExplorerCustomFolderViewModel)
                {
                    oldItem = oldList.FirstOrDefault(i => i.Name == item.Name);
                }
                else
                {
                    oldItem = oldList.FirstOrDefault(i =>
                        item.QualifiedSelection != null && i.QualifiedSelection != null &&
                        i.QualifiedSelection.Value.QualifiedName.ProjectId ==
                        item.QualifiedSelection.Value.QualifiedName.ProjectId &&
                        i.QualifiedSelection.Value.QualifiedName.ComponentName ==
                        item.QualifiedSelection.Value.QualifiedName.ComponentName &&
                        i.QualifiedSelection.Value.Selection == item.QualifiedSelection.Value.Selection);
                }

                if (oldItem != null)
                {
                    item.IsExpanded = oldItem.IsExpanded;

                    if (oldItem.Items.Any() && item.Items.Any())
                    {
                        UpdateNodes(oldItem.Items, item.Items);
                    }
                }
            }
        }

        private void ParserState_ModuleStateChanged(object sender, Parsing.ParseProgressEventArgs e)
        {
            // todo: figure out a way to handle error state.
            // the problem is that the _projects collection might not contain our failing module yet.
        }

        private void ExecuteRefreshCommand(object param)
        {
            _state.OnParseRequested(this);
        }

        private void ExecuteAddTestModuleCommand(object param)
        {
            _newUnitTestModuleCommand.NewUnitTestModule();
        }

        private void ExecuteAddStdModuleCommand(object param)
        {
            _vbe.ActiveVBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
        }

        private void ExecuteAddClsModuleCommand(object param)
        {
            _vbe.ActiveVBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
        }

        private void ExecuteAddFormCommand(object param)
        {
            _vbe.ActiveVBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_MSForm);
        }

        private void ExecuteIndenterCommand(object param)
        {
            if (!SelectedItem.QualifiedSelection.HasValue)
            {
                return;
            }

            if (SelectedItem is CodeExplorerProjectViewModel)
            {
                _indenter.Indent(SelectedItem.QualifiedSelection.Value.QualifiedName.Project);
            }

            if (SelectedItem is CodeExplorerComponentViewModel)
            {
                _indenter.Indent(SelectedItem.QualifiedSelection.Value.QualifiedName.Component);
            }

            if (SelectedItem is CodeExplorerMemberViewModel)
            {
                var arg = new NavigateCodeEventArgs(SelectedItem.QualifiedSelection.Value);
                NavigateCommand.Execute(arg);

                _indenter.IndentCurrentProcedure();
            }
        }

        private void ExecuteRenameCommand(object obj)
        {
            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(_vbe, view, _state, new MessageBox(), _wrapperFactory);
                var refactoring = new RenameRefactoring(_vbe, factory, new MessageBox(), _state);

                refactoring.Refactor(GetSelectedDeclaration());
            }
        }

        private void ExecuteFindAllReferencesCommand(object obj)
        {
            _findAllReferences.Execute(GetSelectedDeclaration());
        }

        private void ExecuteFindAllImplementationsCommand(object obj)
        {
            _findAllImplementations.Execute(GetSelectedDeclaration());
        }

        private void ExecuteContextMenuNavigateCommand(object obj)
        {
            // ReSharper disable once PossibleInvalidOperationException
            // CanExecute protects against this
            var arg = new NavigateCodeEventArgs(SelectedItem.QualifiedSelection.Value);

            NavigateCommand.Execute(arg);
        }

        private Declaration GetSelectedDeclaration()
        {
            if (SelectedItem is CodeExplorerProjectViewModel)
            {
                return ((CodeExplorerProjectViewModel) SelectedItem).Declaration;
            }

            if (SelectedItem is CodeExplorerComponentViewModel)
            {
                return ((CodeExplorerComponentViewModel) SelectedItem).Declaration;
            }

            if (SelectedItem is CodeExplorerMemberViewModel)
            {
                return ((CodeExplorerMemberViewModel) SelectedItem).Declaration;
            }

            return null;
        }
    }
}
