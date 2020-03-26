﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.QuickFixes.Logistics
{
    internal class QuickFixProvider : IQuickFixProvider
    {
        private readonly IRewritingManager _rewritingManager;
        private readonly IQuickFixFailureNotifier _failureNotifier;
        private readonly Dictionary<Type, HashSet<IQuickFix>> _quickFixes = new Dictionary<Type, HashSet<IQuickFix>>();

        public QuickFixProvider(IRewritingManager rewritingManager, IQuickFixFailureNotifier failureNotifier, IEnumerable<IQuickFix> quickFixes)
        {
            _rewritingManager = rewritingManager;
            _failureNotifier = failureNotifier;
            foreach (var quickFix in quickFixes)
            {
                foreach (var supportedInspection in quickFix.SupportedInspections)
                {
                    if (_quickFixes.ContainsKey(supportedInspection))
                    {
                        _quickFixes[supportedInspection].Add(quickFix);
                    }
                    else
                    {
                        _quickFixes.Add(supportedInspection, new HashSet<IQuickFix> {quickFix});
                    }
                }
            }
        }

        public IEnumerable<IQuickFix> QuickFixes(Type inspectionType)
        {
            if (!_quickFixes.ContainsKey(inspectionType))
            {
                return Enumerable.Empty<IQuickFix>();
            }

            return _quickFixes[inspectionType];
        }

        public IEnumerable<IQuickFix> QuickFixes(IInspectionResult result)
        {
            return QuickFixes(result.Inspection.GetType())
                .Where(fix => !result.DisabledQuickFixes.Contains(fix.GetType().Name))
                .OrderBy(fix => fix.SupportedInspections.Count); // most specific fixes first; keeps "ignore once" last
        }

        private bool CanFix(IQuickFix fix, IInspectionResult result)
        {
            return QuickFixes(result).Contains(fix);
        }

        public void Fix(IQuickFix fix, IInspectionResult result)
        {
            if (!CanFix(fix, result))
            {
                return;
            }

            var rewriteSession = RewriteSession(fix.TargetCodeKind);
            try
            {
                fix.Fix(result, rewriteSession);
            }
            catch (RewriteFailedException)
            {
                _failureNotifier.NotifyQuickFixExecutionFailure(rewriteSession.Status);
            }
            Apply(rewriteSession);
        }

        private void Apply(IExecutableRewriteSession rewriteSession)
        {
            if (!rewriteSession.TryRewrite())
            {
                _failureNotifier.NotifyQuickFixExecutionFailure(rewriteSession.Status);
            }
        }

        private IExecutableRewriteSession RewriteSession(CodeKind targetCodeKind)
        {
            switch (targetCodeKind)
            {
                case CodeKind.CodePaneCode:
                    return _rewritingManager.CheckOutCodePaneSession();
                case CodeKind.AttributesCode:
                    return _rewritingManager.CheckOutAttributesSession();
                default:
                    throw new NotSupportedException(nameof(targetCodeKind));
            }
        }

        public void FixInProcedure(IQuickFix fix, QualifiedMemberName? qualifiedMember, Type inspectionType, IEnumerable<IInspectionResult> results)
        {
            Debug.Assert(qualifiedMember.HasValue, "Null qualified member.");

            var filteredResults = results.Where(result => result.Inspection.GetType() == inspectionType && result.QualifiedMemberName == qualifiedMember).ToList();

            if (!filteredResults.Any())
            {
                return;
            }

            var rewriteSession = RewriteSession(fix.TargetCodeKind);
            foreach (var result in filteredResults)
            {
                if (!CanFix(fix, result))
                {
                    continue;
                }

                fix.Fix(result, rewriteSession);
            }
            Apply(rewriteSession);
        }

        public void FixInModule(IQuickFix fix, QualifiedSelection selection, Type inspectionType, IEnumerable<IInspectionResult> results)
        {
            var filteredResults = results.Where(result => result.Inspection.GetType() == inspectionType && result.QualifiedSelection.QualifiedName == selection.QualifiedName).ToList();

            if (!filteredResults.Any())
            {
                return;
            }

            var rewriteSession = RewriteSession(fix.TargetCodeKind);
            foreach (var result in filteredResults)
            {
                if (!CanFix(fix, result))
                {
                    continue;
                }

                fix.Fix(result, rewriteSession);
            }
            Apply(rewriteSession);
        }

        public void FixInProject(IQuickFix fix, QualifiedSelection selection, Type inspectionType, IEnumerable<IInspectionResult> results)
        {
            var filteredResults = results.Where(result => result.Inspection.GetType() == inspectionType && result.QualifiedSelection.QualifiedName.ProjectId == selection.QualifiedName.ProjectId).ToList();

            if (!filteredResults.Any())
            {
                return;
            }

            var rewriteSession = RewriteSession(fix.TargetCodeKind);
            foreach (var result in filteredResults)
            {
                if (!CanFix(fix, result))
                {
                    continue;
                }

                fix.Fix(result, rewriteSession);
            }
            Apply(rewriteSession);
        }

        public void FixAll(IQuickFix fix, Type inspectionType, IEnumerable<IInspectionResult> results)
        {
            var filteredResults = results.Where(result => result.Inspection.GetType() == inspectionType).ToArray();

            if (!filteredResults.Any())
            {
                return;
            }

            var rewriteSession = RewriteSession(fix.TargetCodeKind);
            foreach (var result in filteredResults)
            {
                if (!CanFix(fix, result))
                {
                    continue;
                }

                fix.Fix(result, rewriteSession);
            }
            Apply(rewriteSession);
        }

        public bool HasQuickFixes(IInspectionResult inspectionResult)
        {
            return _quickFixes.ContainsKey(inspectionResult.Inspection.GetType()) && _quickFixes[inspectionResult.Inspection.GetType()].Any();
        }
    }
}