using System.Runtime.InteropServices;
using NLog;
using Rubberduck.UI.RegexAssistant;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command
{

    public struct RegexWithOption
    {
        public string Pattern;
        public bool? IgnoreCaseFlag;
        public bool? GlobalFlag;
    }

    /// <summary>
    /// A command that displays the RegexAssistantDialog
    /// </summary>
    [ComVisible(false)]
    public class RegexAssistantWithSelectionCommand : CommandBase
    {
        private static readonly string kRegExpTypeName = "RegExp";

        protected readonly IVBE Vbe;
        private readonly RubberduckParserState _state;

        public RegexAssistantWithSelectionCommand(IVBE vbe, RubberduckParserState state) : base(LogManager.GetCurrentClassLogger())
        {
            Vbe = vbe;
            _state = state;
        }

        protected override void OnExecute(object parameter)
        {
            //get code in selection
            var pane = Vbe.ActiveCodePane;
            var module = pane.CodeModule;
            var qualifiedSelection = pane.GetQualifiedSelection();
            if (!qualifiedSelection.HasValue || module.IsWrappingNullReference)
            {
                return;
            }
            
            //search RegExp variable declaration under the cursor
            var selectedDec = _state.FindSelectedDeclaration(pane);
            RegexWithOption regex = new RegexWithOption();

            if (selectedDec.AsTypeName == kRegExpTypeName && selectedDec.DeclarationType == DeclarationType.Variable)
            {
                var rawPatternText = GetRegExpPropertyAssigningText(qualifiedSelection, selectedDec, "Pattern");
                regex.Pattern = ToStringFromParsedText(rawPatternText);

                var rawGlobalText = GetRegExpPropertyAssigningText(qualifiedSelection, selectedDec, "Global");
                regex.GlobalFlag = ToBoolFromParsedText(rawGlobalText);

                var rawIgnoreCaseText = GetRegExpPropertyAssigningText(qualifiedSelection, selectedDec, "IgnoreCase");
                regex.IgnoreCaseFlag = ToBoolFromParsedText(rawIgnoreCaseText);
            }

            using (var window = new RegexAssistantDialog(regex))
            {
                window.ShowDialog();
            }
        }

        private string ToStringFromParsedText(string parsedText)
        {
            return
                parsedText.Length <= 1 ? null : parsedText.Substring(1, parsedText.Length - 2);
        }

        private bool? ToBoolFromParsedText(string parsedText)
        {
            if (!Boolean.TryParse(parsedText, out bool result))
            {
                return null;
            }

            return result;
        }

        private string GetRegExpPropertyAssigningText(QualifiedSelection? qualifiedSelection, Declaration selectedDec, string propertyName)
        {
            //find declaration of property let expression of regex object
            var propertyDec =
                _state.DeclarationFinder.MatchName(propertyName)
                    .Where(dec =>
                        dec.DeclarationType == DeclarationType.PropertyLet && dec.ParentDeclaration.AsTypeName == "RegExp")
                    .FirstOrDefault();

            if (propertyDec == null)
            {
                return null;
            }

            var propertyRefs =
                propertyDec.References
                    .Where(propRef =>
                        propRef.QualifiedModuleName == qualifiedSelection.Value.QualifiedName &&
                        selectedDec.References.Any(varRef =>
                            IsSameContext(propRef, varRef)
                        )
                    );

            var propertyRef = propertyRefs.FirstOrDefault();
            
            if (propertyRef == null)
            {
                return null;
            }

            var patternText =
                ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(propertyRef.Context)
                    ?.expression()
                    ?.GetText();

            return patternText;
        }

        private bool IsSameContext(IdentifierReference ref1, IdentifierReference ref2)
        {
            var context1 =
                ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(ref1.Context);
            var context2 =
                ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(ref2.Context);

            return context1 == context2;
        }
    }
}
