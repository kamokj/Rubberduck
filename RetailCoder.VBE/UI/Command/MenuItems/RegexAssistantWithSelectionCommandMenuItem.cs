using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RegexAssistantWithSelectionCommandMenuItem : CommandMenuItemBase
    {
        public RegexAssistantWithSelectionCommandMenuItem(CommandBase command) : base(command)
        {
        }

        public override string Key => "RubberduckMenu_RegexAssistantWithSelection";
        public override bool BeginGroup => true;
        public override int DisplayOrder => (int)CodePaneContextMenuItemDisplayOrder.RegexAssistantWithSelection;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return true;
        }

    }
}
