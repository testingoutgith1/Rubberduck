using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class CodeExplorerCommandMenuItem : CommandMenuItemBase
    {
        public CodeExplorerCommandMenuItem(CommandBase command) 
            : base(command)
        {
        }

        public override string Key => "RubberduckMenu_CodeExplorer";
        public override int DisplayOrder => (int)NavigationMenuItemDisplayOrder.CodeExplorer;
    }
}
