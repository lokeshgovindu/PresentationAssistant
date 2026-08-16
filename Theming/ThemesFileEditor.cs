using System;
using System.ComponentModel;
using System.Drawing.Design;

namespace PresentationAssistant.Theming
{
    /// <summary>
    /// Puts a "..." button on the Themes File row of the options page. Pressing it opens
    /// themes.json in the editor rather than returning a value, so the property itself is
    /// left untouched.
    /// </summary>
    internal sealed class ThemesFileEditor : UITypeEditor
    {
        public override UITypeEditorEditStyle GetEditStyle(ITypeDescriptorContext context)
        {
            return UITypeEditorEditStyle.Modal;
        }

        public override object EditValue(
            ITypeDescriptorContext context, IServiceProvider provider, object value)
        {
            PresentationAssistantPackage.OpenThemesFile();

            // The path is not user-editable; hand back exactly what we were given.
            return value;
        }
    }
}
