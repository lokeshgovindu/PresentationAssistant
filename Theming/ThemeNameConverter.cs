using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace PresentationAssistant.Theming
{
    /// <summary>
    /// Gives the Theme setting a dropdown in the options property grid, populated from
    /// <see cref="ThemeCatalog"/> so themes added to themes.json show up without a code
    /// change. Not exclusive: a name can still be typed by hand, which matters when a
    /// theme is added to the file while the dialog is open.
    /// </summary>
    internal sealed class ThemeNameConverter : StringConverter
    {
        public override bool GetStandardValuesSupported(ITypeDescriptorContext context) => true;

        public override bool GetStandardValuesExclusive(ITypeDescriptorContext context) => false;

        public override StandardValuesCollection GetStandardValues(ITypeDescriptorContext context)
        {
            // No need to force a reload: the catalog re-reads themes.json by itself when
            // the file changes, and the property grid calls this often.
            List<string> names = ThemeCatalog.Current.Names.ToList();
            return new StandardValuesCollection(names);
        }
    }
}
