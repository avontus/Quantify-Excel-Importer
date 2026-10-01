using System.Security.Principal;
using System.Windows.Forms;
using ApplicationContext = Avontus.Core.ApplicationContext;

namespace Avontus.Quantify.ExcelImporter
{
	///<summary>
	///		Really not sure why a SerializationException happens whenever MessageBoxHelper.Show is invoked and the AvontusPrincipal is set on the curren thread, but it does.  This class gets us around that problem.
	///</summary>
	public class MessageBoxHelper
	{
		public static void Show(string message)
		{
			Show(message, "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.None);
		}

		public static DialogResult Show(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
		{
			var oldPrincipal = ApplicationContext.User;
			try
			{
				ApplicationContext.User = new GenericPrincipal(new GenericIdentity("unknown"), new string[0]);
				return MessageBox.Show(message, title, buttons, icon);
			}
			finally
			{
				ApplicationContext.User = oldPrincipal;
			}
		}
	}
}