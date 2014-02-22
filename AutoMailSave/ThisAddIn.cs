using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace AutoMailSave
{
	public partial class ThisAddIn
	{
		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			// Create an event handler for when items are sent
			Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(SaveEmail);
		}


		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
		}

		private void SaveEmail(object Item, ref bool Cancel)
		{
			var message = Item as Outlook.MailItem;

			string fileName = FileName(message.Recipients[1]);
			string savePath = PathName();

			message.SaveAs(savePath + fileName, Outlook.OlSaveAsType.olMSG);
		}

		private string PathName()
		{
			string pathName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\MailSave\" + DateTime.Now.ToString("yyyy-MM-dd");

			Directory.CreateDirectory(pathName);

			return pathName;
		}

		private string FileName(Outlook.Recipient recipient)
		{
			string fileName = @"\" + recipient.Address + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".msg";

			return fileName;
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
				this.Startup += new System.EventHandler(ThisAddIn_Startup);
				this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}
				
		#endregion
	}
}
