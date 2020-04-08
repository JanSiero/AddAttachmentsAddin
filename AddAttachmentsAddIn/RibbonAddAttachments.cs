using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace AddAttachmentsAddin
{
    public partial class RibbonAddAttachments
    {
        private void RibbonAddAttachments_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void buttonAddAttachments_Click(object sender, RibbonControlEventArgs e)
        {
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                Outlook.MAPIFolder outBox = (Outlook.MAPIFolder)(Globals.ThisAddIn.Application.
                    ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox));
                Outlook.Items items = (Outlook.Items)(outBox.Items);

                if (items.Count == 0) {
                    MessageBox.Show("There are no messages in OUTBOX.", "Add Attachments");
                }
                else {
                    openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    openFileDialog.Filter = "All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 1;
                    openFileDialog.RestoreDirectory = true;
                    openFileDialog.Multiselect = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK) {
                        try
                        {
                            foreach (Outlook.MailItem eMail in items)
                            {
                                foreach (String file in openFileDialog.FileNames)
                                    eMail.Attachments.Add(file, Outlook.OlAttachmentType.olByValue);

                                eMail.Save();
                            }

                            MessageBox.Show("Files were attached.", "Add Attachments");
                        }
                        catch (UnauthorizedAccessException uae)
                        {
                            MessageBox.Show("An error: \n" + uae.Message + "\n" + uae.Source + "\n" +uae.HResult, 
                                "Add Attachments",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);
                        }

                    }
                }
            }
        }
    }
}
