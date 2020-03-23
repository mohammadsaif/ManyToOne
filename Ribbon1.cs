using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace ManyToOne
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnManyToOne_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveDocument.MailMerge.MainDocumentType == Microsoft.Office.Interop.Word.WdMailMergeMainDocType.wdEMail)
            {
                ManyToOneForm manToOneForm = new ManyToOneForm();
                manToOneForm.ShowDialog();
            }
            else
            {
                MessageBox.Show("This Add-in works with Mail Merge documents only", "Information", MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
           

        }
    }
}
