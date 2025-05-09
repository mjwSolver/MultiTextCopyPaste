using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;

namespace MultiTextCopyPaste
{
    public partial class MTRibbon
    {
        private List<string> copiedTexts = new List<string>();

        private void MTRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnMultiCopy_Click(object sender, RibbonControlEventArgs e)
        {
            copiedTexts.Clear();
            Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            Selection selection = presentation.Application.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        copiedTexts.Add(shape.TextFrame.TextRange.Text);
                    }
                }

                MessageBox.Show($"Copied {copiedTexts.Count} text boxes.", "MultiCopy", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Please select text boxes before clicking MultiCopy!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnMultiPaste_Click(object sender, RibbonControlEventArgs e)
        {
            Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            Selection selection = presentation.Application.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                if (selection.ShapeRange.Count != copiedTexts.Count)
                {
                    MessageBox.Show("Mismatch: Select exactly the same number of text boxes as copied!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                for (int i = 1, j = 0; i <= copiedTexts.Count; i++, j++)
                {
                    selection.ShapeRange[i].TextFrame.TextRange.Text = copiedTexts[j]; // Apply text in same order as copied
                }

                MessageBox.Show("MultiPaste completed successfully.", "MultiPaste", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Please select text boxes before clicking MultiPaste!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
