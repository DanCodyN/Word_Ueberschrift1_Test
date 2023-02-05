using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Word_Ueberschrift1_Test
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Ueberschrift1_Click(object sender, RibbonControlEventArgs e)
        {
            Application wordApp = Globals.ThisAddIn.Application;
            Template buildingBlock = wordApp.Templates["Ueberschrift1"];
            buildingBlock.BuildingBlockEntries.Item(1).Insert(wordApp.Selection.Range, Type.Missing);
        }
    }
}
