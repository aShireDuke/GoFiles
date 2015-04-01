// Created 20150401 By Andrea Dukeshire
// Event handler code for the ribbon.
// Good resource http://www.codeproject.com/Articles/192724/Justin-s-VSTO-Knowledge-Base-First-VSTO-Applica

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace GoWordAddIn
{
    public partial class RibbonGo
    {
        private void RibbonGo_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Hello World!");
        }
    }
}
