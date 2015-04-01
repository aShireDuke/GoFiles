// Created 20150401 By Andrea Dukeshire
// Event handler code for the ribbon.
// See MSDN example Walkthrough: Creating a Custom Tab by Using the Ribbon Designer
// https://msdn.microsoft.com/en-us/library/bb386104.aspx

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

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


            //string path = "C:\\GoFiles\\GoWordAddIn\\GoWordAddIn\\";

            //var dialog = Globals.ThisAddIn.Application.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFileSaveAs];
            //dialog.GetType().InvokeMember("Name", System.Reflection.BindingFlags.SetProperty, null, dialog, new object[] { path });
            //if (dialog.Show() == -1)
            //{
            //  dialog.Execute();
            //}

            // This shows save dialog, but can't custom what showsup!
            // I think need to program in VB if want custom ability, otherwise have to use 
            // win32 saveas see here:http://software-solutions-online.com/2014/03/13/vba-save-file-dialog-filedialogmsofiledialogsaveas/
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("FileSaveAs");

            //Word.Application wapp = Globals.ThisAddIn.Application;
            //Word.Dialog saveDialog = 
            //var wordDialog = Globals.ThisAddIn.Application.FileDialog("msoFileDialogSaveAs");

                    //Dim dlgSaveAs As FileDialog
                    //dlgSaveAs = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
                    //dlgSaveAs.InitialFileName = ActiveDocument.path & "\" & clientName & "-" & goFileName
                    //dlgSaveAs.InitialView = msoFileDialogViewDetails    ' shows details
                    //dlgSaveAs.title = "SaveAsGo Macro" & " - SaveAsGo"
                    //dlgSaveAs.ButtonName = "Save GoFile"
                    //If dlgSaveAs.Show = -1 Then dlgSaveAs.Execute()
                    //Exit Sub

    
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
