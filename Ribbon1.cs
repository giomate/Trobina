using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest
{
    public partial class Ribbon1
    {
        private int counter = 1;
        private bool talking = false;
      //  this.ActionsPane.Controls.Add(actions);
       // ActionsPaneControl1 actionsPane1 = new ActionsPaneControl1();
     //   ActionsPaneControl2 actionsPane2 = new ActionsPaneControl2();
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //Globals.ThisWorkbook.ActionsPane.Controls.Add(actionsPane1);
           // Globals.ThisWorkbook.ActionsPane.Controls.Add(actionsPane2);
           // actionsPane1.Hide();
           // actionsPane2.Hide();
           // Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = false;

       //     this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler( this.button1_Click);
          
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            string text = Globals.ThisAddIn.GetReferenceWord();
            TextSynthesizer speaker = new TextSynthesizer();
            speaker.SpeakWord(text);
            
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            talking = !talking;
            string text;
            TextSynthesizer speaker= new TextSynthesizer();
            SpeechCommand listener = new SpeechCommand();
            if (talking)
            {
                button2.Label = "Talking";
            }
            else
            {
                button2.Label = "Silence";
            }
            button2.PerformDynamicLayout();
            if (talking)
            {
                int rec = listener.StartRecognition();
                if (rec > 0)
                {
                    switch (rec)
                    {
                        case 2:
                            counter++;
                            text = Globals.ThisAddIn.GetNextNumber(counter);
                            speaker.SpeakWord(text);

                            break;
                        case 8:
                            counter--;
                            text = Globals.ThisAddIn.GetNextNumber(counter);
                            speaker.SpeakWord(text);

                            break;
                        case 5:
                             text = Globals.ThisAddIn.GetCurrentCell();
                            speaker.SpeakWord(text);

                            break;
                        default:
                            break;
                    }
                    rec = 0;

                }
            }
            else
            {
                listener.CancelRecognition();
            }

            
        }
    }
}
