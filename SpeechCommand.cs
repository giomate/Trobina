using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Speech.Recognition;
using System.Globalization;
using System.Threading;
using System.Diagnostics;

namespace ExcelAddInTest
{
    public class SpeechCommand
    {
        SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine(new CultureInfo("en-US"));
        public bool completed { get; set; }
        public  int command { get; set; }
        public SpeechCommand listener;
        private int counter = 1;
        public SpeechCommand()
        {
            listener = this;
            recognizer.SpeechRecognized +=   new EventHandler<SpeechRecognizedEventArgs>( SpeechRecognizedHandler);
            recognizer.EmulateRecognizeCompleted +=  new EventHandler<EmulateRecognizeCompletedEventArgs>(EmulateRecognizeCompletedHandler);
            recognizer.SpeechDetected += new EventHandler<SpeechDetectedEventArgs>(SpeechDetectedHandler);
            completed = false;
            recognizer.InitialSilenceTimeout = TimeSpan.FromSeconds(3);
            recognizer.BabbleTimeout = TimeSpan.FromSeconds(2);
            recognizer.EndSilenceTimeout = TimeSpan.FromSeconds(1);
            recognizer.EndSilenceTimeoutAmbiguous = TimeSpan.FromSeconds(1.5);
            LoadGrammarRecognizer();

        }
        

         private void LoadGrammarRecognizer()
        {
            Choices nextChoices = new Choices(new string[] { "next", "nechts"});
            Grammar nextGrammar =    new Grammar(nextChoices);
            Choices upChoices = new Choices(new string[] { "upper cell" });
            Grammar upGrammar = new Grammar(upChoices);
            Choices currentChoices = new Choices(new string[] { "current cell" });
            Grammar currentGrammar = new Grammar(currentChoices);
            nextGrammar.Name = "Next";
            recognizer.LoadGrammar(nextGrammar);
            recognizer.LoadGrammar(upGrammar);
            recognizer.LoadGrammar(currentGrammar);
            recognizer.SetInputToDefaultAudioDevice();
            command = 0;
        }
        public int StartRecognition()
        {
            bool confi = false;
            command = 0;
            completed = false;
            
            //  RecognitionResult result = recognizer.EmulateRecognize("next");
            recognizer.RecognizeAsync(RecognizeMode.Multiple);
           
         //  recognizer.Dispose();
            return command;

        }

        public void CancelRecognition()
        {
            recognizer.RecognizeAsyncStop();
            recognizer.RecognizeAsyncCancel();
        }
        
        private  void SpeechRecognizedHandler(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result != null)
            {
                Debug.WriteLine("Recognition result = "+e.Result.Text,
                  e.Result.Text ?? "<no text>");
            }
            else
            {
                Debug.WriteLine("No recognition result");
            }
            if (e.Result.Confidence>0.9)
            {
                
                if ((e.Result.Text== "next")| (e.Result.Text == "nechts"))
                {
                    command = 2;
                }
                else if ((e.Result.Text == "upper cell") )
                {
                    command = 8;
                }
                else if ((e.Result.Text == "current cell") )
                {
                    command = 5;
                }
                completed=true;
                SayNumber();
                
            }
        }

        public  void SayNumber()
        {
            string text;
            TextSynthesizer speaker = new TextSynthesizer();
           
            if (command > 0)
            {
                switch (command)
                {
                    case 2:
                        
                        text = Globals.ThisAddIn.GetNextNumber(command);
                        speaker.SpeakWord(text);

                        break;
                    case 8:
                        counter--;
                        text = Globals.ThisAddIn.GetNextNumber(command);
                        speaker.SpeakWord(text);

                        break;
                    case 5:
                        text = Globals.ThisAddIn.GetCurrentCell();
                        speaker.SpeakWord(text);

                        break;
                    default:
                        break;
                }
                command = 0;

            }
        }

        public static void SpeechDetectedHandler(object sender, SpeechDetectedEventArgs e)
        {
            Debug.WriteLine(" In SpeechDetectedHandler:");
            Debug.WriteLine(" - AudioPosition = {0}", e.AudioPosition);
        }
        public void EmulateRecognizeCompletedHandler(object sender, EmulateRecognizeCompletedEventArgs e)
        {
            if (e.Result == null)
            {
                Debug.WriteLine("No result generated.");
            }

            // Indicate the asynchronous operation is complete.  
            completed = true;
        }
        public  void recognizer_StateChanged(object sender, StateChangedEventArgs e)
        {
            if (e.RecognizerState != RecognizerState.Stopped)
            {
                recognizer.EmulateRecognizeAsync("Start listening");
            }
        }
    }
   
}
