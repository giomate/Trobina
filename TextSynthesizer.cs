using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Speech.Synthesis;
using System.Reflection;

namespace ExcelAddInTest
{
    public class TextSynthesizer
    {
        public  bool isTalking { get; set; }
        public TextSynthesizer textSynthesizer;
        public TextSynthesizer()
        {
            textSynthesizer = this;
        }
        
        public  void  SpeakWord(string t)
        {
            SpeechSynthesizer synth = new SpeechSynthesizer();

            // Configure the audio output.   
            synth.SetOutputToDefaultAudioDevice();

            isTalking = true;
            synth.SpeakCompleted += new EventHandler<SpeakCompletedEventArgs>(synthFinish);
            // Speak a string.  
            synth.Speak(t);
            

            while (isTalking && (synth.State==SynthesizerState.Speaking))
            {
                ;

            }
            synth.Dispose();
        }
        void synthFinish(object sender, SpeakCompletedEventArgs e)
        {
            isTalking = false;
        }


    }
    
}
