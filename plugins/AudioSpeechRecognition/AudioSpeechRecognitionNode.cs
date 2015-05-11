// By Lev Panov, 2015 (lev.panov@gmail.com)
// Further work
// 1. Repeated/optional phrases https://msdn.microsoft.com/en-us/library/ms576572(v=vs.110).aspx
// Possible to arrange repetition count pins in a way similar to group (ex9 priority)
// 2. Semantic information
// 3. Free text dictation: https://msdn.microsoft.com/en-us/library/ms576565(v=vs.110).aspx
// 4. Different languages
// 5. Configurable single/multiple recognition modes
// 6. Output phrase index (for complex grammar - slice of indices for each choice?)

#region usings
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Threading;
using System.Threading.Tasks;

using VVVV.PluginInterfaces.V1;
using VVVV.PluginInterfaces.V2;
using VVVV.Utils.Streams;
using VVVV.Utils.VColor;
using VVVV.Utils.VMath;

using VVVV.Core.Logging;

using Microsoft.Speech.Recognition;
#endregion usings

namespace VVVV.Nodes
{
	#region PluginInfo
	[PluginInfo(Name = "SpeechRecognition",
                Category = "Microsoft Speech Platform",
                Version = "Predefined Grammar",
                Help = "Speech recognition for different languages with configurable predefined grammar",
                Tags = "speech",
                Author = "lev")]
	#endregion PluginInfo
	public class AudioSpeechRecognitionNode : IPluginEvaluate, IPartImportsSatisfiedNotification, IDisposable
	{
		#region fields & pins
        //[Input("Choices", IsPinGroup = true)]
        //public IDiffSpread<ISpread<string>> FChoices;
        public Spread<IIOContainer<ISpread<string>>> FChoices = new Spread<IIOContainer<ISpread<string>>>();
        public Spread<IIOContainer<ISpread<int>>> FRepetitionsNumber = new Spread<IIOContainer<ISpread<int>>>();

        [Config("Choices Count", DefaultValue = 1, MinValue = 1)]
        public IDiffSpread<int> FChoicesCount;        

        [Input("Culture Name", IsSingle = true, DefaultString = "en-US", Order = 1)]
        IDiffSpread<string> FCultureName;

        [Input("Enabled", IsSingle = true, DefaultBoolean = true, Order = 0)]
        IDiffSpread<bool> FEnabled;

        [Input("Confidence Threshold", IsSingle = true, DefaultValue = 0.5, Order = 2)]
        IDiffSpread<double> FConfidenceThreshold;

        // TODO
        //[Output("Choices indices", IsPinGroup = true, Order = 4)]
        //ISpread<ISpread<string>> FChoices;

        [Output("Recognition Result", IsSingle = true)]
		public ISpread<String> FRecognitionResult;

        [Output("Confidence", IsSingle = true)]
        public ISpread<double> FConfidence;

        [Output("On Speech Detected", IsSingle = true, IsToggle = true)]
        public ISpread<bool> FOnSpeechDetected;

        [Output("On Recognized", IsSingle = true, IsBang = true)]
        public ISpread<bool> FOnRecognized;

        [Output("Recognizer for culture found", IsSingle = true, IsToggle = true)]
        public ISpread<bool> FRecognizerForCultureFound;

        [Output("Grammar Loaded", IsSingle = true, IsToggle = true)]
        public ISpread<bool> FGrammarLoaded;

		[Import()]
		public ILogger FLogger;
        [Import] ///!!!
        public IIOFactory FIOFactory;

        // Speech recognition engine 
        private SpeechRecognitionEngine recognizer = null;
        // Currently loaded grammar
        private Grammar currentGrammar = null;
        // Support bang behavior for FOnRecognized pin
        private bool onRecognizedBangFrameElapsed = true;
        // True or false whether recognition has been started or not
        private bool recognitionStarted;
        private bool RecognitionStarted
        {
            get { return recognitionStarted; }
            set { recognitionStarted = value; }
        }
        private Object startRecognitionLock = new Object();
        private Object asyncActionMonitor = new Object();
		#endregion fields & pins

        #region pin management
        public void OnImportsSatisfied()
        {
            FChoicesCount.Changed += HandleInputCountChanged;
        }

        private void HandlePinCountChanged<T>(ISpread<int> countSpread, Spread<IIOContainer<T>> pinSpread,
                                              Func<int, IOAttribute> ioAttributeFactory) where T : class
        {
            pinSpread.ResizeAndDispose(
                countSpread[0],
                (i) =>
                {
                    var ioAttribute = ioAttributeFactory(i + 1);
                    return FIOFactory.CreateIOContainer<T>(ioAttribute);
                }
            );
        }

        InputAttribute createInputAttributeWithChangeCheck(int i)
        {
            var ia = new InputAttribute(string.Format("Choices {0}", i));
            ia.CheckIfChanged = false;
            return ia;
        }

        private void HandleInputCountChanged(IDiffSpread<int> sender)
        {
            //HandlePinCountChanged(sender, FChoices, (i) => new InputAttribute(string.Format("Choices {0}", i)));
            HandlePinCountChanged(sender, FChoices, createInputAttributeWithChangeCheck);
            HandlePinCountChanged(sender, FRepetitionsNumber, (i) => new InputAttribute(string.Format("Repetitions Number {0}", i)));
        }
        #endregion

        #region ctor & dispose
        //public AudioSpeechRecognitionNode()
        //{
        //    // Create new engine
        //    //reinitialize();
        //}

        public void Dispose()
        {
            if (recognizer != null)
            {
                ((IDisposable)recognizer).Dispose();
                recognizer = null;
            }
        }
        #endregion

        #region private methods

        void reinitialize(string cultureName)
        {
            // Dispose old recognizer if present
            Dispose();
            // Reset all needed class memebers
            recognitionStarted = false;

            try
            {
                var culture = new System.Globalization.CultureInfo(cultureName);

                try
                {
                    // Create SpeechRecognitionEngine instance
                    recognizer = new SpeechRecognitionEngine(culture);
                }
                catch (ArgumentException)
                {
                    if (FLogger != null)
                    {
                        FLogger.Log(LogType.Error,
                                    "No recognized found for culture \"" + cultureName + "\"");
                    }
                    FRecognizerForCultureFound[0] = false;
                    return;
                }

                FRecognizerForCultureFound[0] = true;                
            }
            catch (System.Globalization.CultureNotFoundException)
            {
                if (FLogger != null)
                {
                    FLogger.Log(LogType.Error,
                                "Culture with name \"" + cultureName + "\" not found");
                }
                FRecognizerForCultureFound[0] = false;
                return;
            }

            // Add a handler for the speech recognized event
            recognizer.SpeechRecognized +=
                new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);
            // Attach other event handlers
            recognizer.RecognizerUpdateReached +=
                new EventHandler<RecognizerUpdateReachedEventArgs>(recognizer_RecognizerUpdateReached);
            recognizer.SpeechRecognitionRejected +=
                new EventHandler<SpeechRecognitionRejectedEventArgs>(recognizer_SpeechRecognitionRejected);
            recognizer.SpeechDetected +=
                new EventHandler<SpeechDetectedEventArgs>(recognizer_SpeechDetected);
            recognizer.SpeechDetected +=
                new EventHandler<SpeechDetectedEventArgs>(recognizer_SpeechDetected);
            recognizer.LoadGrammarCompleted +=
                new EventHandler<LoadGrammarCompletedEventArgs>(recognizer_LoadGrammarCompleted);

            // Set an empty grammar
            // FIXME: it's possible to refactor here:
            // don't start recognition until any grammar is provided using input pins
            //setGrammar(new string[][] { new string[] { "default" } });

            // Configure the input to the speech recognizer
            recognizer.SetInputToDefaultAudioDevice();

            // Start the recognition
            //startRecognition();
        }

        void startSpeechRecognition()
        {
            lock (startRecognitionLock)
            {
                if (!RecognitionStarted && isGrammarLoaded())
                {
                    // Start asynchronous, continuous speech recognition
                    recognizer.RecognizeAsync(RecognizeMode.Multiple);
                    RecognitionStarted = true;
                }
            }
        }

        void stopSpeechRecongition()
        {
            // Terminate asynchronous recognition immediately
            recognizer.RecognizeAsyncCancel();
            RecognitionStarted = false;
        }

        bool isGrammarLoaded()
        {
            return recognizer != null && 
                   currentGrammar != null &&
                   recognizer.Grammars.Count > 0;
        }

        void unloadCurrentGrammar()
        {
            if (!isGrammarLoaded())
            {
                return;
            }
            lock (asyncActionMonitor)
            {
                // Request an update and unload the current grammar
                recognizer.RequestRecognizerUpdate(
                    new grammarActionDelegate(grammarActionUnloadAllGrammars));
                // Syncronization
                Monitor.Wait(asyncActionMonitor);
            } 


            // Reset the current grammar instance to null
            currentGrammar = null;
        }

        // Delegate for performing grammar updates
        public delegate void grammarActionDelegate();

        public void grammarActionUnloadAllGrammars()
        {
            lock (asyncActionMonitor)
            {
                recognizer.UnloadAllGrammars();
                // Syncronization
                Monitor.PulseAll(asyncActionMonitor);
            } 
        }

        public void grammarActionLoadPreparedGrammar()
        {
            if (currentGrammar == null)
            {
                return;
            }

            //lock (asyncActionMonitor)
            //{
            // Request an update and unload the current grammar
            //recognizer.LoadGrammarAsync(currentGrammar);
            try
            {
                recognizer.LoadGrammar(currentGrammar);
            }
            catch (InvalidOperationException)
            {
                if (FLogger != null)
                {
                    FLogger.Log(LogType.Error,
                                "The grammar isn't suitable for the current culture. Try either changing the culture name, or the grammar");
                }
            }

            // Syncronization
            //    Monitor.Wait(asyncActionMonitor);
            //}
        }

        //public void grammarActionLoadPreparedGrammarAndStartRecognition()
        //{
        //    // Load grammar
        //    grammarActionLoadPreparedGrammar();
        //    // Start recognition if this is initial grammar
        //    // Won't affect anything if recognition already started
        //    startRecognition();
        //}

        void setGrammar(string[][] choicesStringsArray, string cultureName)
        {
            // Unload the old (current) grammar
            unloadCurrentGrammar();

            // Construct the new grammar
            GrammarBuilder grammarBuilder = new GrammarBuilder();
            try
            {
                // Thanks to http://stackoverflow.com/a/25190005/1155958
                // More examples: http://blog.qurbit.com/tutorials/building-grammars-in-net/
                grammarBuilder.Culture = new System.Globalization.CultureInfo(cultureName);
            }
            catch (System.Globalization.CultureNotFoundException)
            {
                if (FLogger != null)
                {
                    FLogger.Log(LogType.Error,
                                "Culture with name \"" + cultureName + "\" not found");
                }
                FRecognizerForCultureFound[0] = false;
                return;
            }

            // Whether we have at least one phrase/choice to construct grammar
            bool grammarNotEmpty = false;

            // Append all choices strings to the new grammar
            foreach (string[] choicesStrings in choicesStringsArray)
            {
                try
                {
                    if (choicesStrings.Length > 1)
                    {
                        // Set of choices
                        Choices choices = new Choices(choicesStrings);
                        grammarBuilder.Append(choices);    
                    }
                    else
                    {
                        // A single phrase
                        grammarBuilder.Append(choicesStrings[0]);
                    }
                    // Now we have at least one element in the grammar
                    grammarNotEmpty = true;
                }
                catch (ArgumentException)
                {
                    if (FLogger != null)
                    {
                        FLogger.Log(LogType.Error, "Invalid argument for grammar builder");
                    }
                    continue;
                }
            }

            // Construct the new grammar only if there is at least one element for it
            if (grammarNotEmpty)
            {
                // Create a Grammar object from the GrammarBuilder and load it to the recognizer
                currentGrammar = new Grammar(grammarBuilder);
                currentGrammar.Name = "vvvv";
                // Request an update and load the new grammar
                //var actionDelegate = startRecognition ? new grammarActionDelegate(grammarActionLoadPreparedGrammarAndStartRecognition) : 
                //                                        new grammarActionDelegate(grammarActionLoadPreparedGrammar);
                lock (asyncActionMonitor)
                {
                    recognizer.RequestRecognizerUpdate(new grammarActionDelegate(grammarActionLoadPreparedGrammar));
                    // Syncronization
                    Monitor.Wait(asyncActionMonitor);
                }
            }
        }
        #endregion

        #region evalueate
		public void Evaluate(int SpreadMax)
		{
            // Bang behavior
            if (onRecognizedBangFrameElapsed/* && FOnRecognized[0] == true*/)
            {
                FRecognitionResult[0] = "";
                FOnRecognized[0] = false;
            }
            onRecognizedBangFrameElapsed = true;

            // Culture name has been changed
            if (FCultureName.IsChanged)
            {
                reinitialize(FCultureName[0]);

                if (recognizer != null)
                {
                    lock (asyncActionMonitor)
                    {
                        recognizer.RequestRecognizerUpdate(new grammarActionDelegate(grammarActionLoadPreparedGrammar));
                        // Syncronization
                        Monitor.Wait(asyncActionMonitor);
                    }
                }
                else
                {
                    if (FLogger != null)
                    {
                        FLogger.Log(LogType.Error,
                                    "startRecognition(): SpeechRecognitionEngine instance hasn't been created yet");
                    }
                }

                // Update "grammar loaded" output pin value 
                FGrammarLoaded[0] = isGrammarLoaded();

                if (FEnabled[0] == true)
                {
                    startSpeechRecognition();
                }
            }

            // Enable pin has been changed
            if (FEnabled.IsChanged)
            {
                if (recognizer != null)
                {
                    if (FEnabled[0] == true)
                    {
                        startSpeechRecognition();
                    }
                    else
                    {
                        stopSpeechRecongition();
                    }
                }
                else
                {
                    if (FLogger != null)
                    {
                        FLogger.Log(LogType.Error,
                                    "startRecognition(): SpeechRecognitionEngine instance hasn't been created yet");
                    }
                }
            }

            // Grammar has been changed
            // FIXME: this is always being executed :(
            // Manual check of array equality?
            //if (FChoices.IsChanged)
            bool changedAtLeastOneChoices = false;
            {
                int numChoices = FChoicesCount[0];
                // Convert spread to array
                string[][] choicesStringsArray = new string[FChoices.SliceCount][];
                for (int i = 0; i < numChoices; ++i)
                {
                    var choicesSpread = FChoices[i].IOObject;
                    if (choicesSpread.IsChanged)
                    {
                        changedAtLeastOneChoices = true;
                    }
                    choicesStringsArray[i] = new string[choicesSpread.SliceCount];
                    for (int j = 0; j < choicesSpread.SliceCount; ++j)
                    {
                        choicesStringsArray[i][j] = choicesSpread[j];
                    }
                }

                if (changedAtLeastOneChoices)
                {
                    if (recognizer != null)
                    {
                        // Try to load grammar
                        setGrammar(choicesStringsArray, FCultureName[0]);
                    }
                    else
                    {
                        if (FLogger != null)
                        {
                            FLogger.Log(LogType.Error,
                                        "startRecognition(): SpeechRecognitionEngine instance hasn't been created yet");
                        }
                    }

                    // Update "grammar loaded" output pin value 
                    FGrammarLoaded[0] = isGrammarLoaded();

                    if (FEnabled[0] == true)
                    {
                        startSpeechRecognition();
                    }
                }
            }

            if (FEnabled.IsChanged || FCultureName.IsChanged || changedAtLeastOneChoices)
            {
                // Reset main output pins
                FRecognitionResult[0] = "";
                FOnRecognized[0] = false;
            }
		}
        
        #endregion

        #region speech event handlers
        // At the update, get the names and enabled status of the currently loaded grammars
        void recognizer_RecognizerUpdateReached(object sender, RecognizerUpdateReachedEventArgs e)
        {
            lock (asyncActionMonitor)
            {
                // Recognized is ready for update: call the delegate method
                grammarActionDelegate action = (grammarActionDelegate)e.UserToken;
                action();

                // Update "grammar loaded" pin value after update
                //FGrammarLoaded[0] = isGrammarLoaded();

                if (FLogger != null)
                {
                    FLogger.Log(LogType.Debug, "Update reached:");

                    string qualifier;
                    List<Grammar> grammars = new List<Grammar>(recognizer.Grammars);
                    foreach (Grammar g in grammars)
                    {
                        qualifier = (g.Enabled) ? "enabled" : "disabled";
                        FLogger.Log(LogType.Debug, "  {0} grammar is loaded and {1}.", g.Name, qualifier);
                    }
                }

                // Syncronization
                Monitor.PulseAll(asyncActionMonitor);
            } 
        }

        // Handle the SpeechRecognized event.
        void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Confidence >= FConfidenceThreshold[0])
            {
                FOnRecognized[0] = true;
            }
            FRecognitionResult[0] = e.Result.Text;
            FConfidence[0] = e.Result.Confidence;
            FOnSpeechDetected[0] = false;

            onRecognizedBangFrameElapsed = false;

            if (FLogger != null)
            {
                FLogger.Log(LogType.Debug, "Recognized text: " + e.Result.Text +
                                            String.Format("; Confidence = {0}", e.Result.Confidence));
            }
        }

        // Write a message to the console when recognition fails.
        void recognizer_SpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            if (FLogger != null)
            {
                FLogger.Log(LogType.Debug, "Recognition attempt failed");
            }

            FOnSpeechDetected[0] = false;
        }

        void recognizer_SpeechDetected(object sender, SpeechDetectedEventArgs e)
        {
            FOnSpeechDetected[0] = true;
        }

        void recognizer_LoadGrammarCompleted(object sender, LoadGrammarCompletedEventArgs e)
        {
            lock (asyncActionMonitor)
            {
                // Syncronization
                Monitor.PulseAll(asyncActionMonitor);
            } 
        }
        #endregion
    }
}
