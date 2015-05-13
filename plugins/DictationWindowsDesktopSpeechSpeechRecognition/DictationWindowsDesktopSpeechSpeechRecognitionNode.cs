// By Lev Panov, 2015 (lev.panov@gmail.com)
// Further work
// + 1. Repeated/optional phrases https://msdn.microsoft.com/en-us/library/ms576572(v=vs.110).aspx
// + Possible to arrange repetition count pins in a way similar to group (ex9 priority)
// 2. Semantic information
// 3. Free text dictation: https://msdn.microsoft.com/en-us/library/ms576565(v=vs.110).aspx
// Not possible with server-side Microsoft.Speech, need to use System.Speech
// + 4. Different languages
// 5. Configurable single/multiple recognition modes
// 6. Output phrase index (for complex grammar - slice of indices for each choice?)
// http://vvvv.org/documentation/dynamic-plugins-reference
// http://vvvv.org/forum/dynamic-amount-of-input-pins-in-a-plugin-(like-stallone)?full=1
// https://searchcode.com/codesearch/view/26524082/
// http://vvvv.org/pluginspecs/html/N_VVVV_PluginInterfaces_V2.htm
// http://stackoverflow.com/questions/12353605/speech-recognition-in-c-sharp-with-sapi-5-4-or-ms-speech-sdk-v11-using-a-memorys
// http://stackoverflow.com/questions/15258987/does-this-system-speech-recognition-recognition-code-make-use-of-speech-trainin
// http://stackoverflow.com/questions/2977338/what-is-the-difference-between-system-speech-recognition-and-microsoft-speech-re

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

using System.Speech.Recognition;
#endregion usings

namespace VVVV.Nodes
{
    #region PluginInfo
    [PluginInfo(Name = "SpeechRecognition",
                Category = "Windows Desktop Speech",
                Version = "Dictation",
                Help = "Speech recognition using dictation in different languages",
                Tags = "speech",
                Author = "lev")]
    #endregion PluginInfo

    public class DictationWindowsDesktopSpeechSpeechRecognitionNode :
        IPluginEvaluate, IPartImportsSatisfiedNotification, IDisposable
    {
        #region input pins

        [Input("Enabled", IsSingle = true, DefaultBoolean = true, Order = 0)]
        IDiffSpread<bool> FEnabled;

        [Input("Culture Name", EnumName = "CultureNameEnum", Order = 1)]
        public IDiffSpread<EnumEntry> FCultureNameEnum;

        [Input("Confidence Threshold", IsSingle = true, DefaultValue = 0.7, Order = 2)]
        IDiffSpread<double> FConfidenceThreshold;

        public Spread<IIOContainer<ISpread<string>>> FChoices = new Spread<IIOContainer<ISpread<string>>>();
        public Spread<IIOContainer<ISpread<bool>>> FIsOptional = new Spread<IIOContainer<ISpread<bool>>>();

        [Config("Choices Count", DefaultValue = 1, MinValue = 1)]
        public IDiffSpread<int> FChoicesCount;

        #endregion

        #region output pins

        // TODO
        //[Output("Choices indices", IsPinGroup = true, Order = 4)]
        //ISpread<ISpread<string>> FChoices;

        [Output("Recognition Result", IsSingle = true)]
        public ISpread<String> FRecognitionResult;

        [Output("On Recognized", IsSingle = true, IsBang = true)]
        public ISpread<bool> FOnRecognized;

        [Output("Confidence", IsSingle = true)]
        public ISpread<double> FConfidence;

        [Output("On Speech Detected", IsSingle = true, IsToggle = true)]
        public ISpread<bool> FOnSpeechDetected;

        [Output("Recognizer for culture found", IsSingle = true, IsToggle = true)]
        public ISpread<bool> FRecognizerForCultureFound;

        [Output("Grammar Loaded", IsSingle = true, IsToggle = true)]
        public ISpread<bool> FGrammarLoaded;

        #endregion

        #region imports

        [Import()]
        public ILogger FLogger;
        [Import]
        public IIOFactory FIOFactory;

        #endregion

        #region fields

        private int choicesOrderOffset = 3;
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

        // This region code is mostly copy-pasted from Template (Value DynamicPins)
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

        InputAttribute createChoicesInputAttributeWithChangeCheck(int i)
        {
            var ia = new InputAttribute(string.Format("Choices {0}", i));
            ia.CheckIfChanged = true;
            ia.Order = choicesOrderOffset + i * 2;
            return ia;
        }

        InputAttribute createIsOptionalInputAttributeWithChangeCheck(int i)
        {
            // http://vvvv.org/pluginspecs/html/AllMembers_T_VVVV_PluginInterfaces_V2_InputAttribute.htm#propertyTableToggle
            var ia = new InputAttribute(string.Format("Is Optional {0}", i));
            ia.CheckIfChanged = true;
            ia.Order = choicesOrderOffset + i * 2 + 1;
            ia.DefaultBoolean = false;
            ia.Visibility = PinVisibility.OnlyInspector; // OnlyInspector/Hidden
            return ia;
        }

        private void HandleInputCountChanged(IDiffSpread<int> sender)
        {
            HandlePinCountChanged(sender, FChoices, createChoicesInputAttributeWithChangeCheck);
            HandlePinCountChanged(sender, FIsOptional, createIsOptionalInputAttributeWithChangeCheck);
        }
        #endregion

        #region ctor & dispose

        [ImportingConstructor]
        public DictationWindowsDesktopSpeechSpeechRecognitionNode()
        {
            // Populate enum with available for speech recognition cultures
            // TODO: Select a speech recognizer that supports English by default:
            // https://msdn.microsoft.com/en-us/library/system.speech.recognition.speechrecognitionengine.installedrecognizers(v=vs.110).aspx
            var recognizers = SpeechRecognitionEngine.InstalledRecognizers();
            bool[] isValidRecognizer = new bool[recognizers.Count];
            int numValidRecognizers = 0;
            for (int i = 0; i < recognizers.Count; ++i)
            {
                isValidRecognizer[i] = recognizers[i].Culture.Name.Length > 0;
                numValidRecognizers += Convert.ToInt32(isValidRecognizer[i]);
            }

            var recognizersCultureNames = new string[numValidRecognizers];
            int enumIndex = 0;
            for (int i = 0; i < recognizers.Count; ++i)
            {
                if (isValidRecognizer[i])
                {
                    recognizersCultureNames[enumIndex] = recognizers[i].Culture.Name;
                    enumIndex += 1;
                }
            }

            EnumManager.UpdateEnum("CultureNameEnum", recognizersCultureNames[0], recognizersCultureNames);
        }

        public void Dispose()
        {
            if (recognizer != null)
            {
                ((IDisposable)recognizer).Dispose();
                recognizer = null;
            }
        }
        #endregion

        #region log helpers

        void logError(string message)
        {
            if (FLogger != null)
                FLogger.Log(LogType.Error, message);
        }
        void logDebug(string message)
        {
            if (FLogger != null)
                FLogger.Log(LogType.Debug, message);
        }
        void logDebug(string message, object arg)
        {
            if (FLogger != null)
                FLogger.Log(LogType.Debug, message, arg);
        }
        void logDebug(string message, object arg1, object arg2)
        {
            if (FLogger != null)
                FLogger.Log(LogType.Debug, message, arg1, arg2);
        }

        #endregion

        #region private methods

        void reinitialize(string cultureName)
        {
            // Dispose old recognizer if present
            Dispose();
            // Reset all needed class memebers
            recognitionStarted = false;
            // Two nested try-catch blocks: ugly, I know
            try
            {
                // Can throw System.Globalization.CultureNotFoundException
                var culture = new System.Globalization.CultureInfo(cultureName);
                // Culture info created, will try to create a recognition engine now...
                try
                {
                    // Can throw ArgumentException in case if no recognizer for specified culture exists 
                    recognizer = new SpeechRecognitionEngine(culture);
                }
                catch (ArgumentException)
                {
                    logError("No recognized found for culture \"" + cultureName + "\"");
                    return;
                }
            }
            catch (System.Globalization.CultureNotFoundException)
            {
                logError("Culture with name \"" + cultureName + "\" not found");
                return;
            }

            // Attach handlers to all needed events
            recognizer.SpeechRecognized +=
                new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);
            recognizer.RecognizerUpdateReached +=
                new EventHandler<RecognizerUpdateReachedEventArgs>(recognizer_RecognizerUpdateReached);
            recognizer.SpeechRecognitionRejected +=
                new EventHandler<SpeechRecognitionRejectedEventArgs>(recognizer_SpeechRecognitionRejected);
            recognizer.SpeechDetected +=
                new EventHandler<SpeechDetectedEventArgs>(recognizer_SpeechDetected);
            recognizer.LoadGrammarCompleted +=
                new EventHandler<LoadGrammarCompletedEventArgs>(recognizer_LoadGrammarCompleted);

            // Configure the input to the speech recognizer
            recognizer.SetInputToDefaultAudioDevice();
        }

        void startSpeechRecognition()
        {
            lock (startRecognitionLock)
            {
                if (!RecognitionStarted && GrammarLoaded)
                {
                    // Start asynchronous, continuous speech recognition
                    recognizer.RecognizeAsync(RecognizeMode.Multiple);
                    RecognitionStarted = true;
                }
            }
        }

        void stopSpeechRecongition()
        {
            if (RecognitionStarted)
            {
                // Terminate asynchronous recognition immediately
                recognizer.RecognizeAsyncCancel();
                RecognitionStarted = false;

            }
        }
        private bool RecognizerPresent
        {
            get
            {
                return recognizer != null;
            }
        }

        private bool GrammarLoaded
        {
            get
            {
                return RecognizerPresent &&
                       currentGrammar != null &&
                       recognizer.Grammars.Count > 0;
            }
        }

        void unloadCurrentGrammar()
        {
            if (!GrammarLoaded)
                return;

            lock (asyncActionMonitor)
            {
                // Request an update and unload the current grammar
                recognizer.RequestRecognizerUpdate(
                    new grammarActionDelegate(grammarActionUnloadAllGrammars));
                // Syncronization
                Monitor.Wait(asyncActionMonitor);
            }
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

            // FIXME: remove this unused async code at some point
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
                logError("The grammar isn't suitable for the current culture. " +
                         "Try either changing the culture name, or the grammar");
            }
            // Syncronization
            //    Monitor.Wait(asyncActionMonitor);
            //}
        }

        void setGrammar(string[][] choicesStringsArray, bool[] isOptionalArray, string cultureName)
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
                logError("Culture with name \"" + cultureName + "\" not found");
                return;
            }

            // Whether we have at least one phrase/choice to construct grammar
            bool grammarNotEmpty = false;

            // Append all choices strings to the new grammar
            for (int i = 0; i < choicesStringsArray.Length; ++i)
            {
                string[] choicesStrings = choicesStringsArray[i];
                bool isOptional = isOptionalArray[i];

                try
                {
                    int min = isOptional ? 0 : 1, max = 1;
                    if (choicesStrings.Length > 1)
                    {
                        // Set of choices
                        Choices choices = new Choices(choicesStrings);
                        // https://msdn.microsoft.com/en-us/library/system.speech.recognition.choices(v=vs.110).aspx
                        grammarBuilder.Append(new GrammarBuilder(choices), minRepeat: min, maxRepeat: max);
                    }
                    else
                    {
                        // A single phrase
                        grammarBuilder.Append(choicesStrings[0], minRepeat: min, maxRepeat: max);
                    }
                    // Now we have at least one element in the grammar
                    grammarNotEmpty = true;
                }
                catch (ArgumentException)
                {
                    logError("Invalid argument for grammar builder");
                    continue;
                }
            }

            // Construct the new grammar only if there is at least one element for it
            if (grammarNotEmpty)
            {
                // Create a Grammar object from the GrammarBuilder and load it to the recognizer
                currentGrammar = new Grammar(grammarBuilder);
                currentGrammar.Name = "vvvv";

                lock (asyncActionMonitor)
                {
                    // Request an update and load the new grammar
                    recognizer.RequestRecognizerUpdate(new grammarActionDelegate(grammarActionLoadPreparedGrammar));
                    // Syncronization
                    Monitor.Wait(asyncActionMonitor);
                }
            }
        }
        #endregion

        #region evaluate

        public void Evaluate(int SpreadMax)
        {
            // Bang behavior
            // FIXME: second statement needed or not?
            if (onRecognizedBangFrameElapsed/* && FOnRecognized[0] == true*/)
            {
                FRecognitionResult[0] = "";
                FOnRecognized[0] = false;
            }
            onRecognizedBangFrameElapsed = true;

            // Culture name has been changed
            if (FCultureNameEnum.IsChanged)
            {
                // Need to reinitialize the recognition engine in this case
                reinitialize(FCultureNameEnum[0].Name);
                // If all went ok, reload the grammar next
                if (RecognizerPresent)
                {
                    lock (asyncActionMonitor)
                    {
                        recognizer.RequestRecognizerUpdate(new grammarActionDelegate(grammarActionLoadPreparedGrammar));
                        // Syncronization
                        Monitor.Wait(asyncActionMonitor);
                    }
                }

                // Update output pins values 
                FGrammarLoaded[0] = GrammarLoaded;
                FRecognizerForCultureFound[0] = RecognizerPresent;

                // Start speech recognition if needed
                if (FEnabled[0] == true)
                    startSpeechRecognition();
            }

            // Enable pin has been changed
            if (FEnabled.IsChanged)
            {
                if (RecognizerPresent)
                {
                    if (FEnabled[0] == true)
                        startSpeechRecognition();
                    else
                        stopSpeechRecongition();
                }
                else
                {
                    logError("Can't start/stop speech recognition: engine is not present");
                }
            }

            // Check if grammar structure changed: can't use FChoices.IsChanged here: it's always true
            bool grammarChanged = false;
            int numChoices = FChoicesCount[0];
            for (int i = 0; i < numChoices; ++i)
            {
                var choicesSpread = FChoices[i].IOObject;
                var isOptionalSpread = FIsOptional[i].IOObject;

                if (choicesSpread.IsChanged || isOptionalSpread.IsChanged)
                {
                    grammarChanged = true;
                    break;
                }
            }
            // Reload grammar if it's structure has been changed
            if (grammarChanged)
            {
                // Convert spreads to arrays
                string[][] choicesStringsArray = new string[numChoices][];
                bool[] isOptionalArray = new bool[numChoices];
                for (int i = 0; i < numChoices; ++i)
                {
                    var choicesSpread = FChoices[i].IOObject;
                    choicesStringsArray[i] = new string[choicesSpread.SliceCount];
                    for (int j = 0; j < choicesSpread.SliceCount; ++j)
                    {
                        choicesStringsArray[i][j] = choicesSpread[j];
                    }

                    var isOptionalSpread = FIsOptional[i].IOObject;
                    isOptionalArray[i] = isOptionalSpread[i];
                }

                if (RecognizerPresent)
                {
                    // Try to load grammar
                    setGrammar(choicesStringsArray, isOptionalArray, FCultureNameEnum[0].Name);
                }
                else
                {
                    logError("Can't reload grammar: speech recognition engine is not present");
                }

                // Update "grammar loaded" output pin value 
                FGrammarLoaded[0] = GrammarLoaded;

                // Start recognition if needed
                if (FEnabled[0] == true)
                {
                    startSpeechRecognition();
                }
            }

            // Reset needed output pins if any parameters has been changed
            if (FEnabled.IsChanged || FCultureNameEnum.IsChanged || grammarChanged)
            {
                FRecognitionResult[0] = "";
                FOnRecognized[0] = false;
                FConfidence[0] = 0.0;
            }
        }

        #endregion

        #region speech event handlers

        void recognizer_RecognizerUpdateReached(object sender, RecognizerUpdateReachedEventArgs e)
        {
            // Monitor actions should be performed inside a lock()
            lock (asyncActionMonitor)
            {
                // Recognized is ready for update: obtain and call the delegate method
                grammarActionDelegate action = (grammarActionDelegate)e.UserToken;
                action();

                // Debug: get the names and enabled status of the currently loaded grammars at the update
                logDebug("Update reached:");

                string qualifier;
                List<Grammar> grammars = new List<Grammar>(recognizer.Grammars);
                foreach (Grammar g in grammars)
                {
                    qualifier = (g.Enabled) ? "enabled" : "disabled";
                    logDebug("  {0} grammar is loaded and {1}.", g.Name, qualifier);
                }

                // Tell the main process that the asyncronious action is finished
                Monitor.PulseAll(asyncActionMonitor);
            }
        }

        void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            // OnRecognized bang is set to true only if the confidence threshold has been exceeded
            if (e.Result.Confidence >= FConfidenceThreshold[0])
            {
                FOnRecognized[0] = true;
            }
            FRecognitionResult[0] = e.Result.Text;
            FConfidence[0] = e.Result.Confidence;
            FOnSpeechDetected[0] = false;

            onRecognizedBangFrameElapsed = false;

            logDebug("Recognized text: " + e.Result.Text + String.Format("; Confidence = {0}", e.Result.Confidence));
        }

        void recognizer_SpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            //logDebug("Recognition attempt failed");

            FOnSpeechDetected[0] = false;
        }

        void recognizer_SpeechDetected(object sender, SpeechDetectedEventArgs e)
        {
            FOnSpeechDetected[0] = true;
        }

        void recognizer_LoadGrammarCompleted(object sender, LoadGrammarCompletedEventArgs e)
        {
            // Monitor actions should be performed inside a lock()
            lock (asyncActionMonitor)
            {
                // Tell the main process that the asyncronious grammar loading is finished
                Monitor.PulseAll(asyncActionMonitor);
            }
        }
        #endregion
    }
}
