using System;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Speech.Synthesis;
using System.Threading;
using Rainmeter;

namespace PluginSpeech
{
    class Measure
    {
        static public implicit operator Measure(IntPtr data)
        {
            return (Measure)GCHandle.FromIntPtr(data).Target;
        }
        public IntPtr buffer = IntPtr.Zero;
        public SpeechSynthesizer synth;
        public ReadOnlyCollection<InstalledVoice> voices;
        public int voiceCount;
        public VoiceGender gender;
        public int index;
        public Prompt prompt;
        public String selectedName;
    }

    public class Plugin
    {
        [DllExport]
        public static void Initialize(ref IntPtr data, IntPtr rm)
        {
            data = GCHandle.ToIntPtr(GCHandle.Alloc(new Measure()));
            Rainmeter.API api = (Rainmeter.API)rm;

            Measure measure = (Measure)data;
            measure.synth = new SpeechSynthesizer();
            measure.prompt = new Prompt("");

            measure.voices = measure.synth.GetInstalledVoices();
            measure.voiceCount = measure.voices.Count;

            measure.synth.SetOutputToDefaultAudioDevice();

            int debug = api.ReadInt("Debug", 0);

            if (debug == 1)
            {
                api.Log(API.LogType.Notice, "------------------------------");
                api.Log(API.LogType.Notice, "* Speech.dll - GetInstalledVoices() API");

                int index = 1;
                foreach (InstalledVoice voice in measure.voices)
                {
                    VoiceInfo info = voice.VoiceInfo;

                    api.LogF(API.LogType.Notice, "Index: {0}", index);
                    api.LogF(API.LogType.Notice, "  Name          : {0}", info.Name);
                    api.LogF(API.LogType.Notice, "  Gender        : {0}", info.Gender);
                    api.LogF(API.LogType.Notice, "  Culture       : {0}", info.Culture);
                    api.LogF(API.LogType.Notice, "  Age           : {0}", info.Age);
                    api.LogF(API.LogType.Notice, "  Description   : {0}", info.Description);
                    api.LogF(API.LogType.Notice, "  ID            : {0}", info.Id);
                    api.LogF(API.LogType.Notice, "  Enabled       : {0}", voice.Enabled);

                    ++index;
                }

                api.Log(API.LogType.Notice, "------------------------------");
            }
        }

        [DllExport]
        public static void Finalize(IntPtr data)
        {
            Measure measure = (Measure)data;

            if (!measure.prompt.IsCompleted)
            {
                measure.synth.SpeakAsyncCancelAll();
            }

            measure.synth.Dispose();

            if (measure.buffer != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(measure.buffer);
            }
            GCHandle.FromIntPtr(data).Free();
        }

        [DllExport]
        public static void Reload(IntPtr data, IntPtr rm, ref double maxValue)
        {
            Measure measure = (Measure)data;
            Rainmeter.API api = (Rainmeter.API)rm;

            string name = api.ReadString("Name", "");
            bool nameExists = name != "" ? true : false;

            VoiceGender gender = VoiceGender.NotSet;
            string genderStr = api.ReadString("Gender", "").ToUpper();
            if (genderStr == "MALE")
            {
                gender = VoiceGender.Male;
            }
            else if (genderStr == "FEMALE")
            {
                gender = VoiceGender.Female;
            }
            else if (genderStr != "")
            {
                api.Log(API.LogType.Warning, "Speech.dll: Invalid gender");
            }
            bool genderExists = gender != VoiceGender.NotSet;
            measure.gender = gender;

            int index = api.ReadInt("Index", 0);
            measure.index = index;

            int volume = api.ReadInt("Volume", 100);
            if (volume < 0 || volume > 100)
            {
                volume = 100;
            }

            int rate = api.ReadInt("Rate", 0);
            if (rate < -10 || rate > 10)
            {
                rate = 0;
            }

            measure.synth.Volume = volume;
            measure.synth.Rate = rate;

            // Setup voice
            int i = 1;
            String selectedName = "";
            foreach (InstalledVoice voice in measure.voices)
            {
                // If no options are "defined", use the first voice.
                if (!nameExists && !genderExists && (index <= 0))
                {
                    selectedName = voice.VoiceInfo.Name;
                    break;
                }

                // If name is defined, use this voice.
                if (nameExists && (name.ToUpper() == voice.VoiceInfo.Name.ToUpper()))
                {
                    selectedName = voice.VoiceInfo.Name;
                    break;
                }

                // If gender is defined, only look for voices of the same gender
                if (!nameExists && genderExists)
                {
                    if (gender != voice.VoiceInfo.Gender) continue;

                    if (index > 0)
                    {
                        if (index == i)
                        {
                            selectedName = voice.VoiceInfo.Name;
                            break;
                        }

                        ++i;
                        continue;
                    }
                    else
                    {
                        selectedName = voice.VoiceInfo.Name;
                        break;
                    }
                }

                // Select by index
                if (!nameExists && index == i)
                {
                    selectedName = voice.VoiceInfo.Name;
                    break;
                }

                ++i;
            }

            measure.selectedName = selectedName;

            // Something went wrong
            //  Name is incorrect
            //  No matching gender was found
            //  Invalid index
            if (selectedName == "")
            {
                api.Log(API.LogType.Warning, "Speech.dll: Invalid Name, Gender and/or Index. Using best matching valid voice.");
            }
        }

        [DllExport]
        public static double Update(IntPtr data)
        {
            Measure measure = (Measure)data;
            return Convert.ToDouble(measure.voiceCount);
        }

        [DllExport]
        public static void ExecuteBang(IntPtr data, [MarshalAs(UnmanagedType.LPWStr)]String args)
        {
            Measure measure = (Measure)data;

            if (!measure.prompt.IsCompleted)
            {
                measure.synth.SpeakAsyncCancelAll();
                Thread.Sleep(10);
            }

            if (measure.selectedName == "")
            {
                measure.synth.SelectVoiceByHints(measure.gender, VoiceAge.NotSet, measure.index);
            }
            else
            {
                measure.synth.SelectVoice(measure.selectedName);
            }

            measure.prompt = measure.synth.SpeakAsync(args);
        }

        [DllExport]
        public static IntPtr GetString(IntPtr data)
        {
            Measure measure = (Measure)data;
            if (measure.buffer != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(measure.buffer);
                measure.buffer = IntPtr.Zero;
            }

            measure.buffer = Marshal.StringToHGlobalUni(measure.selectedName);

            return measure.buffer;
        }

        //[DllExport]
        //public static IntPtr (IntPtr data, int argc,
        //    [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 1)] string[] argv)
        //{
        //    Measure measure = (Measure)data;
        //    if (measure.buffer != IntPtr.Zero)
        //    {
        //        Marshal.FreeHGlobal(measure.buffer);
        //        measure.buffer = IntPtr.Zero;
        //    }
        //
        //    measure.buffer = Marshal.StringToHGlobalUni("");
        //
        //    return measure.buffer;
        //}
    }
}
