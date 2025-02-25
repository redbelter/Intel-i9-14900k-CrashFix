using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intel_i9_14900k_CrashFix
{
    public class SetIntelCPUSpeed
    {
        public const string INTEL_TUNING_PATH = "C:\\Program Files\\Intel\\Intel(R) Extreme Tuning Utility\\Client\\PerfTune.exe";
        public const string X_SPEED_50 = "50x";
        public const string TUNING_COMBO_BOX_ID = "TuningCombo:3489660933";
        public const string INTEL_PROCESS_NAME = "PerfTune";
        public const string INTEL_HEADER_PROCESSOR_NAME_ID = "HeaderBrandString";
        public static bool CheckIfIntelTuningIsInstalled()
        {
            if (File.Exists(INTEL_TUNING_PATH))
            {
                Console.WriteLine("Intel Tuning Utility is installed");
                return true;
            }
            else
            {
                Console.WriteLine("Intel Tuning Utility is not installed");
                return false;
            }
        }
        public static void KillIntelProcess()
        {
            var process = Process.GetProcessesByName(INTEL_PROCESS_NAME);
            process[0].Kill();
            process[0].WaitForExit();
        }
        
        public static Window? TryToFindWindow(UIA3Automation automation, bool nested = false)
        {
            KillIntelProcess(); //THis makes the code below more stable
            ProcessStartInfo startInfo = new ProcessStartInfo(INTEL_TUNING_PATH);

            Console.WriteLine("Launching or attaching to intel process");
            var app = FlaUI.Core.Application.AttachOrLaunch(startInfo);
            var window = app.GetMainWindow(automation);
            var combo = FindTuningComboBox(window);
            int giveupinitial = 100;
            Console.WriteLine("Trying to find intel window...");
            while (combo == null)
            {
                //Wait for the window to pop up
                window = app.GetMainWindow(automation);
                combo = FindTuningComboBox(window);
                Thread.Sleep(500);
                giveupinitial--;
                if (giveupinitial == 0)
                {
                    Console.WriteLine("Window might be in systray, killing process");
                    KillIntelProcess();

                    //Relaunch it
                    app = FlaUI.Core.Application.AttachOrLaunch(startInfo);
                    if (!nested) { 
                        window = TryToFindWindow(automation, true); //Try one more time
                    }
                }
            }
            if(window != null)
            {
                Console.WriteLine("Found intel window");
            }
            return window;
        }

        public static ComboBox? FindTuningComboBox(Window window)
        {
            var combo = window.FindFirstDescendant(cf => cf.ByAutomationId(TUNING_COMBO_BOX_ID));
            if(combo == null)
            {
                return null;
            }
            return combo.AsComboBox();
        }

        public static bool IsBadProcessor(Window window)
        {
            var header = window.FindFirstDescendant(cf => cf.ByAutomationId(INTEL_HEADER_PROCESSOR_NAME_ID));
            if (header.Name.Equals("Intel(R) Core(TM) i9-14900K"))
            {
                Console.WriteLine("Known bad processor");
                return true; //supported;
            }
            Console.WriteLine("We dont know if this processor is good or bad.");
            return false;
        }

        public static void DoIt()
        {
            if (!CheckIfIntelTuningIsInstalled()){
                Environment.Exit(1);
            }
            using (var automation = new UIA3Automation())
            {
                var window = TryToFindWindow(automation, false);
                if(window == null)
                {
                    Console.WriteLine("Failed to find the intel window");
                    Environment.Exit(1);
                }
                if(!IsBadProcessor(window))
                {
                    Console.WriteLine("This processor is not a known bad one");
                    Environment.Exit(1);
                }

                var cpuspeedcombo = FindTuningComboBox(window);
                if(cpuspeedcombo == null)
                {
                    Console.WriteLine("Failed to find the cpu speed combo box");
                    Environment.Exit(1);
                }

                if (cpuspeedcombo.SelectedItem.Text == X_SPEED_50)
                {
                    Console.WriteLine("Already set to " + X_SPEED_50);
                    window.Close();
                    return;
                }
                cpuspeedcombo.Select(X_SPEED_50);
                cpuspeedcombo.Click();
                var apply = window.FindFirstDescendant(cf => cf.ByText("Apply"));
                int maxWait = 1000;
                while (!apply.IsEnabled)
                {
                    Thread.Sleep(10);
                    maxWait--;
                    if(maxWait == 0)
                    {
                        Console.WriteLine("Apply button was grayed out, already good?");
                        Environment.Exit(0);
                    }
                }
                while (apply.IsEnabled) 
                { 
                    apply.Click();
                    Thread.Sleep(10);
                    if (maxWait == 0)
                    {
                        Console.WriteLine("Failed to apply");
                        Environment.Exit(1);
                    }
                }
                Console.WriteLine("Underclocked to 5ghz for stability");

                window.Close();
            }
        }
    }
}
