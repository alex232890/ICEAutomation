using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.AutomationElements.Infrastructure;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;


namespace ImageComposeEditorAutomation
{
    public class AdvancedComposeAppService
    {
        Action<string> onEvent;
        Action<int> onProgress;

        public void Compose(string[] images, Action<string> onEvent = null, Action<int> onProgress = null)
        {

            string[] lines = System.IO.File.ReadAllLines(@"C:\Projects\ICEAutomation\parameters.txt");
            IDictionary<string, string> parameters = new Dictionary<string, string>();
            foreach (string line in lines)
            {
                string[] words = line.Split('=');
                if (words.Count() > 1)
                {
                    parameters.Add(words[0], words[1]);
                }
            }
            this.onEvent = onEvent;
            this.onProgress = onProgress;
            var appStr = parameters["ice-app-path"];
            var appName = Path.GetFileName(appStr);

            var structuredPanoramaBtnLabel = ConfigurationManager.AppSettings["Structured-panorama-btn-label"];
            var primaryDirectionCountTextBox = ConfigurationManager.AppSettings["Primary-direction-count-text-box"];

            var startTopLeftHelpText = ConfigurationManager.AppSettings["Start-top-left-help-text"];
            var startTopRightHelpText = ConfigurationManager.AppSettings["Start-top-right-help-text"];
            var startBottomLeftHelpText = ConfigurationManager.AppSettings["Start-bottom-left-help-text"];
            var startBottomRightHelpText = ConfigurationManager.AppSettings["Start-bottom-right-help-text"];

            var moveRightHelpText = ConfigurationManager.AppSettings["Start-moving-right-help-text"];
            var moveLeftHelpText = ConfigurationManager.AppSettings["Start-moving-left-help-text"];
            var moveUpHelpText = ConfigurationManager.AppSettings["Start-moving-up-help-text"];
            var moveDownHelpText = ConfigurationManager.AppSettings["Start-moving-down-help-text"];

            var cornerHelpText = startBottomLeftHelpText;
            var directionHelpText = moveRightHelpText;

            if (Convert.ToBoolean(parameters["start-upper-left-corner"]) == true)
            {
                cornerHelpText = startTopLeftHelpText;

                if (Convert.ToBoolean(parameters["move-right"]) == true)
                {
                    directionHelpText = moveRightHelpText;
                }
                else if (Convert.ToBoolean(parameters["move-down"]) == true)
                {
                    directionHelpText = moveDownHelpText;
                }
            }
            else if (Convert.ToBoolean(parameters["start-upper-right-corner"]) == true)
            {
                cornerHelpText = startTopRightHelpText;

                if (Convert.ToBoolean(parameters["move-left"]) == true)
                {
                    directionHelpText = moveLeftHelpText;
                }
                else if (Convert.ToBoolean(parameters["move-down"]) == true)
                {
                    directionHelpText = moveDownHelpText;
                }
            }
            else if (Convert.ToBoolean(parameters["start-bottom-right-corner"]) == true)
            {
                cornerHelpText = startBottomRightHelpText;

                if (Convert.ToBoolean(parameters["move-left"]) == true)
                {
                    directionHelpText = moveLeftHelpText;
                }
                else if (Convert.ToBoolean(parameters["move-up"]) == true)
                {
                    directionHelpText = moveUpHelpText;
                }
            }
            else
            {
                cornerHelpText = startBottomLeftHelpText;

                if (Convert.ToBoolean(parameters["move-right"]) == true)
                {
                    directionHelpText = moveRightHelpText;
                }
                else if (Convert.ToBoolean(parameters["move-up"]))
                {
                    directionHelpText = moveUpHelpText;
                }
            }

            var serpentineBtnLabel = ConfigurationManager.AppSettings["Serpentine-btn-label"];
            var zigZagBtnLabel = ConfigurationManager.AppSettings["Zig-zag-btn-label"];

            var angularRangeBtnLabel = ConfigurationManager.AppSettings["Less-than-360-angular-range-btn-label"];

            if (Convert.ToBoolean(parameters["360-horizontal-angular-range"]) == true)
                angularRangeBtnLabel = ConfigurationManager.AppSettings["360-horizontal-btn-label"];
            else if (Convert.ToBoolean(parameters["360-vertical-angular-range"]) == true)
                angularRangeBtnLabel = ConfigurationManager.AppSettings["360-vertical-btn-label"];

            var horizontalOverlapTextBox = ConfigurationManager.AppSettings["Horizontal-overlap-text-box"];
            var verticalOverlapTextBox = ConfigurationManager.AppSettings["Vertical-overlap-text-box"];
            var searchRadiusTextBox = ConfigurationManager.AppSettings["Search-radius-text-box"];

            var exportBtnLabel = ConfigurationManager.AppSettings["Export-btn-label"];
            var exportToDiskBtnLabel = ConfigurationManager.AppSettings["ExportToDisk-btn-label"];
            var exportPanoramaBtnLabel = ConfigurationManager.AppSettings["ExportPanorama-btn-label"];
            var saveBtnLabel = ConfigurationManager.AppSettings["Save-btn-label"];
            int saveWait = int.Parse(ConfigurationManager.AppSettings["Save-wait"]);

            var imgStr = string.Join(" ", images);
            var processStartInfo = new ProcessStartInfo(fileName: appStr, arguments: imgStr);
            Console.WriteLine(appStr);
            var app = FlaUI.Core.Application.Launch(processStartInfo);

            using (var automation = new UIA3Automation())
            {
                string title = null;
                Window window = null;
                AutomationElement importPage = null;
                AutomationElement accordionPage = null;
                AutomationElement structuredPanorama = null;
                AutomationElement structuredPanoramaSettings = null;
                AutomationElement layout = null;
                AutomationElement overlap = null;

                do
                {
                    app = FlaUI.Core.Application.Attach(appName);
                    Console.WriteLine("app");
                    window = app.GetMainWindow(automation);
                    Console.WriteLine("automation");
                    importPage = window.FindFirstDescendant(cf => cf.ByClassName("ImportPage"));
                    Console.WriteLine("import");
                    accordionPage = importPage.FindFirstDescendant(cf => cf.ByClassName("Accordion"));
                    Console.WriteLine("accordion");
                    structuredPanorama = accordionPage.FindFirstDescendant(cf => cf.ByName("Structured panorama"));
                    Console.WriteLine("structured");
                    structuredPanoramaSettings = structuredPanorama.FindFirstDescendant(cf => cf.ByClassName("StructuredPanoramaSettings"));
                    Console.WriteLine("settings");
                    layout = structuredPanoramaSettings.FindFirstDescendant(cf => cf.ByName("Layout"));
                    Console.WriteLine("layout");
                    overlap = structuredPanoramaSettings.FindFirstDescendant(cf => cf.ByName("Overlap"));
                    Console.WriteLine("overlap");
                    title = window.Title;
                    OnEvent("Opened :" + title);
                } while (string.IsNullOrWhiteSpace(title));

                OnEvent("files: " + imgStr);

                try
                {
                    AutomationElement button1 = null;
                    do
                    {
                        button1 = structuredPanorama.FindFirstDescendant(cf => cf.ByText(structuredPanoramaBtnLabel));
                        if (button1 == null)
                        {
                            OnEvent(".");
                        }
                    } while (button1 == null);

                    if (button1.ControlType != ControlType.Button)
                        button1 = button1.AsButton().Parent;
                    button1?.AsButton().Click();
                    AutomationElement button2 = null;
                    do
                    {
                        button2 = layout.FindFirstDescendant(cf => cf.ByHelpText(cornerHelpText));
                        if (button2 == null)
                        {
                            OnEvent(".");
                        }
                    } while (button2 == null);

                    if (button2.ControlType != ControlType.RadioButton)
                        button2 = button2.AsRadioButton().Parent;

                    button2?.AsRadioButton().Click();

                    AutomationElement[] button3Possibilities = null;
                    AutomationElement button3 = null;
                    do
                    {
                        button3Possibilities = layout.FindAllDescendants(cf => cf.ByHelpText(directionHelpText));
                        for (int i = 0; i < button3Possibilities.Length; i += 1)
                        {
                            if (button3Possibilities[i].IsEnabled == true)
                            {
                                button3 = button3Possibilities[i];
                            }
                        }
                        if (button3 == null)
                        {
                            OnEvent(".");
                        }
                    } while (button3 == null);

                    if (button3.ControlType != ControlType.RadioButton)
                        button3 = button3.AsRadioButton().Parent;

                    button3?.AsRadioButton().Click();

                    AutomationElement button4 = null;
                    do
                    {
                        string imgOrderBtnLabel = null;
                        if (Convert.ToBoolean(parameters["is-serpentine"]) == true)
                        {
                            imgOrderBtnLabel = serpentineBtnLabel;
                        }
                        else
                        {
                            imgOrderBtnLabel = zigZagBtnLabel;
                        }
                        button4 = layout.FindFirstDescendant(cf => cf.ByAutomationId(imgOrderBtnLabel));
                        if (button4 == null)
                        {
                            OnEvent(".");
                        }
                    } while (button4 == null);

                    if (button4.ControlType != ControlType.RadioButton)
                        button4 = button4.AsRadioButton().Parent;

                    button4?.AsRadioButton().Click();

                    AutomationElement button5 = null;
                    do
                    {
                        button5 = layout.FindFirstDescendant(cf => cf.ByName(angularRangeBtnLabel));
                        if (button5 == null)
                        {
                            OnEvent(".");
                        }
                    } while (button5 == null);

                    if (button5.ControlType != ControlType.RadioButton)
                        button5 = button5.AsRadioButton().Parent;

                    button5?.AsRadioButton().Click();

                    AutomationElement textBox1 = null;
                    do
                    {
                        textBox1 = layout.FindFirstDescendant(cf => cf.ByAutomationId(primaryDirectionCountTextBox));
                        if (textBox1 == null)
                        {
                            OnEvent(".");
                        }
                    } while (textBox1 == null);

                    if (textBox1.ControlType != ControlType.Edit)
                        textBox1 = textBox1.AsTextBox().Parent;

                    textBox1?.AsTextBox().Click();
                    textBox1?.AsTextBox().Enter(parameters["major-axis-image-count"]);

                    AutomationElement textBox2 = null;
                    do
                    {
                        textBox2 = overlap.FindFirstDescendant(cf => cf.ByAutomationId(horizontalOverlapTextBox));
                        if (textBox2 == null)
                        {
                            OnEvent(".");
                        }
                    } while (textBox2 == null);

                    if (textBox2.ControlType != ControlType.Edit)
                        textBox2 = textBox2.AsTextBox().Parent;

                    textBox2?.AsTextBox().Enter(parameters["horizontal-overlap"]);

                    AutomationElement textBox3 = null;
                    do
                    {
                        textBox3 = overlap.FindFirstDescendant(cf => cf.ByAutomationId(verticalOverlapTextBox));
                        if (textBox3 == null)
                        {
                            OnEvent(".");
                        }
                    } while (textBox3 == null);

                    if (textBox3.ControlType != ControlType.Edit)
                        textBox3 = textBox3.AsTextBox().Parent;

                    textBox3?.AsTextBox().Enter(parameters["vertical-overlap"]);

                    AutomationElement textBox4 = null;
                    do
                    {
                        textBox4 = overlap.FindFirstDescendant(cf => cf.ByAutomationId(searchRadiusTextBox));
                        if (textBox4 == null)
                        {
                            OnEvent(".");
                        }
                    } while (textBox4 == null);

                    if (textBox4.ControlType != ControlType.Edit)
                        textBox4 = textBox4.AsTextBox().Parent;

                    textBox4?.AsTextBox().Enter(parameters["search-radius"]);

                    AutomationElement button6 = null;
                    do
                    {
                        button6 = window.FindFirstDescendant(cf => cf.ByText(exportBtnLabel));
                        if (button6 == null)
                        {
                            OnEvent(".");
                        }
                    } while (button6 == null);

                    if (button6.ControlType != ControlType.Button)
                        button6 = button6.AsButton().Parent;

                    button6?.AsButton().Invoke();
                    bool finished = false;
                    onEvent("composing.");
                    do
                    {
                        var window2 = app.GetMainWindow(automation);
                        var button7 = window.FindFirstDescendant(cf => cf.ByText(exportToDiskBtnLabel));

                        title = window2.Title;
                        finished = button7 != null && title.StartsWith("U");
                        int percent = 0;
                        if (!finished)
                        {
                            var percentStr = title.Substring(0, 2);
                            var numStr = percentStr[1] == '%' ? percentStr.Substring(0, 1) : percentStr;
                            if (int.TryParse(numStr, out percent))
                                onProgress?.Invoke(percent);
                        }
                    } while (!finished);
                }
                catch (Exception ex)
                {
                    OnEvent(ex.Message);
                }
                try
                {
                    var button7 = window.FindFirstDescendant(cf => cf.ByText(exportToDiskBtnLabel));
                    if (button7 != null && button7.ControlType != ControlType.Button)
                        button7 = button7.AsButton().Parent;
                    button7?.AsButton().Invoke();
                    OnEvent("exporting to disk...");
                }
                catch (Exception ex)
                {
                    OnEvent(ex.Message);
                }
                try
                {
                    var saveDlg = window.ModalWindows.FirstOrDefault(w => w.Name == exportPanoramaBtnLabel);

                    var pane40965 = saveDlg.FindFirstDescendant(cf => cf.ByAutomationId("40965"));
                    var pane41477 = pane40965.FindFirstDescendant(cf => cf.ByAutomationId("41477"));
                    var progressBar = pane41477.FindFirstDescendant(cf => cf.ByClassName("msctls_progress32"));
                    var breadcrumbParent = progressBar.FindFirstDescendant(cf => cf.ByClassName("Breadcrumb Parent"));
                    var toolbar = breadcrumbParent.FindFirstDescendant(cf => cf.ByClassName("ToolbarWindow32"));

                    toolbar.Click();
                    FlaUI.Core.Input.Keyboard.TypeSimultaneously(FlaUI.Core.WindowsAPI.VirtualKeyShort.CONTROL, FlaUI.Core.WindowsAPI.VirtualKeyShort.KEY_V);
                    FlaUI.Core.Input.Keyboard.Type(FlaUI.Core.WindowsAPI.VirtualKeyShort.ENTER);

                    var buttonSave = saveDlg.FindFirstDescendant(cf => cf.ByText(saveBtnLabel)).AsButton();
                    buttonSave?.DoubleClick();

                    Thread.Sleep(saveWait);
                }
                catch (Exception ex)
                {
                    OnEvent(ex.Message);
                }
            }
            app.Kill();
        }
        private void OnEvent(string message)
        {
            this.onEvent?.Invoke(message);
        }
    }
}