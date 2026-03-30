using EnvDTE;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;
using TCatSysManagerLib;

namespace TwinCATCppProjectCreator
{
    internal class Program
    {
        private static readonly List<string> SupportedVsDteVersions = new List<string>
        {
            "VisualStudio.DTE.15.0",
            "VisualStudio.DTE.16.0",
            "VisualStudio.DTE.17.0"
        };

        private static readonly Dictionary<int, (string TemplateName, string WizardId)> ModuleTemplates = new Dictionary<int, (string, string)>
        {
            {1, ("TwinCAT Module Class", "TcModuleClassWizard")},
            {2, ("TwinCAT Module Class with ADS port", "TcModuleClassWithAdsPortWizard")},
            {3, ("TwinCAT Module Class with cyclic caller", "TcModuleCyclicCallerWizard")},
            {4, ("TwinCAT Module Class with cyclic input/output", "TcModuleCyclicIoWizard")},
            {5, ("TwinCAT Module Class with data pointer", "TcModuleClassWithDataPointerWizard")},
            {6, ("TwinCAT Module Class for real-time context", "TcModuleClassForRealTimeContextWizard")},
            {7, ("TwinCAT Module Class with Online Changeable capability", "TcModuleClassWithOnlineChangeableWizard")}
        };

        private static readonly Dictionary<int, (string TemplateName, string WizardId)> ProjectTemplates = new Dictionary<int, (string, string)>
        {
            {1, ("Versioned C++ projects", "TcVersionedDriverWizard")},
            {2, ("Driver C++ project", "TcDriverWizard")}
        };

        private const string DEFAULT_CPP_PROJECT_NAME = "NewCppProject";
        private const string DEFAULT_MODULE_NAME = "NewModule";

        // 4024 / 4026 常见模块发布目录候选
        // TwinCAT 在不同版本/安装方式下，模块发布目录并不固定。
        // 这里按“常见安装目录 -> ProgramData 回退目录”的顺序尝试，
        // 这样 4024 / 4026 都可以共用同一套查找逻辑，而不是把路径写死后直接失败。
        private static readonly string[] DefaultTmcBasePaths =
        {
            @"C:\TwinCAT\3.1\CustomConfig\Modules",
            @"C:\TwinCAT\3.1\Config\Modules",
            @"C:\ProgramData\Beckhoff\TwinCAT\3.1\CustomConfig\Modules",
            @"C:\ProgramData\Beckhoff\TwinCAT\3.1\Config\Modules"
        };

        // 4026 仓库目录变化后，很多内容也会落到 ProgramData 下
        // 4026 经常把生成物放进 Repository，而不是旧版本更常见的 CustomConfig。
        // 后面查找 TMC、模块 GUID、发布结果时会同时扫描这些根目录，
        // 避免“文件其实已经生成，但程序因为找错目录而误判失败”。
        private static readonly string[] TwinCatRepositoryRoots =
        {
            @"C:\ProgramData\Beckhoff\TwinCAT\3.1\Repository",
            @"C:\TwinCAT\3.1\Repository"
        };

        private const int DEFAULT_TIMEOUT = 5000;
        private const int CHECK_INTERVAL = 100;
        private const int DEFAULT_FILE_WAIT_TIMEOUT = 120000;
        private const string VS_SOLUTION_FOLDER_KIND = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";

        // 下面这几个字段保存一次运行中的关键上下文。
        // 后续无论是重新查找项目树节点、定位生成文件、修补 TMC，还是发布/挂载模块，
        // 都依赖这些值来避免继续使用已经失效的 COM 引用或旧路径。
        private static string CurrentSolutionDirectory = null;
        private static string CurrentCppProjectName = DEFAULT_CPP_PROJECT_NAME;
        private static string CurrentModuleName = DEFAULT_MODULE_NAME;

        private const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001);
        private const int RPC_E_SERVERCALL_RETRYLATER = unchecked((int)0x8001010A);
        private const int RPC_S_CALL_FAILED = unchecked((int)0x800706BE);

        [STAThread]
        static void Main(string[] args)
        {
            // 主流程按“绑定 VS -> 创建项目 -> 创建模块 -> 生成/修补产物 -> 发布 -> 挂载实例”推进。
            // TwinCAT / Visual Studio 自动化接口的失败通常非常依赖时序，
            // 所以这里故意保留分阶段执行和阶段间保存/等待，便于稳定运行和定位问题。
            DTE selectedDte = null;
            ITcSysManager sysManager = null;
            ITcSmTreeItem cppProject = null;

            MessageFilter.Register();

            try
            {
                Console.WriteLine("========================================");
                Console.WriteLine("===== 绑定Visual Studio项目 =====");
                selectedDte = SelectAndBindVsProject();
                if (selectedDte == null)
                {
                    Console.WriteLine("× 你选择取消绑定，程序即将退出...");
                    return;
                }
                Console.WriteLine("√  项目绑定成功！");

                sysManager = ConnectToBoundTcProject(selectedDte);
                CurrentSolutionDirectory = GetSolutionDirectory(selectedDte);

                Console.WriteLine("\n========================================");
                Console.WriteLine("是否新建TwinCAT C++项目？(Y/N)（5秒无操作默认Y）：");
                string createCppProjInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{createCppProjInput}");

                if (createCppProjInput != "Y")
                {
                    Console.WriteLine("× 你选择不创建C++项目，程序即将退出...");
                    return;
                }

                Console.WriteLine("\n===== 选择C++项目模板 =====");
                foreach (var item in ProjectTemplates)
                {
                    Console.WriteLine($"{item.Key} : {item.Value.TemplateName}");
                }
                Console.WriteLine("请输入模板编号（4026 推荐 1，5秒无操作默认1）：");
                int selectedProjTemplateId = GetNumberInputWithTimeout(1, ProjectTemplates.Keys, DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择的模板编号：{selectedProjTemplateId}");

                var selectedProjTemplate = ProjectTemplates[selectedProjTemplateId];

                // 4026 更稳：先算一个唯一项目名，避免目录撞车
                CurrentCppProjectName = GetUniqueCppProjectName(CurrentSolutionDirectory, DEFAULT_CPP_PROJECT_NAME);
                Console.WriteLine($"→ 本次创建的 C++ 项目名：{CurrentCppProjectName}");

                cppProject = CreateTcCppProject(sysManager, selectedProjTemplate.WizardId, CurrentCppProjectName);
                Console.WriteLine($"√  {selectedProjTemplate.TemplateName}创建完成！");

                SaveAll(selectedDte);
                WaitForCppProjectStabilized(selectedDte, sysManager, CurrentCppProjectName);
                
                // 重新取一次树节点，避免 4026 下刚创建返回的旧句柄失效
                ReleaseComIfNeeded(cppProject);
                cppProject = GetCppProjectTreeItem(sysManager, CurrentCppProjectName);

                if (selectedProjTemplateId == 1)
                {
                    Console.WriteLine("===== 4026 Versioned C++ 模块处理 =====");
                    Console.WriteLine("→ 先检查 4026 的项目向导是否已经自带默认模块；如果没有，再回退到手动选择模块模板并创建模块。");
                    SaveAll(selectedDte);
                    WaitForProjectArtifacts(selectedDte);
                    System.Threading.Thread.Sleep(1500);

                    if (ProjectAlreadyContainsDefaultModuleSkeleton(selectedDte))
                    {
                        string detectedModuleName = TryGetExistingModuleNameFromProjectArtifacts(selectedDte);
                        CurrentModuleName = string.IsNullOrWhiteSpace(detectedModuleName) ? DEFAULT_MODULE_NAME : detectedModuleName;
                        Console.WriteLine($"√  检测到项目向导已自带默认模块：{CurrentModuleName}");
                    }
                    else
                    {
                        Console.WriteLine("→ 未检测到默认模块，回退到模块模板选择流程。");
                        Console.WriteLine("→ 诊断信息：" + GetVersionedProjectMissingDefaultModuleMessage(selectedDte));

                        var selectedTemplate = PromptForModuleTemplateSelection();
                        CurrentModuleName = GetUniqueModuleName(CurrentSolutionDirectory, CurrentCppProjectName, DEFAULT_MODULE_NAME);
                        Console.WriteLine($"→ 本次创建的模块名：{CurrentModuleName}");

                        CreateModuleWithFallbackForVersionedProject(selectedDte, sysManager, CurrentModuleName, selectedTemplate.TemplateId, selectedTemplate.TemplateName, selectedTemplate.WizardId);
                        Console.WriteLine($"√  已通过回退流程创建模块「{selectedTemplate.TemplateName}」。");

                        SaveAll(selectedDte);
                        WaitForProjectArtifacts(selectedDte);
                        System.Threading.Thread.Sleep(1500);
                    }
                }
                else
                {
                    var selectedTemplate = PromptForModuleTemplateSelection();

                    CurrentModuleName = GetUniqueModuleName(CurrentSolutionDirectory, CurrentCppProjectName, DEFAULT_MODULE_NAME);
                    Console.WriteLine($"→ 本次创建的模块名：{CurrentModuleName}");

                    CreateTcCppModuleStable(selectedDte, sysManager, CurrentCppProjectName, CurrentModuleName, selectedTemplate.WizardId);
                    ValidateModuleIntegratedIntoProjectModel(selectedDte, CurrentModuleName);
                    Console.WriteLine($"√  模块「{selectedTemplate.TemplateName}」创建操作已完成，并已进入工程模型。");

                    SaveAll(selectedDte);
                    WaitForProjectArtifacts(selectedDte);
                    System.Threading.Thread.Sleep(1500);
                }

                cppProject = GetCppProjectTreeItem(sysManager, CurrentCppProjectName);

                Console.WriteLine("\n========================================");
                Console.WriteLine("是否往生成的C++模块里写一点点简单代码？(Y/N)（5秒无操作默认Y）：");
                string writeCppInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{writeCppInput}");
                bool writeCppRequested = writeCppInput == "Y";
                bool simpleCodeWritten = false;

                if (writeCppRequested)
                {
                    simpleCodeWritten = WriteSimpleCodeToGeneratedModule(selectedDte);
                    if (simpleCodeWritten)
                    {
                        Console.WriteLine("√  已向生成的C++模块写入简单示例代码！");
                    }
                    else
                    {
                        Console.WriteLine("→ 当前源码里还没有可安全注入的自动生成函数体，稍后在 Code Generator 之后再尝试一次。");
                    }
                }
                else
                {
                    Console.WriteLine("× 跳过自动写入模块源码...");
                }

                Console.WriteLine("\n========================================");
                Console.WriteLine("是否启用TMC Code Generator？(Y/N)（5秒无操作默认Y）：");
                string tmcGenerateInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{tmcGenerateInput}");

                if (tmcGenerateInput == "Y")
                {
                    cppProject = GetCppProjectTreeItem(sysManager, CurrentCppProjectName);
                    ExecuteTmcCodeGenerator(cppProject);
                    DumpTmcCandidates(selectedDte);
                    SaveAll(selectedDte);
                    System.Threading.Thread.Sleep(3000);
                    Console.WriteLine("√  TMC Code Generator执行成功！");

                    if (writeCppRequested && !simpleCodeWritten)
                    {
                        simpleCodeWritten = WriteSimpleCodeToGeneratedModule(selectedDte);
                        if (simpleCodeWritten)
                        {
                            Console.WriteLine("√  已在 Code Generator 完成后写入简单示例代码。");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("× 跳过TMC Code Generator...");
                }

                Console.WriteLine("\n========================================");
                Console.WriteLine("是否修改工程TMC（保守方式：修改向导默认结构体参数为 Gain / Enable / VelocityLimit）？(Y/N)（5秒无操作默认Y）：");
                string patchTmcInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{patchTmcInput}");

                if (patchTmcInput == "Y")
                {
                    if (TryPatchProjectTmcWithRecovery(selectedDte, sysManager, out string patchedTmcPath))
                    {
                        Console.WriteLine($"√  工程TMC修改成功！→ {patchedTmcPath}");
                    }
                    else
                    {
                        Console.WriteLine("× 未能自动修改工程TMC，已继续后续流程。可在发布后再次执行本工具重试。");
                    }
                }
                else
                {
                    Console.WriteLine("× 跳过修改工程TMC...");
                }

                Console.WriteLine("\n========================================");
                Console.WriteLine("是否发布TcCOM Objects？(Y/N)（5秒无操作默认Y）：");
                string publishModulesInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{publishModulesInput}");

                if (publishModulesInput == "Y")
                {
                    cppProject = GetCppProjectTreeItem(sysManager, CurrentCppProjectName);
                    if (ExecutePublishModules(cppProject))
                    {
                        Console.WriteLine("√  TcCOM Modules发布成功！");
                    }
                    else
                    {
                        Console.WriteLine("× TcCOM Modules未检测到有效发布产物，后续不会继续添加 TcCOM Object。");
                    }
                }
                else
                {
                    Console.WriteLine("× 跳过发布TcCOM Objects...");
                }

                Console.WriteLine("\n========================================");
                Console.WriteLine("是否添加TcCOM Object到配置？(Y/N)（5秒无操作默认Y）：");
                string addTcComInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{addTcComInput}");

                bool tcComAdded = false;
                if (addTcComInput == "Y")
                {
                    cppProject = GetCppProjectTreeItem(sysManager, CurrentCppProjectName);
                    tcComAdded = AddTcComObject(sysManager, cppProject);
                    if (tcComAdded)
                    {
                        Console.WriteLine("√  TcCOM Object添加成功！");
                    }
                    else
                    {
                        Console.WriteLine("× 未添加 TcCOM Object：未找到可用的已发布 TMC，或模块尚未成功发布。");
                    }
                }
                else
                {
                    Console.WriteLine("× 跳过添加TcCOM Object...");
                }

                Console.WriteLine("\n========================================");
                Console.WriteLine("是否编译当前项目？(Y/N)（5秒无操作默认Y）：");
                string buildProjectInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{buildProjectInput}");

                bool buildSucceeded = false;
                if (buildProjectInput == "Y")
                {
                    buildSucceeded = BuildCurrentSolution(selectedDte);
                    if (buildSucceeded)
                    {
                        Console.WriteLine("√  当前项目编译完成！");
                    }
                    else
                    {
                        Console.WriteLine("× 当前项目编译失败，请查看 Visual Studio 错误列表。");
                    }
                }
                else
                {
                    Console.WriteLine("× 跳过编译当前项目...");
                }

                Console.WriteLine("\n========================================");
                Console.WriteLine("是否激活 TwinCAT 配置（Active Configuration）？(Y/N)（5秒无操作默认Y）：");
                string activeConfigInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{activeConfigInput}");

                if (activeConfigInput == "Y")
                {
                    if (buildProjectInput == "Y" && !buildSucceeded)
                    {
                        Console.WriteLine("× 跳过激活 TwinCAT 配置：当前解决方案编译未通过。");
                    }
                    else
                    {
                        ActivateTcConfiguration(selectedDte, sysManager);
                    }
                }
                else
                {
                    Console.WriteLine("× 跳过激活 TwinCAT 配置...");
                }

                Console.WriteLine("\n========================================");
                Console.WriteLine("√  所有操作执行完成！");
            }
            catch (COMException ex)
            {
                Console.WriteLine($"\n× TwinCAT COM错误 → {ex.Message}（错误码：0x{ex.ErrorCode:X8}）");
            }
            catch (InvalidCastException)
            {
                Console.WriteLine("\n× 错误：当前打开的项目不是TwinCAT项目！");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n× 运行错误 → {ex.Message}");
            }
            finally
            {
                MessageFilter.Revoke();

                ReleaseComIfNeeded(cppProject);
                ReleaseComIfNeeded(sysManager);
                ReleaseComIfNeeded(selectedDte);
            }

            Console.WriteLine("\n按任意键关闭窗口...");
            Console.ReadKey();
        }

        #region 输入处理
        private static string GetYesNoInputWithTimeout(string defaultValue, int timeoutMs)
        {
            int remainingTime = timeoutMs;
            while (remainingTime > 0)
            {
                if (Console.KeyAvailable)
                {
                    ConsoleKeyInfo keyInfo = Console.ReadKey(true);
                    char keyChar = char.ToUpperInvariant(keyInfo.KeyChar);
                    ClearConsoleInputBuffer();
                    if (keyChar == 'Y' || keyChar == 'N')
                    {
                        return keyChar.ToString();
                    }
                    return defaultValue;
                }
                System.Threading.Thread.Sleep(CHECK_INTERVAL);
                remainingTime -= CHECK_INTERVAL;
            }
            return defaultValue;
        }

        private static (int TemplateId, string TemplateName, string WizardId) PromptForModuleTemplateSelection()
        {
            Console.WriteLine("\n========================================");
            Console.WriteLine("===== 选择C++模块模板 =====");
            foreach (var item in ModuleTemplates)
            {
                Console.WriteLine($"{item.Key} : {item.Value.TemplateName}");
            }
            Console.WriteLine("请输入模板编号（4026 建议先用 1，5秒无操作默认1）：");
            int selectedTemplateId = GetNumberInputWithTimeout(1, ModuleTemplates.Keys, DEFAULT_TIMEOUT);
            Console.WriteLine($"→ 选择的模板编号：{selectedTemplateId}");
            var template = ModuleTemplates[selectedTemplateId];
            return (selectedTemplateId, template.TemplateName, template.WizardId);
        }

        private static int GetNumberInputWithTimeout(int defaultValue, IEnumerable<int> validValues, int timeoutMs)
        {
            int remainingTime = timeoutMs;
            while (remainingTime > 0)
            {
                if (Console.KeyAvailable)
                {
                    ConsoleKeyInfo keyInfo = Console.ReadKey(true);
                    ClearConsoleInputBuffer();
                    if (int.TryParse(keyInfo.KeyChar.ToString(), out int selectedValue) && validValues.Contains(selectedValue))
                    {
                        return selectedValue;
                    }
                    Console.WriteLine($"× 输入无效，使用默认值：{defaultValue}");
                    return defaultValue;
                }
                System.Threading.Thread.Sleep(CHECK_INTERVAL);
                remainingTime -= CHECK_INTERVAL;
            }
            return defaultValue;
        }

        private static void ClearConsoleInputBuffer()
        {
            while (Console.KeyAvailable)
            {
                Console.ReadKey(true);
            }
        }
        #endregion

        #region 核心方法

        private static DTE SelectAndBindVsProject()
        {
            var dteInstances = GetRunningVsInstances();
            if (dteInstances.Count == 0)
            {
                throw new Exception("当前没有运行的Visual Studio实例！");
            }

            Console.WriteLine("请选择项目编号（0=取消）：");
            foreach (var item in dteInstances)
            {
                string projectName = RetryComCall(() =>
                {
                    if (item.Value.Solution != null && item.Value.Solution.Projects.Count > 0)
                    {
                        return item.Value.Solution.Projects.Item(1).Name;
                    }
                    return "未加载项目";
                });
                Console.WriteLine($"{item.Key} : {projectName} (Visual Studio实例)");
            }

            Console.Write("请输入编号：");
            string input = Console.ReadLine()?.Trim();
            if (input == "0") return null;
            if (!int.TryParse(input, out int selectedIndex) || !dteInstances.ContainsKey(selectedIndex))
            {
                throw new Exception("输入的编号无效！");
            }

            return dteInstances[selectedIndex];
        }

        private static Dictionary<int, DTE> GetRunningVsInstances()
        {
            var dteInstances = new Dictionary<int, DTE>();
            IRunningObjectTable rot;
            IEnumMoniker enumMoniker;
            IMoniker[] monikers = new IMoniker[1];
            int index = 1;

            if (GetRunningObjectTable(0, out rot) != 0)
                return dteInstances;

            rot.EnumRunning(out enumMoniker);
            enumMoniker.Reset();

            IntPtr fetchedPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(int)));
            try
            {
                while (enumMoniker.Next(1, monikers, fetchedPtr) == 0)
                {
                    int instanceCount = Marshal.ReadInt32(fetchedPtr);
                    if (instanceCount != 1) break;

                    IBindCtx bindCtx;
                    CreateBindCtx(0, out bindCtx);
                    string displayName;
                    monikers[0].GetDisplayName(bindCtx, null, out displayName);

                    bool isSupportedVsInstance = SupportedVsDteVersions.Any(version => displayName.Contains(version));
                    if (isSupportedVsInstance)
                    {
                        object obj;
                        rot.GetObject(monikers[0], out obj);
                        if (obj is DTE dte)
                        {
                            dteInstances.Add(index++, dte);
                        }
                    }

                    Marshal.ReleaseComObject(monikers[0]);
                    Marshal.ReleaseComObject(bindCtx);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(fetchedPtr);
            }

            Marshal.ReleaseComObject(enumMoniker);
            Marshal.ReleaseComObject(rot);
            return dteInstances;
        }

        private static ITcSysManager ConnectToBoundTcProject(DTE dte)
        {
            Console.WriteLine($"√  已连接到Visual Studio实例：{RetryComCall(() => dte.Name)} v{RetryComCall(() => dte.Version)}");

            if (RetryComCall(() => dte.Solution) == null || RetryComCall(() => dte.Solution.Projects.Count) == 0)
            {
                throw new Exception("当前实例未加载任何项目！");
            }

            Project tcProject = RetryComCall(() => dte.Solution.Projects.Item(1));
            Console.WriteLine($"√  已定位到TwinCAT项目：{tcProject.Name}");

            return (ITcSysManager)tcProject.Object;
        }

        private static bool BuildCurrentSolution(DTE dte)
        {
            if (dte?.Solution == null)
            {
                throw new Exception("当前未加载解决方案！");
            }

            SolutionBuild sb = RetryComCall(() => dte.Solution.SolutionBuild);
            RetryComCall(() => sb.Build(true));

            const int buildWaitTimeoutMs = 300000;
            int waited = 0;
            while (waited < buildWaitTimeoutMs)
            {
                if (RetryComCall(() => sb.BuildState) != vsBuildState.vsBuildStateInProgress)
                    break;
                System.Threading.Thread.Sleep(500);
                waited += 500;
            }

            bool buildOk = RetryComCall(() => sb.LastBuildInfo) == 0 &&
                           RetryComCall(() => sb.BuildState) == vsBuildState.vsBuildStateDone;

            if (buildOk)
            {
                Console.WriteLine("→ 编译成功。");
            }
            else
            {
                Console.WriteLine("→ 编译已结束，若失败请查看 Visual Studio 错误列表（Error List）。");
            }

            return buildOk;
        }

        private static void ActivateTcConfiguration(DTE dte, ITcSysManager sysManager)
        {
            if (dte?.Solution == null)
            {
                throw new Exception("当前未加载解决方案，无法激活配置。");
            }

            string solutionDir = Path.GetDirectoryName(dte.Solution.FullName);
            string configPath = Path.Combine(solutionDir ?? "", "CurrentConfig.tszip");

            try
            {
                Console.WriteLine("→ 正在保存当前 TwinCAT 配置...");
                sysManager.SaveConfiguration(configPath);
                Console.WriteLine($"→ 配置已保存：{configPath}");
            }
            catch (Exception ex)
            {
                if (!string.IsNullOrEmpty(ex.Message) &&
                    ex.Message.IndexOf("Automation Legacy Mode", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("→ 当前环境不支持通过 SysManager.SaveConfiguration 保存配置，改为继续使用标准激活流程。");
                }
                else
                {
                    Console.WriteLine($"× 保存配置时提示：{ex.Message}（继续尝试激活）");
                }
            }

            bool activated = false;
            string[] tryCommands =
            {
                "TwinCAT.ActivateConfiguration",
                "TcXaeShell.TwinCAT.ActivateConfiguration",
                "Build.ActivateConfiguration"
            };

            foreach (string cmd in tryCommands)
            {
                try
                {
                    RetryComCall(() => dte.ExecuteCommand(cmd));
                    activated = true;
                    Console.WriteLine($"√  已通过命令执行激活：{cmd}");
                    break;
                }
                catch
                {
                }
            }

            if (!activated)
            {
                try
                {
                    Console.WriteLine("→ 正在通过 System Manager 激活配置...");
                    RetryComCall(() => sysManager.ActivateConfiguration());
                    activated = true;
                    Console.WriteLine("√  TwinCAT 配置已激活。");
                }
                catch (Exception ex)
                {
                    throw new Exception($"激活 TwinCAT 配置失败：{ex.Message}");
                }
            }

            if (activated)
            {
                Console.WriteLine("→ 若 TwinCAT 未切到 Run，请在 IDE 中确认或手动点击“激活配置”。");
            }
        }

        private static ITcSmTreeItem CreateTcCppProject(ITcSysManager sysManager, string templateId, string projectName)
        {
            ITcSmTreeItem cppNode = RetryComCall(() => sysManager.LookupTreeItem("TIXC"));
            if (cppNode == null)
            {
                throw new Exception("未找到TwinCAT C++节点（TIXC），请确认项目支持C++开发！");
            }

            try
            {
                ITcSmTreeItem cppProject = RetryComCall(() => cppNode.CreateChild(projectName, 0, "", templateId));
                return cppProject;
            }
            catch (Exception ex)
            {
                throw new Exception($"创建 C++ 项目失败（模板：{templateId}，项目名：{projectName}）：{ex.Message}");
            }
        }

        private static bool ProjectAlreadyContainsDefaultModuleSkeleton(DTE dte)
        {
            string solutionDir = GetSolutionDirectory(dte);
            string projectDir = Path.Combine(solutionDir, CurrentCppProjectName);

            if (!Directory.Exists(projectDir))
            {
                return false;
            }

            // 1) 以 ClassMap 中是否已有实体映射为准
            string classFactoryCpp = Path.Combine(projectDir, CurrentCppProjectName + "ClassFactory.cpp");
            if (HasNonEmptyClassMap(classFactoryCpp))
            {
                return true;
            }

            // 2) 再看 tmc 是否已有真实 Module 记录（而不是 <Modules/> 空壳）
            foreach (string tmcPath in SafeEnumerateFiles(projectDir, "*.tmc"))
            {
                try
                {
                    XDocument doc = XDocument.Load(tmcPath);
                    bool hasModuleEntries = doc.Descendants().Any(x =>
                        x.Name.LocalName == "Modules" &&
                        x.Elements().Any(e => e.Name.LocalName == "Module"));

                    if (hasModuleEntries)
                    {
                        return true;
                    }
                }
                catch
                {
                }
            }

            return false;
        }

        private static string GetVersionedProjectMissingDefaultModuleMessage(DTE dte)
        {
            string solutionDir = GetSolutionDirectory(dte);
            string projectDir = Path.Combine(solutionDir, CurrentCppProjectName);
            string vcxprojPath = Path.Combine(projectDir, CurrentCppProjectName + ".vcxproj");
            string tmcPath = Path.Combine(projectDir, CurrentCppProjectName + ".tmc");

            bool projectDirExists = Directory.Exists(projectDir);
            bool vcxprojExists = File.Exists(vcxprojPath);
            bool tmcExists = File.Exists(tmcPath);
            bool vcxprojContainsModuleCpp = false;
            bool tmcContainsModule = false;

            try
            {
                if (vcxprojExists)
                {
                    string vcxprojText = File.ReadAllText(vcxprojPath);
                    vcxprojContainsModuleCpp =
                        Regex.IsMatch(vcxprojText, @"<ClCompile Include=""(?!TcPch\.cpp)(?!.*ClassFactory)(?!.*Driver)(?!.*Main)[^""]+\.cpp""", RegexOptions.IgnoreCase);
                }
            }
            catch
            {
            }

            try
            {
                if (tmcExists)
                {
                    XDocument doc = XDocument.Load(tmcPath);
                    tmcContainsModule = doc.Descendants().Any(x =>
                        x.Name.LocalName == "Modules" &&
                        x.Elements().Any(e => e.Name.LocalName == "Module"));
                }
            }
            catch
            {
            }

            string visibleFiles = projectDirExists
                ? string.Join(", ", SafeEnumerateFiles(projectDir, "*.*")
                    .Where(IsProjectSourcePath)
                    .Select(Path.GetFileName)
                    .OrderBy(x => x)
                    .Take(20))
                : "(项目目录不存在)";

            return
                "Versioned C++ 项目创建完成后，未检测到默认模块骨架。" +
                $" 已检查：项目目录={projectDirExists}，vcxproj={vcxprojExists}，tmc={tmcExists}，vcxproj含模块cpp={vcxprojContainsModuleCpp}，tmc含Module节点={tmcContainsModule}。" +
                $" 当前工程目录可见文件：{visibleFiles}。" +
                " 这说明本机 4026 的项目向导只创建了项目外壳，没有把默认模块真正创建出来，当前不再继续后续流程。";
        }

        private static string TryGetExistingModuleNameFromProjectTmc(DTE dte)
        {
            string solutionDir = GetSolutionDirectory(dte);
            string projectDir = Path.Combine(solutionDir, CurrentCppProjectName);

            if (!Directory.Exists(projectDir))
            {
                return null;
            }

            foreach (string tmcPath in SafeEnumerateFiles(projectDir, "*.tmc").Where(IsProjectSourcePath).OrderBy(p => p.Length))
            {
                try
                {
                    XDocument doc = XDocument.Load(tmcPath);
                    XElement moduleElement = doc.Descendants().FirstOrDefault(x => x.Name.LocalName == "Module");
                    if (moduleElement == null)
                    {
                        continue;
                    }

                    string moduleName = GetChildElementValue(moduleElement, "Name");
                    if (!string.IsNullOrWhiteSpace(moduleName))
                    {
                        return moduleName.Trim();
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private static string TryGetExistingModuleNameFromProjectArtifacts(DTE dte)
        {
            string moduleNameFromTmc = TryGetExistingModuleNameFromProjectTmc(dte);
            if (!string.IsNullOrWhiteSpace(moduleNameFromTmc))
            {
                return moduleNameFromTmc;
            }

            string solutionDir = GetSolutionDirectory(dte);
            string projectDir = Path.Combine(solutionDir, CurrentCppProjectName);
            if (!Directory.Exists(projectDir))
            {
                return null;
            }

            string vcxprojPath = Path.Combine(projectDir, CurrentCppProjectName + ".vcxproj");
            if (File.Exists(vcxprojPath))
            {
                try
                {
                    XDocument doc = XDocument.Load(vcxprojPath);
                    XNamespace ns = doc.Root?.Name.Namespace ?? XNamespace.None;
                    string moduleCpp = doc.Descendants(ns + "ClCompile")
                        .Select(x => x.Attribute("Include")?.Value)
                        .FirstOrDefault(value =>
                            !string.IsNullOrWhiteSpace(value) &&
                            !value.EndsWith("TcPch.cpp", StringComparison.OrdinalIgnoreCase) &&
                            value.IndexOf("ClassFactory", StringComparison.OrdinalIgnoreCase) < 0 &&
                            value.IndexOf("Driver", StringComparison.OrdinalIgnoreCase) < 0 &&
                            value.IndexOf("Main", StringComparison.OrdinalIgnoreCase) < 0);

                    if (!string.IsNullOrWhiteSpace(moduleCpp))
                    {
                        return Path.GetFileNameWithoutExtension(moduleCpp);
                    }
                }
                catch
                {
                }
            }

            string fallbackCpp = SafeEnumerateFiles(projectDir, "*.cpp")
                .Where(IsProjectSourcePath)
                .Select(Path.GetFileName)
                .FirstOrDefault(name =>
                    !string.IsNullOrWhiteSpace(name) &&
                    !name.Equals("TcPch.cpp", StringComparison.OrdinalIgnoreCase) &&
                    name.IndexOf("ClassFactory", StringComparison.OrdinalIgnoreCase) < 0 &&
                    name.IndexOf("Driver", StringComparison.OrdinalIgnoreCase) < 0 &&
                    name.IndexOf("Main", StringComparison.OrdinalIgnoreCase) < 0);

            return string.IsNullOrWhiteSpace(fallbackCpp) ? null : Path.GetFileNameWithoutExtension(fallbackCpp);
        }

        private static bool HasNonEmptyClassMap(string classFactoryCppPath)
        {
            if (string.IsNullOrWhiteSpace(classFactoryCppPath) || !File.Exists(classFactoryCppPath))
            {
                return false;
            }

            string[] lines = File.ReadAllLines(classFactoryCppPath);
            int start = Array.FindIndex(lines, line => line.IndexOf("///<AutoGeneratedContent id=\"ClassMap\">", StringComparison.OrdinalIgnoreCase) >= 0);
            int end = Array.FindIndex(lines, line => line.IndexOf("///</AutoGeneratedContent>", StringComparison.OrdinalIgnoreCase) >= 0);

            if (start < 0 || end < 0 || end <= start)
            {
                return false;
            }

            for (int i = start + 1; i < end; i++)
            {
                string line = lines[i].Trim();
                if (line.Length == 0) continue;
                if (line.StartsWith("///", StringComparison.Ordinal)) continue;
                return true;
            }

            return false;
        }

        private sealed class BootstrapModuleProfile
        {
            public int TemplateId { get; set; }
            public string TemplateName { get; set; }
            public bool IncludeAdsPort { get; set; }
            public bool IncludeCyclicIoHints { get; set; }
            public bool IncludeDataPointerHint { get; set; }
            public bool IncludeRealtimeContextHint { get; set; }
            public bool IncludeOnlineChangeHint { get; set; }
        }

        private static BootstrapModuleProfile GetBootstrapModuleProfile(int templateId, string templateName)
        {
            return new BootstrapModuleProfile
            {
                TemplateId = templateId,
                TemplateName = string.IsNullOrWhiteSpace(templateName) ? "TwinCAT Module Class" : templateName,
                IncludeAdsPort = templateId == 2,
                IncludeCyclicIoHints = templateId == 4,
                IncludeDataPointerHint = templateId == 5,
                IncludeRealtimeContextHint = templateId == 6,
                IncludeOnlineChangeHint = templateId == 7
            };
        }

        // 4026 的 Versioned C++ project 上，模块向导并不总是稳定。
        // 因此这里先走正常 wizard；如果失败，再退回到“手动补工程骨架”的恢复流程，
        // 尽量把“向导失败”转换成“仍然得到一个后续可生成、可发布的模块项目”。
        private static void CreateModuleWithFallbackForVersionedProject(
            DTE dte,
            ITcSysManager sysManager,
            string moduleName,
            int selectedTemplateId,
            string selectedTemplateName,
            string wizardId)
        {
            try
            {
                CreateTcCppModuleStable(dte, sysManager, CurrentCppProjectName, moduleName, wizardId);
                ValidateModuleIntegratedIntoProjectModel(dte, moduleName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"→ TwinCAT 模块向导不可用，改为手工生成兼容骨架：{ex.Message}");

                if (!TryBootstrapModuleSkeletonForVersionedProject(dte, moduleName, selectedTemplateId, selectedTemplateName, out string bootstrapMessage))
                {
                    throw;
                }

                Console.WriteLine($"√  {bootstrapMessage}");
                ValidateModuleIntegratedIntoProjectModel(dte, moduleName);
            }
        }

        // 这个回退流程直接在磁盘层面补模块需要的关键文件和工程项：
        // .h / .cpp、Services、ClassFactory、TMC、vcxproj 以及 filters。
        // 它的目标不是完全复制 TwinCAT wizard 的所有细节，而是补出一个足够完整、
        // 能让后续代码生成、发布和 TcCOM 挂载继续跑下去的“最小可用模块骨架”。
        private static bool TryBootstrapModuleSkeletonForVersionedProject(
            DTE dte,
            string moduleName,
            int selectedTemplateId,
            string selectedTemplateName,
            out string resultMessage)
        {
            resultMessage = null;

            try
            {
                string solutionDir = GetSolutionDirectory(dte);
                string projectDir = Path.Combine(solutionDir, CurrentCppProjectName);
                if (!Directory.Exists(projectDir))
                {
                    return false;
                }

                string servicesPath = Path.Combine(projectDir, CurrentCppProjectName + "Services.h");
                string classFactoryPath = Path.Combine(projectDir, CurrentCppProjectName + "ClassFactory.cpp");
                string tmcPath = Path.Combine(projectDir, CurrentCppProjectName + ".tmc");
                string vcxprojPath = Path.Combine(projectDir, CurrentCppProjectName + ".vcxproj");
                string moduleHeaderPath = Path.Combine(projectDir, moduleName + ".h");
                string moduleCppPath = Path.Combine(projectDir, moduleName + ".cpp");

                if (!File.Exists(servicesPath) || !File.Exists(classFactoryPath) || !File.Exists(tmcPath) || !File.Exists(vcxprojPath))
                {
                    return false;
                }

                BootstrapModuleProfile profile = GetBootstrapModuleProfile(selectedTemplateId, selectedTemplateName);

                string safeProjectIdentifier = MakeSafeIdentifier(CurrentCppProjectName);
                string safeModuleIdentifier = MakeSafeIdentifier(moduleName);
                string moduleClassName = "C" + safeModuleIdentifier;
                string classIdName = "CID_" + safeProjectIdentifier + moduleClassName;
                Guid moduleGuid = Guid.NewGuid();

                string srvNameMacro = GetServiceNameMacroToken(servicesPath);

                WriteBootstrapModuleHeader(moduleHeaderPath, moduleName, moduleClassName, classIdName, safeModuleIdentifier, profile);
                WriteBootstrapModuleCpp(moduleCppPath, moduleName, moduleClassName, safeModuleIdentifier, profile);

                UpsertClassIdInServicesHeader(servicesPath, classIdName, moduleGuid);
                UpsertClassFactoryMapEntry(classFactoryPath, moduleName, moduleClassName, classIdName, srvNameMacro);
                UpsertModuleDefinitionInTmc(tmcPath, moduleName, CurrentCppProjectName, moduleGuid.ToString("B").ToUpperInvariant(), safeModuleIdentifier, profile);
                AddModuleFilesToProjectFiles(vcxprojPath, moduleName);
                string filtersPath = vcxprojPath + ".filters";
                if (File.Exists(filtersPath))
                {
                    AddModuleFilesToVcxprojFilters(filtersPath, moduleName);
                }

                SaveAll(dte);
                System.Threading.Thread.Sleep(1000);

                resultMessage = $"按模板「{profile.TemplateName}」生成 {moduleName}.h/.cpp，并已写入 ClassFactory / Services / TMC / 项目文件";
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"× 自动补齐模块骨架失败：{ex.Message}");
                return false;
            }
        }

        private static string MakeSafeIdentifier(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return "X";
            }

            char[] chars = rawValue
                .Select(ch => char.IsLetterOrDigit(ch) || ch == '_' ? ch : '_')
                .ToArray();
            string identifier = new string(chars);
            if (identifier.Length == 0)
            {
                identifier = "X";
            }

            if (char.IsDigit(identifier[0]))
            {
                identifier = "_" + identifier;
            }

            return identifier;
        }

        private static string GetServiceNameMacroToken(string servicesHeaderPath)
        {
            string content = File.ReadAllText(servicesHeaderPath);
            Match match = Regex.Match(content, "#define\\s+(SRVNAME_[A-Za-z0-9_]+)\\s+\\\"");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            return "SRVNAME_" + MakeSafeIdentifier(CurrentCppProjectName).ToUpperInvariant();
        }

        private static string FormatGuidAsCtcid(Guid guid)
        {
            string[] parts = guid.ToString("D").Split('-');
            if (parts.Length != 5)
            {
                throw new Exception("无法将GUID转换为CTCID格式。");
            }

            string d1 = parts[0].ToLowerInvariant();
            string d2 = parts[1].ToLowerInvariant();
            string d3 = parts[2].ToLowerInvariant();
            string d4 = parts[3].ToLowerInvariant();
            string d5 = parts[4].ToLowerInvariant();

            var bytes = new List<string>
            {
                "0x" + d4.Substring(0, 2),
                "0x" + d4.Substring(2, 2),
                "0x" + d5.Substring(0, 2),
                "0x" + d5.Substring(2, 2),
                "0x" + d5.Substring(4, 2),
                "0x" + d5.Substring(6, 2),
                "0x" + d5.Substring(8, 2),
                "0x" + d5.Substring(10, 2)
            };

            return "{0x" + d1 + ",0x" + d2 + ",0x" + d3 + ",{" + string.Join(",", bytes) + "}}";
        }

        private static void UpsertClassIdInServicesHeader(string servicesHeaderPath, string classIdName, Guid moduleGuid)
        {
            List<string> lines = File.ReadAllLines(servicesHeaderPath).ToList();
            if (lines.Any(line => line.IndexOf(classIdName, StringComparison.OrdinalIgnoreCase) >= 0))
            {
                return;
            }

            int start = lines.FindIndex(line => line.IndexOf("///<AutoGeneratedContent id=\"ClassIDs\">", StringComparison.OrdinalIgnoreCase) >= 0);
            int end = lines.FindIndex(start + 1, line => line.IndexOf("///</AutoGeneratedContent>", StringComparison.OrdinalIgnoreCase) >= 0);

            string classIdLine = "const CTCID " + classIdName + " = " + FormatGuidAsCtcid(moduleGuid) + ";";

            if (start >= 0 && end > start)
            {
                lines.Insert(end, classIdLine);
            }
            else
            {
                lines.Add("///<AutoGeneratedContent id=\"ClassIDs\">");
                lines.Add(classIdLine);
                lines.Add("///</AutoGeneratedContent>");
            }

            File.WriteAllLines(servicesHeaderPath, lines);
        }

        private static void UpsertClassFactoryMapEntry(string classFactoryPath, string moduleName, string moduleClassName, string classIdName, string srvNameMacro)
        {
            List<string> lines = File.ReadAllLines(classFactoryPath).ToList();

            string includeLine = "#include \"" + moduleName + ".h\"";
            if (!lines.Any(line => string.Equals(line.Trim(), includeLine, StringComparison.OrdinalIgnoreCase)))
            {
                int lastInclude = lines.FindLastIndex(line => line.TrimStart().StartsWith("#include ", StringComparison.Ordinal));
                if (lastInclude >= 0)
                {
                    lines.Insert(lastInclude + 1, includeLine);
                }
                else
                {
                    lines.Insert(0, includeLine);
                }
            }

            if (!lines.Any(line => line.IndexOf(classIdName, StringComparison.OrdinalIgnoreCase) >= 0))
            {
                int start = lines.FindIndex(line => line.IndexOf("///<AutoGeneratedContent id=\"ClassMap\">", StringComparison.OrdinalIgnoreCase) >= 0);
                int end = lines.FindIndex(start + 1, line => line.IndexOf("///</AutoGeneratedContent>", StringComparison.OrdinalIgnoreCase) >= 0);

                if (start >= 0 && end > start)
                {
                    string entryLine = "\tCLASS_ENTRY_LIB(VID_" + MakeSafeIdentifier(CurrentCppProjectName) + ", " + classIdName + ", " + srvNameMacro + " \"!" + moduleClassName + "\", " + moduleClassName + ")";
                    lines.Insert(end, entryLine);
                }
            }

            File.WriteAllLines(classFactoryPath, lines);
        }

        private static void WriteBootstrapModuleHeader(
            string moduleHeaderPath,
            string moduleName,
            string moduleClassName,
            string classIdName,
            string safeModuleIdentifier,
            BootstrapModuleProfile profile)
        {
            if (File.Exists(moduleHeaderPath))
            {
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("///////////////////////////////////////////////////////////////////////////////");
            sb.AppendLine("// " + moduleName + ".h");
            sb.AppendLine("#pragma once");
            sb.AppendLine();
            sb.AppendLine("#include \"" + CurrentCppProjectName + "Interfaces.h\"");
            sb.AppendLine();
            sb.AppendLine("// AUTO_BOOTSTRAP_TEMPLATE: " + profile.TemplateName);

            if (profile.IncludeAdsPort)
            {
                sb.AppendLine("const PTCID PID_" + safeModuleIdentifier + "_DefaultAdsPort = 0x00000001;");
                sb.AppendLine("const PTCID PID_" + safeModuleIdentifier + "_AdsPort = 0x00000002;");
                sb.AppendLine();
            }

            sb.AppendLine("class " + moduleClassName);
            sb.AppendLine("\t: public ITComObject");
            sb.AppendLine("\t, public ITcADI");
            sb.AppendLine("\t, public ITcWatchSource");
            sb.AppendLine("///<AutoGeneratedContent id=\"InheritanceList\">");
            sb.AppendLine("///</AutoGeneratedContent>");
            sb.AppendLine("{");
            sb.AppendLine("public:");
            sb.AppendLine("\tDECLARE_IUNKNOWN()");
            sb.AppendLine("\tDECLARE_IPERSIST(" + classIdName + ")");
            sb.AppendLine("\tDECLARE_ITCOMOBJECT_LOCKOP()");
            sb.AppendLine("\tDECLARE_ITCADI()");
            sb.AppendLine("\tDECLARE_ITCWATCHSOURCE()");
            sb.AppendLine("\tDECLARE_OBJPARAWATCH_MAP()");
            sb.AppendLine("\tDECLARE_OBJDATAAREA_MAP()");
            sb.AppendLine();
            sb.AppendLine("\t" + moduleClassName + "();");
            sb.AppendLine("\tvirtual ~" + moduleClassName + "();");
            sb.AppendLine();
            sb.AppendLine("///<AutoGeneratedContent id=\"InterfaceMembers\">");
            sb.AppendLine("///</AutoGeneratedContent>");
            sb.AppendLine();
            sb.AppendLine("protected:");
            sb.AppendLine("\tDECLARE_ITCOMOBJECT_SETSTATE();");
            sb.AppendLine();
            sb.AppendLine("\tHRESULT AddModuleToCaller();");
            sb.AppendLine("\tVOID RemoveModuleFromCaller();");
            sb.AppendLine();
            sb.AppendLine("private:");
            sb.AppendLine("\tCTcTrace m_Trace;");
            sb.AppendLine("\tUINT m_counter;");

            if (profile.IncludeAdsPort)
            {
                sb.AppendLine("\tWORD m_DefaultAdsPort;");
                sb.AppendLine("\tWORD m_ContextAdsPort;");
            }

            if (profile.IncludeCyclicIoHints)
            {
                sb.AppendLine("\tLREAL m_InputValue;");
                sb.AppendLine("\tLREAL m_OutputValue;");
            }

            if (profile.IncludeDataPointerHint)
            {
                sb.AppendLine("\tPVOID m_pDataPointer;");
            }

            if (profile.IncludeOnlineChangeHint)
            {
                sb.AppendLine("\tBOOL m_EnableOnlineChange;");
            }

            sb.AppendLine("///<AutoGeneratedContent id=\"Members\">");
            sb.AppendLine("///</AutoGeneratedContent>");
            sb.AppendLine("};");
            File.WriteAllText(moduleHeaderPath, sb.ToString());
        }

        private static void WriteBootstrapModuleCpp(
            string moduleCppPath,
            string moduleName,
            string moduleClassName,
            string safeModuleIdentifier,
            BootstrapModuleProfile profile)
        {
            if (File.Exists(moduleCppPath))
            {
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("///////////////////////////////////////////////////////////////////////////////");
            sb.AppendLine("// " + moduleName + ".cpp");
            sb.AppendLine("#include \"TcPch.h\"");
            sb.AppendLine("#pragma hdrstop");
            sb.AppendLine();
            sb.AppendLine("#include \"" + moduleName + ".h\"");
            sb.AppendLine();
            sb.AppendLine("#ifdef _DEBUG");
            sb.AppendLine("#define new DEBUG_NEW");
            sb.AppendLine("#endif");
            sb.AppendLine();
            sb.AppendLine("// AUTO_BOOTSTRAP_TEMPLATE: " + profile.TemplateName);
            sb.AppendLine();
            sb.AppendLine("BEGIN_INTERFACE_MAP(" + moduleClassName + ")");
            sb.AppendLine("\tINTERFACE_ENTRY_ITCOMOBJECT()");
            sb.AppendLine("\tINTERFACE_ENTRY(IID_ITcADI, ITcADI)");
            sb.AppendLine("\tINTERFACE_ENTRY(IID_ITcWatchSource, ITcWatchSource)");
            sb.AppendLine("///<AutoGeneratedContent id=\"InterfaceMap\">");
            sb.AppendLine("///</AutoGeneratedContent>");
            sb.AppendLine("END_INTERFACE_MAP()");
            sb.AppendLine();
            sb.AppendLine("IMPLEMENT_ITCOMOBJECT(" + moduleClassName + ")");
            sb.AppendLine("IMPLEMENT_ITCOMOBJECT_SETSTATE_LOCKOP2(" + moduleClassName + ")");
            sb.AppendLine("IMPLEMENT_ITCADI(" + moduleClassName + ")");
            sb.AppendLine("IMPLEMENT_ITCWATCHSOURCE(" + moduleClassName + ")");
            sb.AppendLine();
            sb.AppendLine("BEGIN_SETOBJPARA_MAP(" + moduleClassName + ")");
            sb.AppendLine("\tSETOBJPARA_DATAAREA_MAP()");
            sb.AppendLine("///<AutoGeneratedContent id=\"SetObjectParameterMap\">");
            sb.AppendLine("///</AutoGeneratedContent>");
            sb.AppendLine("END_SETOBJPARA_MAP()");
            sb.AppendLine();
            sb.AppendLine("BEGIN_GETOBJPARA_MAP(" + moduleClassName + ")");
            sb.AppendLine("\tGETOBJPARA_DATAAREA_MAP()");
            sb.AppendLine("///<AutoGeneratedContent id=\"GetObjectParameterMap\">");
            sb.AppendLine("///</AutoGeneratedContent>");
            sb.AppendLine("END_GETOBJPARA_MAP()");
            sb.AppendLine();
            sb.AppendLine("BEGIN_OBJPARAWATCH_MAP(" + moduleClassName + ")");
            sb.AppendLine("\tOBJPARAWATCH_DATAAREA_MAP()");
            sb.AppendLine("///<AutoGeneratedContent id=\"ObjectParameterWatchMap\">");
            sb.AppendLine("///</AutoGeneratedContent>");
            sb.AppendLine("END_OBJPARAWATCH_MAP()");
            sb.AppendLine();
            sb.AppendLine("BEGIN_OBJDATAAREA_MAP(" + moduleClassName + ")");
            sb.AppendLine("///<AutoGeneratedContent id=\"ObjectDataAreaMap\">");
            sb.AppendLine("///</AutoGeneratedContent>");
            sb.AppendLine("END_OBJDATAAREA_MAP()");
            sb.AppendLine();
            sb.AppendLine(moduleClassName + "::" + moduleClassName + "()");
            sb.AppendLine("\t: m_Trace(m_TraceLevelMax, m_spSrv)");
            sb.AppendLine("\t, m_TraceLevelMax(tlAlways)");
            sb.AppendLine("\t, m_counter(0)");
            sb.AppendLine("{");
            if (profile.IncludeAdsPort)
            {
                sb.AppendLine("\tm_DefaultAdsPort = 0;");
                sb.AppendLine("\tm_ContextAdsPort = 0;");
            }
            if (profile.IncludeCyclicIoHints)
            {
                sb.AppendLine("\tm_InputValue = 0.0;");
                sb.AppendLine("\tm_OutputValue = 0.0;");
            }
            if (profile.IncludeDataPointerHint)
            {
                sb.AppendLine("\tm_pDataPointer = NULL;");
            }
            if (profile.IncludeOnlineChangeHint)
            {
                sb.AppendLine("\tm_EnableOnlineChange = TRUE;");
            }
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine(moduleClassName + "::~" + moduleClassName + "()");
            sb.AppendLine("{");
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine("IMPLEMENT_ITCOMOBJECT_SETOBJSTATE_IP_PI(" + moduleClassName + ")");
            sb.AppendLine();
            sb.AppendLine("HRESULT " + moduleClassName + "::SetObjStatePS(PTComInitDataHdr pInitData)");
            sb.AppendLine("{");
            sb.AppendLine("\tHRESULT hr = S_OK;");
            sb.AppendLine("\tIMPLEMENT_ITCOMOBJECT_EVALUATE_INITDATA(pInitData);");
            sb.AppendLine("\treturn hr;");
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine("HRESULT " + moduleClassName + "::SetObjStateSO()");
            sb.AppendLine("{");
            sb.AppendLine("\tHRESULT hr = S_OK;");
            sb.AppendLine("\thr = FAILED(hr) ? hr : AddModuleToCaller();");
            sb.AppendLine("\tif (FAILED(hr))");
            sb.AppendLine("\t{");
            sb.AppendLine("\t\tRemoveModuleFromCaller();");
            sb.AppendLine("\t}");
            sb.AppendLine("\treturn hr;");
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine("HRESULT " + moduleClassName + "::SetObjStateOS()");
            sb.AppendLine("{");
            sb.AppendLine("\tHRESULT hr = S_OK;");
            sb.AppendLine("\tRemoveModuleFromCaller();");
            sb.AppendLine("\treturn hr;");
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine("HRESULT " + moduleClassName + "::SetObjStateSP()");
            sb.AppendLine("{");
            sb.AppendLine("\tHRESULT hr = S_OK;");
            sb.AppendLine("\treturn hr;");
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine("///<AutoGeneratedContent id=\"ImplementationOf_ITcCyclic\">");
            sb.AppendLine("///</AutoGeneratedContent>");
            sb.AppendLine();
            sb.AppendLine("HRESULT " + moduleClassName + "::AddModuleToCaller()");
            sb.AppendLine("{");
            sb.AppendLine("\tHRESULT hr = S_OK;");
            sb.AppendLine("\tif (m_spCyclicCaller.HasOID())");
            sb.AppendLine("\t{");
            sb.AppendLine("\t\tif (SUCCEEDED_DBG(hr = m_spSrv->TcQuerySmartObjectInterface(m_spCyclicCaller)))");
            sb.AppendLine("\t\t{");
            sb.AppendLine("\t\t\tif (FAILED(hr = m_spCyclicCaller->AddModule(m_spCyclicCaller, THIS_CAST(ITcCyclic))))");
            sb.AppendLine("\t\t\t{");
            sb.AppendLine("\t\t\t\tm_spCyclicCaller = NULL;");
            sb.AppendLine("\t\t\t}");
            sb.AppendLine("\t\t}");
            sb.AppendLine("\t}");
            sb.AppendLine("\telse");
            sb.AppendLine("\t{");
            sb.AppendLine("\t\thr = ADS_E_INVALIDOBJID;");
            sb.AppendLine("\t\tSUCCEEDED_DBGT(hr, \"Invalid OID specified for caller task\");");
            sb.AppendLine("\t}");
            sb.AppendLine("\treturn hr;");
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine("VOID " + moduleClassName + "::RemoveModuleFromCaller()");
            sb.AppendLine("{");
            sb.AppendLine("\tif (m_spCyclicCaller)");
            sb.AppendLine("\t{");
            sb.AppendLine("\t\tm_spCyclicCaller->RemoveModule(m_spCyclicCaller);");
            sb.AppendLine("\t}");
            sb.AppendLine("\tm_spCyclicCaller = NULL;");
            sb.AppendLine("}");

            File.WriteAllText(moduleCppPath, sb.ToString());
        }

        private static void UpsertModuleDefinitionInTmc(
            string tmcPath,
            string moduleName,
            string classFactoryName,
            string moduleGuid,
            string safeModuleIdentifier,
            BootstrapModuleProfile profile)
        {
            XDocument doc = XDocument.Load(tmcPath, LoadOptions.PreserveWhitespace);
            XNamespace ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;

            XElement root = doc.Root;
            if (root == null)
            {
                throw new Exception("TMC文件内容为空。");
            }

            XElement modulesElement = root.Elements().FirstOrDefault(e => e.Name.LocalName == "Modules");
            if (modulesElement == null)
            {
                modulesElement = new XElement(ns + "Modules");
                XElement libraryElement = root.Elements().FirstOrDefault(e => e.Name.LocalName == "Library");
                if (libraryElement != null)
                {
                    libraryElement.AddBeforeSelf(modulesElement);
                }
                else
                {
                    root.Add(modulesElement);
                }
            }

            bool moduleAlreadyExists = modulesElement.Elements()
                .Any(e => e.Name.LocalName == "Module" &&
                    string.Equals(GetChildElementValue(e, "Name"), moduleName, StringComparison.OrdinalIgnoreCase));
            if (moduleAlreadyExists)
            {
                return;
            }

            XElement parametersElement = new XElement(ns + "Parameters",
                new XElement(ns + "Parameter", new XAttribute("HideParameter", "true"),
                    new XElement(ns + "Name", "TraceLevelMax"),
                    new XElement(ns + "Comment", "Controls the amount of log messages."),
                    new XElement(ns + "BaseType", new XAttribute("GUID", "{8007ae3b-86bb-40f2-b385-ef87fcc239a4}"), "TcTraceLevel"),
                    new XElement(ns + "PTCID", "#x03002103"),
                    new XElement(ns + "ContextId", "1")));

            if (profile.IncludeAdsPort)
            {
                parametersElement.Add(new XElement(ns + "Parameter",
                    new XElement(ns + "Name", "DefaultAdsPort"),
                    new XElement(ns + "Comment", "Auto bootstrap by template: with ADS port"),
                    new XElement(ns + "Type", "WORD"),
                    new XElement(ns + "PTCID", "#x00000001"),
                    new XElement(ns + "ContextId", "1")));
            }

            XElement propertiesElement = new XElement(ns + "Properties",
                new XElement(ns + "Property",
                    new XElement(ns + "Name", "BootstrapTemplate"),
                    new XElement(ns + "Value", profile.TemplateName)));

            if (profile.IncludeCyclicIoHints)
            {
                propertiesElement.Add(new XElement(ns + "Property",
                    new XElement(ns + "Name", "CyclicIoHint"),
                    new XElement(ns + "Value", "true")));
            }

            if (profile.IncludeDataPointerHint)
            {
                propertiesElement.Add(new XElement(ns + "Property",
                    new XElement(ns + "Name", "DataPointerHint"),
                    new XElement(ns + "Value", "true")));
            }

            if (profile.IncludeRealtimeContextHint)
            {
                propertiesElement.Add(new XElement(ns + "Property",
                    new XElement(ns + "Name", "RealtimeContextHint"),
                    new XElement(ns + "Value", "true")));
            }

            if (profile.IncludeOnlineChangeHint)
            {
                propertiesElement.Add(new XElement(ns + "Property",
                    new XElement(ns + "Name", "OnlineChangeHint"),
                    new XElement(ns + "Value", "true")));
            }

            XElement moduleElement = new XElement(ns + "Module",
                new XAttribute("GUID", moduleGuid),
                new XAttribute("Group", "C++"),
                new XElement(ns + "Name", moduleName),
                new XElement(ns + "CLSID", new XAttribute("ClassFactory", classFactoryName), moduleGuid),
                new XElement(ns + "Licenses",
                    new XElement(ns + "License",
                        new XElement(ns + "LicenseId", "{304D006A-8299-4560-AB79-438534B50288}"),
                        new XElement(ns + "Comment", "TC3 C++"))),
                new XElement(ns + "InitSequence", "PSO"),
                new XElement(ns + "Contexts",
                    new XElement(ns + "Context",
                        new XElement(ns + "Id", "1"))),
                new XElement(ns + "Interfaces",
                    new XElement(ns + "Interface",
                        new XAttribute("DisableCodeGeneration", "true"),
                        new XElement(ns + "Type", new XAttribute("GUID", "{00000012-0000-0000-E000-000000000064}"), "ITComObject")),
                    new XElement(ns + "Interface",
                        new XElement(ns + "Type", new XAttribute("GUID", "{03000010-0000-0000-E000-000000000064}"), "ITcCyclic")),
                    new XElement(ns + "Interface",
                        new XAttribute("DisableCodeGeneration", "true"),
                        new XElement(ns + "Type", new XAttribute("GUID", "{03000012-0000-0000-E000-000000000064}"), "ITcADI")),
                    new XElement(ns + "Interface",
                        new XAttribute("DisableCodeGeneration", "true"),
                        new XElement(ns + "Type", new XAttribute("GUID", "{03000018-0000-0000-E000-000000000064}"), "ITcWatchSource"))),
                parametersElement,
                new XElement(ns + "DataAreas"),
                new XElement(ns + "InterfacePointers",
                    new XElement(ns + "InterfacePointer",
                        new XElement(ns + "PTCID", "#x03002060"),
                        new XElement(ns + "Name", "CyclicCaller"),
                        new XElement(ns + "Type", new XAttribute("GUID", "{0300001e-0000-0000-e000-000000000064}"), "ITcCyclicCaller"))),
                new XElement(ns + "DataPointers"),
                new XElement(ns + "Deployment"),
                propertiesElement);

            modulesElement.Add(moduleElement);
            doc.Save(tmcPath, SaveOptions.DisableFormatting);
        }

        private static void AddModuleFilesToProjectFiles(string vcxprojPath, string moduleName)
        {
            XDocument doc = XDocument.Load(vcxprojPath, LoadOptions.PreserveWhitespace);
            XNamespace ns = doc.Root?.Name.Namespace ?? XNamespace.None;

            bool hasCpp = doc.Descendants(ns + "ClCompile")
                .Any(x => string.Equals(x.Attribute("Include")?.Value, moduleName + ".cpp", StringComparison.OrdinalIgnoreCase));
            bool hasHeader = doc.Descendants(ns + "ClInclude")
                .Any(x => string.Equals(x.Attribute("Include")?.Value, moduleName + ".h", StringComparison.OrdinalIgnoreCase));

            if (!hasCpp)
            {
                XElement compileGroup = doc.Descendants(ns + "ItemGroup")
                    .FirstOrDefault(g => g.Elements(ns + "ClCompile").Any());
                if (compileGroup == null)
                {
                    compileGroup = new XElement(ns + "ItemGroup");
                    doc.Root?.Add(compileGroup);
                }

                compileGroup.Add(new XElement(ns + "ClCompile", new XAttribute("Include", moduleName + ".cpp")));
            }

            if (!hasHeader)
            {
                XElement includeGroup = doc.Descendants(ns + "ItemGroup")
                    .FirstOrDefault(g => g.Elements(ns + "ClInclude").Any());
                if (includeGroup == null)
                {
                    includeGroup = new XElement(ns + "ItemGroup");
                    doc.Root?.Add(includeGroup);
                }

                includeGroup.Add(new XElement(ns + "ClInclude", new XAttribute("Include", moduleName + ".h")));
            }

            doc.Save(vcxprojPath, SaveOptions.DisableFormatting);
        }

        private static void AddModuleFilesToVcxprojFilters(string filtersPath, string moduleName)
        {
            XDocument doc = XDocument.Load(filtersPath, LoadOptions.PreserveWhitespace);
            XNamespace ns = doc.Root?.Name.Namespace ?? XNamespace.None;

            bool hasCpp = doc.Descendants(ns + "ClCompile")
                .Any(x => string.Equals(x.Attribute("Include")?.Value, moduleName + ".cpp", StringComparison.OrdinalIgnoreCase));
            bool hasHeader = doc.Descendants(ns + "ClInclude")
                .Any(x => string.Equals(x.Attribute("Include")?.Value, moduleName + ".h", StringComparison.OrdinalIgnoreCase));

            XElement itemGroup = doc.Root?.Elements(ns + "ItemGroup")
                .FirstOrDefault(g => g.Elements().Any(e => e.Name.LocalName == "ClCompile" || e.Name.LocalName == "ClInclude"));

            if (itemGroup == null)
            {
                itemGroup = new XElement(ns + "ItemGroup");
                doc.Root?.Add(itemGroup);
            }

            if (!hasHeader)
            {
                itemGroup.Add(
                    new XElement(ns + "ClInclude",
                        new XAttribute("Include", moduleName + ".h"),
                        new XElement(ns + "Filter", "Header Files")));
            }

            if (!hasCpp)
            {
                itemGroup.Add(
                    new XElement(ns + "ClCompile",
                        new XAttribute("Include", moduleName + ".cpp"),
                        new XElement(ns + "Filter", "Source Files")));
            }

            doc.Save(filtersPath, SaveOptions.DisableFormatting);
        }

        private static void EnsureModuleFilesInProject(DTE dte, string moduleName, string moduleHeaderPath, string moduleCppPath)
        {
            Project project = FindProjectByName(dte?.Solution, CurrentCppProjectName);
            if (project == null)
            {
                throw new Exception($"未在解决方案中找到项目：{CurrentCppProjectName}");
            }

            EnsureFileAddedToProject(project, moduleHeaderPath);
            EnsureFileAddedToProject(project, moduleCppPath);
        }

        private static Project FindProjectByName(Solution solution, string projectName)
        {
            if (solution == null || string.IsNullOrWhiteSpace(projectName))
            {
                return null;
            }

            foreach (Project project in EnumerateProjects(solution.Projects))
            {
                if (project != null &&
                    string.Equals(project.Name, projectName, StringComparison.OrdinalIgnoreCase))
                {
                    return project;
                }
            }

            return null;
        }

        private static IEnumerable<Project> EnumerateProjects(Projects projects)
        {
            if (projects == null)
            {
                yield break;
            }

            for (int i = 1; i <= projects.Count; i++)
            {
                Project project = null;
                try
                {
                    project = projects.Item(i);
                }
                catch
                {
                    continue;
                }

                if (project == null)
                {
                    continue;
                }

                if (string.Equals(project.Kind, VS_SOLUTION_FOLDER_KIND, StringComparison.OrdinalIgnoreCase))
                {
                    foreach (Project nested in EnumerateProjectItems(project.ProjectItems))
                    {
                        yield return nested;
                    }
                }
                else
                {
                    yield return project;
                }
            }
        }

        private static IEnumerable<Project> EnumerateProjectItems(ProjectItems projectItems)
        {
            if (projectItems == null)
            {
                yield break;
            }

            for (int i = 1; i <= projectItems.Count; i++)
            {
                ProjectItem item = null;
                try
                {
                    item = projectItems.Item(i);
                }
                catch
                {
                    continue;
                }

                Project subProject = null;
                try
                {
                    subProject = item?.SubProject;
                }
                catch
                {
                }

                if (subProject == null)
                {
                    continue;
                }

                if (string.Equals(subProject.Kind, VS_SOLUTION_FOLDER_KIND, StringComparison.OrdinalIgnoreCase))
                {
                    foreach (Project nested in EnumerateProjectItems(subProject.ProjectItems))
                    {
                        yield return nested;
                    }
                }
                else
                {
                    yield return subProject;
                }
            }
        }

        private static void EnsureFileAddedToProject(Project project, string filePath)
        {
            if (project == null || string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return;
            }

            string fullPath = Path.GetFullPath(filePath);
            if (ProjectContainsFile(project.ProjectItems, fullPath))
            {
                return;
            }

            RetryComCall(() => project.ProjectItems.AddFromFile(fullPath), 5, 300);
        }

        private static bool ProjectContainsFile(ProjectItems projectItems, string fullPath)
        {
            if (projectItems == null || string.IsNullOrWhiteSpace(fullPath))
            {
                return false;
            }

            for (int i = 1; i <= projectItems.Count; i++)
            {
                ProjectItem item = null;
                try
                {
                    item = projectItems.Item(i);
                }
                catch
                {
                    continue;
                }

                if (item == null)
                {
                    continue;
                }

                try
                {
                    if (item.FileCount > 0)
                    {
                        for (short fileIndex = 1; fileIndex <= item.FileCount; fileIndex++)
                        {
                            string itemPath = item.FileNames[fileIndex];
                            if (string.Equals(Path.GetFullPath(itemPath), fullPath, StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                        }
                    }
                }
                catch
                {
                }

                try
                {
                    if (ProjectContainsFile(item.ProjectItems, fullPath))
                    {
                        return true;
                    }
                }
                catch
                {
                }
            }

            return false;
        }

        // 这里的重点不是“调用一次 CreateChild”，而是把模块创建动作包装成稳定流程：
        // 1. 先确认项目目录、vcxproj、TwinCAT 树节点都已经出现
        // 2. 每轮都重新获取项目节点，避免继续使用已经失效的 COM 对象
        // 3. 对 RPC busy / server retry later 这类典型 COM 时序错误做重试
        private static void CreateTcCppModuleStable(DTE dte, ITcSysManager sysManager, string cppProjectName, string moduleName, string wizardId)
        {
            Exception lastException = null;

            for (int attempt = 1; attempt <= 6; attempt++)
            {
                ITcSmTreeItem cppProject = null;
                try
                {
                    Console.WriteLine($"→ 正在尝试创建模块（第 {attempt} 次）...");
                    SaveAll(dte);

                    EnsureCppProjectReadyForModuleCreation(dte, sysManager, cppProjectName);

                    cppProject = GetCppProjectTreeItem(sysManager, cppProjectName);
                    if (cppProject == null)
                    {
                        throw new Exception($"未找到 C++ 项目树节点：TIXC^{cppProjectName}");
                    }

                    RetryComCall(() => cppProject.CreateChild(moduleName, 1, "", wizardId), 30, 500);

                    SaveAll(dte);
                    WaitForGeneratedModuleFiles(dte, moduleName);

                    Console.WriteLine("→ 模块创建成功并已检测到生成文件。");
                    return;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    Console.WriteLine($"× 第 {attempt} 次创建模块失败：{ex.Message}");

                    // 关键：如果是 wizard 本体失败，不要再重试
                    if (!string.IsNullOrEmpty(ex.Message) &&
                        ex.Message.IndexOf("Add module class wizard failed", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        throw new Exception(
                            "创建模块失败：TwinCAT 4026 的 Add Module Class Wizard 本体执行失败。" +
                            "这通常不是代码时序问题，而是本机 TwinCAT C++ Module Wizard 安装/注册不完整。" +
                            "请先在 Visual Studio 里手工尝试给该 C++ Project 添加一个 TwinCAT Module Class；" +
                            "如果手工也失败，就需要修复 4026 的 C++ 向导环境。");
                    }

                    RecoverVsAfterWizardFailure(dte);
                    System.Threading.Thread.Sleep(2000 + attempt * 1000);
                }
                finally
                {
                    ReleaseComIfNeeded(cppProject);
                }
            }

            throw new Exception($"创建模块时发生异常（模板：{wizardId}）：{lastException?.Message}");
        }

        // 创建模块前先做“工程就绪检查”。
        // 如果工程目录、工程文件或 TIXC 节点还没稳定，继续往下调用向导通常只会得到误导性的异常。
        private static void EnsureCppProjectReadyForModuleCreation(DTE dte, ITcSysManager sysManager, string cppProjectName)
        {
            SaveAll(dte);
            WaitForCppProjectStabilized(dte, sysManager, cppProjectName);

            ITcSmTreeItem cppProject = null;
            try
            {
                cppProject = GetCppProjectTreeItem(sysManager, cppProjectName);
                if (cppProject == null)
                {
                    throw new Exception("C++ 项目树节点尚未出现在 TIXC 下。");
                }
            }
            finally
            {
                ReleaseComIfNeeded(cppProject);
            }

            System.Threading.Thread.Sleep(1500);
        }

        // “工程稳定”至少满足三件事：
        // 磁盘上有工程目录、磁盘上有 .vcxproj、TwinCAT 配置树里能查到 TIXC^<project>。
        // 少任何一个条件，后面无论建模块还是发布，都会变成高概率随机失败。
        private static void WaitForCppProjectStabilized(DTE dte, ITcSysManager sysManager, string cppProjectName)
        {
            DateTime endTime = DateTime.Now.AddMilliseconds(DEFAULT_FILE_WAIT_TIMEOUT);
            string solutionDir = GetSolutionDirectory(dte);

            while (DateTime.Now < endTime)
            {
                bool dirExists = Directory.Exists(Path.Combine(solutionDir, cppProjectName));
                bool hasProjectFile = SafeEnumerateFiles(solutionDir, "*.vcxproj")
                    .Any(p => p.IndexOf("\\" + cppProjectName + "\\", StringComparison.OrdinalIgnoreCase) >= 0);

                bool treeReady = false;
                try
                {
                    ITcSmTreeItem item = GetCppProjectTreeItem(sysManager, cppProjectName);
                    treeReady = item != null;
                    ReleaseComIfNeeded(item);
                }
                catch
                {
                    treeReady = false;
                }

                if (dirExists && hasProjectFile && treeReady)
                {
                    System.Threading.Thread.Sleep(1500);
                    return;
                }

                System.Threading.Thread.Sleep(500);
            }

            throw new Exception($"等待 C++ 项目稳定超时：{cppProjectName}");
        }

        private static ITcSmTreeItem GetCppProjectTreeItem(ITcSysManager sysManager, string cppProjectName)
        {
            return RetryComCall(() => sysManager.LookupTreeItem("TIXC^" + cppProjectName), 20, 300);
        }

        private static void RecoverVsAfterWizardFailure(DTE dte)
        {
            try
            {
                SaveAll(dte);
            }
            catch
            {
            }

            try
            {
                RetryComCall(() => dte.SuppressUI = false, 3, 100);
            }
            catch
            {
            }

            System.Threading.Thread.Sleep(1000);
        }

        private static void ExecuteTmcCodeGenerator(ITcSmTreeItem cppProject)
        {
            string tmcGeneratorXml = @"<?xml version=""1.0"" encoding=""UTF-16""?>
<TreeItem>
  <CppProjectDef>
    <Methods>
      <StartTmcCodeGenerator>
        <Active>true</Active>
      </StartTmcCodeGenerator>
    </Methods>
  </CppProjectDef>
</TreeItem>";

            RetryComCall(() => cppProject.ConsumeXml(tmcGeneratorXml));
        }

        private static void DumpTmcCandidates(DTE dte)
        {
            Console.WriteLine("→ 开始扫描 TMC 候选文件...");

            string solutionDir = GetSolutionDirectory(dte);
            string projectDir = Path.Combine(solutionDir, CurrentCppProjectName);

            if (Directory.Exists(projectDir))
            {
                foreach (var f in SafeEnumerateFiles(projectDir, "*.tmc").OrderBy(x => x))
                {
                    Console.WriteLine("  [工程内] " + f);
                }
            }

            foreach (string root in DefaultTmcBasePaths)
            {
                try
                {
                    if (!Directory.Exists(root)) continue;
                    foreach (var f in Directory.EnumerateFiles(root, "*.tmc", SearchOption.AllDirectories)
                             .Where(x => x.IndexOf(CurrentCppProjectName, StringComparison.OrdinalIgnoreCase) >= 0)
                             .Take(20))
                    {
                        Console.WriteLine("  [发布目录] " + f);
                    }
                }
                catch { }
            }

            foreach (string root in TwinCatRepositoryRoots)
            {
                try
                {
                    if (!Directory.Exists(root)) continue;
                    foreach (var f in Directory.EnumerateFiles(root, "*.tmc", SearchOption.AllDirectories)
                             .Where(x => x.IndexOf(CurrentCppProjectName, StringComparison.OrdinalIgnoreCase) >= 0)
                             .Take(20))
                    {
                        Console.WriteLine("  [Repository] " + f);
                    }
                }
                catch { }
            }
        }

        private static bool ExecutePublishModules(ITcSmTreeItem cppProject)
        {
            DateTime publishStartTime = DateTime.Now;
            string publishModulesXml = @"<?xml version=""1.0"" encoding=""UTF-16""?>
<TreeItem>
  <CppProjectDef>
    <Methods>
      <PublishModules>
        <Active>true</Active>
      </PublishModules>
    </Methods>
  </CppProjectDef>
</TreeItem>";

            RetryComCall(() => cppProject.ConsumeXml(publishModulesXml));
            Console.WriteLine("→ 等待5秒，确保发布文件同步完成...");
            System.Threading.Thread.Sleep(5000);

            return TryGetPublishedTmcFilePath(publishStartTime) != null;
        }

        private static void SaveAll(DTE dte)
        {
            if (dte == null) return;

            try
            {
                RetryComCall(() => dte.ExecuteCommand("File.SaveAll"), 5, 300);
            }
            catch
            {
            }

            try
            {
                RetryComCall(() =>
                {
                    if (dte.Documents != null)
                    {
                        dte.Documents.SaveAll();
                    }
                }, 5, 300);
            }
            catch
            {
            }
        }

        private static void WaitForProjectArtifacts(DTE dte)
        {
            DateTime endTime = DateTime.Now.AddMilliseconds(DEFAULT_FILE_WAIT_TIMEOUT);

            while (DateTime.Now < endTime)
            {
                string cppFilePath = TryGetGeneratedModuleCppFilePath(dte);

                // 对 4026 自带骨架场景，先只要求找到可编辑的 cpp
                if (!string.IsNullOrEmpty(cppFilePath))
                {
                    System.Threading.Thread.Sleep(1000);
                    return;
                }

                System.Threading.Thread.Sleep(500);
            }

            throw new Exception("等待模块源码生成超时，请确认 Visual Studio 已完成向导文件落盘。");
        }

        private static void WaitForGeneratedModuleFiles(DTE dte, string moduleName)
        {
            DateTime endTime = DateTime.Now.AddMilliseconds(DEFAULT_FILE_WAIT_TIMEOUT);

            while (DateTime.Now < endTime)
            {
                string cppFilePath = TryFindProjectOwnedFile(dte, moduleName + ".cpp");
                if (!string.IsNullOrEmpty(cppFilePath))
                {
                    System.Threading.Thread.Sleep(1000);
                    return;
                }

                System.Threading.Thread.Sleep(500);
            }

            throw new Exception($"等待模块源文件生成超时：{moduleName}.cpp");
        }

        private static void ValidateModuleIntegratedIntoProjectModel(DTE dte, string moduleName)
        {
            string solutionDir = GetSolutionDirectory(dte);
            string projectDir = Path.Combine(solutionDir, CurrentCppProjectName);
            string vcxprojPath = Path.Combine(projectDir, CurrentCppProjectName + ".vcxproj");
            string moduleCppPath = Path.Combine(projectDir, moduleName + ".cpp");
            string moduleHeaderPath = Path.Combine(projectDir, moduleName + ".h");

            if (!File.Exists(moduleCppPath) || !File.Exists(moduleHeaderPath))
            {
                throw new Exception($"模块文件未生成完整：{moduleName}.h/.cpp");
            }

            if (!File.Exists(vcxprojPath))
            {
                throw new Exception($"未找到项目文件：{vcxprojPath}");
            }

            string vcxprojText = File.ReadAllText(vcxprojPath);
            bool hasCpp = vcxprojText.IndexOf(moduleName + ".cpp", StringComparison.OrdinalIgnoreCase) >= 0;
            bool hasHeader = vcxprojText.IndexOf(moduleName + ".h", StringComparison.OrdinalIgnoreCase) >= 0;

            if (!hasCpp || !hasHeader)
            {
                throw new Exception(
                    $"TwinCAT 向导没有把模块真正加入项目模型：{moduleName}.h/.cpp 未出现在 vcxproj 中。" +
                    "当前不会再继续走发布/TMC/TcCOM 流程，请先检查本机 4026 C++ Module Wizard 是否可正常在 IDE 中手工添加模块。");
            }
        }

        private static string PatchProjectTmc(DTE dte)
        {
            string tmcFilePath = GetProjectTmcFilePath(dte);
            Console.WriteLine($"→ 准备修改工程TMC：{tmcFilePath}");

            if (!File.Exists(tmcFilePath))
            {
                throw new FileNotFoundException("工程TMC文件不存在", tmcFilePath);
            }

            string backupPath = tmcFilePath + ".bak";
            File.Copy(tmcFilePath, backupPath, true);

            XDocument tmcDoc = XDocument.Load(tmcFilePath, LoadOptions.PreserveWhitespace);

            XElement moduleElement = tmcDoc.Descendants().FirstOrDefault(x => x.Name.LocalName == "Module");

            // 4026 的某些 project-level tmc 结构可能没有 <Module>，直接兜底在全局查 Parameters
            XElement parametersElement = null;

            if (moduleElement != null)
            {
                parametersElement = moduleElement.Descendants().FirstOrDefault(x => x.Name.LocalName == "Parameters");
            }
            else
            {
                Console.WriteLine("→ 当前工程TMC中未找到<Module>节点，尝试按 4026 的 project-level TMC 结构直接查找 <Parameters> ...");
                parametersElement = tmcDoc.Descendants().FirstOrDefault(x => x.Name.LocalName == "Parameters");
            }

            if (parametersElement == null)
            {
                throw new Exception("工程TMC中未找到<Parameters>节点，当前 TMC 可能不是可直接补丁的模块描述文件。");
            }

            List<XElement> parameterElements = parametersElement.Elements()
                .Where(x => x.Name.LocalName == "Parameter")
                .ToList();

            string parameterNames = string.Join(", ", parameterElements
                .Select(x => GetChildElementValue(x, "Name"))
                .Where(x => !string.IsNullOrEmpty(x)));
            Console.WriteLine($"→ 当前工程TMC里的参数：{parameterNames}");

            bool patched = false;

            XElement inlineStructuredParameter = parameterElements.FirstOrDefault(x =>
                !string.Equals(GetChildElementValue(x, "Name"), "TraceLevelMax", StringComparison.OrdinalIgnoreCase) &&
                x.Elements().Count(e => e.Name.LocalName == "SubItem") >= 3);

            if (inlineStructuredParameter != null)
            {
                PatchStructuredTypeElement(inlineStructuredParameter);
                patched = true;
                Console.WriteLine("→ 已找到模块里现成的结构体参数，直接改成 Gain / Enable / VelocityLimit。");
            }

            if (!patched)
            {
                XElement editableParameterElement = parameterElements.FirstOrDefault(x =>
                    !string.Equals(GetChildElementValue(x, "Name"), "TraceLevelMax", StringComparison.OrdinalIgnoreCase));

                if (editableParameterElement != null)
                {
                    string parameterTypeName = GetChildElementValue(editableParameterElement, "Type");
                    XElement referencedDataType = FindReferencedDataType(tmcDoc, parameterTypeName);

                    if (referencedDataType != null && referencedDataType.Elements().Count(e => e.Name.LocalName == "SubItem") >= 3)
                    {
                        PatchStructuredTypeElement(referencedDataType);
                        SetOrCreateChildElementValue(editableParameterElement, "Name", "Parameter");
                        patched = true;
                        Console.WriteLine("→ 已找到参数引用的数据类型结构体，并完成修改。");
                    }
                }
            }

            if (!patched)
            {
                XElement fallbackStructuredDataType = tmcDoc.Descendants()
                    .FirstOrDefault(x =>
                        x.Name.LocalName == "DataType" &&
                        x.Elements().Count(e => e.Name.LocalName == "SubItem") >= 3 &&
                        (GetChildElementValue(x, "Name")?.IndexOf("Parameter", StringComparison.OrdinalIgnoreCase) >= 0));

                if (fallbackStructuredDataType != null)
                {
                    PatchStructuredTypeElement(fallbackStructuredDataType);
                    patched = true;
                    Console.WriteLine("→ 已找到候选结构体类型，并完成修改。");
                }
            }

            if (!patched)
            {
                XElement templateParameterElement = parameterElements.FirstOrDefault();
                if (templateParameterElement == null)
                {
                    throw new Exception("工程TMC中没有任何现成的Parameter节点，无法克隆模板参数。");
                }

                Console.WriteLine("→ 当前模板里没有默认结构体参数，改为克隆一个现有有效参数模板来新增三个标量参数...");

                UpsertScalarParameterByCloningTemplate(parametersElement, templateParameterElement, "Gain", "LREAL", "自动添加参数：Gain");
                UpsertScalarParameterByCloningTemplate(parametersElement, templateParameterElement, "Enable", "BOOL", "自动添加参数：Enable");
                UpsertScalarParameterByCloningTemplate(parametersElement, templateParameterElement, "VelocityLimit", "LREAL", "自动添加参数：VelocityLimit");
                patched = true;
            }

            if (!patched)
            {
                throw new Exception("工程TMC修改失败。");
            }

            tmcDoc.Save(tmcFilePath, SaveOptions.DisableFormatting);
            SaveAll(dte);

            Console.WriteLine($"→ 已备份原始工程TMC：{backupPath}");
            return tmcFilePath;
        }

        // 第一次补 TMC 失败时，不急着直接报错。
        // 更常见的真实原因是 Code Generator / Publish 的产物还没刷新完成，
        // 所以这里会先主动刷新相关文件，再给一次重试机会。
        private static bool TryPatchProjectTmcWithRecovery(DTE dte, ITcSysManager sysManager, out string patchedTmcPath)
        {
            patchedTmcPath = null;
            Exception lastException = null;

            for (int attempt = 1; attempt <= 2; attempt++)
            {
                try
                {
                    patchedTmcPath = PatchProjectTmc(dte);
                    return true;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    Console.WriteLine($"× 第 {attempt} 次修改工程TMC失败：{ex.Message}");

                    if (attempt == 1)
                    {
                        Console.WriteLine("→ 尝试自动刷新TMC产物（Code Generator + Publish Modules）后重试...");
                        TryRefreshTmcArtifacts(dte, sysManager);
                        DumpTmcCandidates(dte);
                    }
                }
            }

            Console.WriteLine($"→ 诊断信息：{lastException?.Message}");
            return false;
        }

        // 这里不是做业务修改，而是强制 TwinCAT 重新吐出一轮最新产物，
        // 避免后续继续命中过期 TMC、旧 GUID 或半生成状态的文件。
        private static void TryRefreshTmcArtifacts(DTE dte, ITcSysManager sysManager)
        {
            ITcSmTreeItem cppProject = null;
            try
            {
                cppProject = GetCppProjectTreeItem(sysManager, CurrentCppProjectName);
                if (cppProject == null)
                {
                    Console.WriteLine("× 未找到当前C++项目树节点，无法自动刷新TMC产物。");
                    return;
                }

                ExecuteTmcCodeGenerator(cppProject);
                SaveAll(dte);
                System.Threading.Thread.Sleep(3000);

                ExecutePublishModules(cppProject);
                SaveAll(dte);
                System.Threading.Thread.Sleep(3000);

                Console.WriteLine("√  已完成TMC产物刷新。");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"× 自动刷新TMC产物失败：{ex.Message}");
            }
            finally
            {
                ReleaseComIfNeeded(cppProject);
            }
        }

        private static void PatchStructuredTypeElement(XElement structuredElement)
        {
            List<XElement> subItems = structuredElement.Elements()
                .Where(x => x.Name.LocalName == "SubItem")
                .ToList();

            if (subItems.Count < 3)
            {
                throw new Exception("默认结构体参数的SubItem数量不足3个，无法安全改成 Gain / Enable / VelocityLimit。");
            }

            if (structuredElement.Name.LocalName == "Parameter")
            {
                SetOrCreateChildElementValue(structuredElement, "Name", "Parameter");
            }

            PatchStructuredSubItem(subItems[0], "Gain", "LREAL", 64, 0);
            PatchStructuredSubItem(subItems[1], "Enable", "BOOL", 8, 64);
            PatchStructuredSubItem(subItems[2], "VelocityLimit", "LREAL", 64, 128);

            if (subItems.Count == 3)
            {
                SetOrCreateChildElementValue(structuredElement, "BitSize", "192");
            }
        }

        private static void PatchStructuredSubItem(XElement subItem, string fieldName, string typeName, int bitSize, int bitOffs)
        {
            SetOrCreateChildElementValue(subItem, "Name", fieldName);

            XElement typeElement = GetOrCreateChildElement(subItem, "Type");
            typeElement.RemoveAttributes();
            typeElement.Value = typeName;

            SetOrCreateChildElementValue(subItem, "BitSize", bitSize.ToString());
            SetOrCreateChildElementValue(subItem, "BitOffs", bitOffs.ToString());

            XElement defaultElement = subItem.Elements().FirstOrDefault(x => x.Name.LocalName == "Default");
            if (defaultElement != null)
            {
                defaultElement.Remove();
            }
        }

        // 只注入一段非常小的示例代码，并通过 marker 保证幂等。
        // 这样重复运行工具时不会不断追加同一段代码，也更不容易覆盖用户后续手写逻辑。
        private static bool WriteSimpleCodeToGeneratedModule(DTE dte)
        {
            string cppFilePath = GetGeneratedModuleCppFilePath(dte);
            string sourceCode = File.ReadAllText(cppFilePath);

            const string marker = "// AUTO_WRITTEN_BY_TOOL";
            if (sourceCode.Contains(marker))
            {
                Console.WriteLine("→ 已检测到自动写入标记，跳过重复写入。");
                return true;
            }

            int cyclePos = sourceCode.IndexOf("CycleUpdate", StringComparison.Ordinal);
            if (cyclePos < 0)
            {
                Console.WriteLine("→ 还未找到 CycleUpdate 实现，当前跳过自动写入示例代码。");
                return false;
            }

            int bracePos = sourceCode.IndexOf('{', cyclePos);
            if (bracePos < 0)
            {
                Console.WriteLine("→ 找到了 CycleUpdate 声明，但还没有函数体，当前跳过自动写入示例代码。");
                return false;
            }

            int insertPos = bracePos + 1;
            string injectCode =
                Environment.NewLine +
                "    // AUTO_WRITTEN_BY_TOOL" + Environment.NewLine +
                "    static int s_autoCounter = 0;" + Environment.NewLine +
                "    ++s_autoCounter;" + Environment.NewLine +
                Environment.NewLine;

            File.Copy(cppFilePath, cppFilePath + ".bak", true);
            sourceCode = sourceCode.Insert(insertPos, injectCode);
            File.WriteAllText(cppFilePath, sourceCode);

            Console.WriteLine($"→ 已自动修改模块源码：{cppFilePath}");
            Console.WriteLine($"→ 已自动备份原始文件：{cppFilePath}.bak");
            return true;
        }
        private static void UpsertScalarParameterByCloningTemplate(XElement parametersElement, XElement templateParameterElement, string parameterName, string typeName, string comment)
        {
            XElement targetParameterElement = parametersElement.Elements()
                .FirstOrDefault(x =>
                    x.Name.LocalName == "Parameter" &&
                    string.Equals(GetChildElementValue(x, "Name"), parameterName, StringComparison.OrdinalIgnoreCase));

            if (targetParameterElement == null)
            {
                targetParameterElement = new XElement(templateParameterElement);
                parametersElement.Add(targetParameterElement);
            }

            XAttribute hideAttr = targetParameterElement.Attributes()
                .FirstOrDefault(a => a.Name.LocalName == "HideParameter");
            hideAttr?.Remove();

            List<XElement> childElementsToRemove = targetParameterElement.Elements()
                .Where(e =>
                {
                    string localName = e.Name.LocalName;
                    return localName == "SubItem" ||
                           localName == "EnumInfo" ||
                           localName == "ArrayInfo" ||
                           localName == "Type" ||
                           localName == "BaseType";
                })
                .ToList();

            foreach (XElement element in childElementsToRemove)
            {
                element.Remove();
            }

            SetOrCreateChildElementValue(targetParameterElement, "Name", parameterName);
            SetOrCreateChildElementValue(targetParameterElement, "Comment", comment);
            SetOrCreateChildElementValue(targetParameterElement, "PTCID", GetNextAvailableUserDefinedPtcId(parametersElement, targetParameterElement));

            string templateContextId = GetChildElementValue(templateParameterElement, "ContextId");
            if (!string.IsNullOrEmpty(templateContextId))
            {
                SetOrCreateChildElementValue(targetParameterElement, "ContextId", templateContextId);
            }

            SetParameterTypeInfo(targetParameterElement, typeName);
            SetParameterSizeInfo(targetParameterElement, typeName);
            SetParameterDefaultValue(targetParameterElement, typeName);
            SetParameterConstantName(targetParameterElement, parameterName);
            RemoveTraceSpecificMetadata(targetParameterElement);
        }

        private static void SetParameterTypeInfo(XElement parameterElement, string typeName)
        {
            XElement baseTypeElement = parameterElement.Elements().FirstOrDefault(x => x.Name.LocalName == "BaseType");
            XElement typeElement = parameterElement.Elements().FirstOrDefault(x => x.Name.LocalName == "Type");

            if (baseTypeElement != null)
            {
                baseTypeElement.RemoveAttributes();
                baseTypeElement.RemoveNodes();
                baseTypeElement.Value = typeName;
            }
            else if (typeElement != null)
            {
                typeElement.RemoveAttributes();
                typeElement.RemoveNodes();
                typeElement.Value = typeName;
            }
            else
            {
                parameterElement.Add(new XElement(parameterElement.GetDefaultNamespace() + "BaseType", typeName));
            }

            if (baseTypeElement != null && typeElement != null)
            {
                typeElement.Remove();
            }
        }

        private static void SetParameterSizeInfo(XElement parameterElement, string typeName)
        {
            int bitSize = GetBitSizeForSimpleType(typeName);
            int byteSize = Math.Max(1, bitSize / 8);

            SetOrCreateChildElementValue(parameterElement, "BitSize", bitSize.ToString());
            SetOrCreateChildElementValue(parameterElement, "Size", byteSize.ToString());
            SetOrCreateChildElementValue(parameterElement, "SizeX64", byteSize.ToString());
        }

        private static void SetParameterDefaultValue(XElement parameterElement, string typeName)
        {
            string defaultValue = string.Equals(typeName, "BOOL", StringComparison.OrdinalIgnoreCase) ? "FALSE" : "0";

            XElement defaultElement = parameterElement.Elements().FirstOrDefault(x => x.Name.LocalName == "Default");
            XElement defaultValueElement = parameterElement.Elements().FirstOrDefault(x => x.Name.LocalName == "DefaultValue");

            if (defaultElement != null)
            {
                defaultElement.Value = defaultValue;
            }
            else if (defaultValueElement == null)
            {
                SetOrCreateChildElementValue(parameterElement, "Default", defaultValue);
            }

            if (defaultValueElement != null)
            {
                defaultValueElement.Value = defaultValue;
            }
        }

        private static void SetParameterConstantName(XElement parameterElement, string parameterName)
        {
            XElement constantNameElement = parameterElement.Elements().FirstOrDefault(x => x.Name.LocalName == "ConstantName");
            if (constantNameElement != null)
            {
                constantNameElement.Value = "PID_" + parameterName;
            }
        }

        private static void RemoveTraceSpecificMetadata(XElement parameterElement)
        {
            List<XElement> removeList = parameterElement.Descendants()
                .Where(x =>
                {
                    string localName = x.Name.LocalName;
                    return localName.IndexOf("Predefined", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           localName.IndexOf("Trace", StringComparison.OrdinalIgnoreCase) >= 0;
                })
                .ToList();

            foreach (XElement element in removeList)
            {
                element.Remove();
            }

            foreach (XElement element in parameterElement.DescendantsAndSelf())
            {
                List<XAttribute> removeAttributes = element.Attributes()
                    .Where(x =>
                    {
                        string localName = x.Name.LocalName;
                        return localName.IndexOf("Predefined", StringComparison.OrdinalIgnoreCase) >= 0 ||
                               localName.IndexOf("Trace", StringComparison.OrdinalIgnoreCase) >= 0;
                    })
                    .ToList();

                foreach (XAttribute attribute in removeAttributes)
                {
                    attribute.Remove();
                }
            }
        }

        private static int GetBitSizeForSimpleType(string typeName)
        {
            switch ((typeName ?? string.Empty).ToUpperInvariant())
            {
                case "BOOL": return 8;
                case "LREAL": return 64;
                case "REAL": return 32;
                case "ULINT":
                case "LINT": return 64;
                case "UDINT":
                case "DINT": return 32;
                case "UINT":
                case "INT": return 16;
                case "USINT":
                case "SINT": return 8;
                default: return 32;
            }
        }

        private static string GetNextAvailableUserDefinedPtcId(XElement parametersElement, XElement currentParameterElement)
        {
            HashSet<int> usedIds = new HashSet<int>();

            foreach (XElement parameterElement in parametersElement.Elements().Where(x => x.Name.LocalName == "Parameter"))
            {
                if (object.ReferenceEquals(parameterElement, currentParameterElement))
                {
                    continue;
                }

                string rawPtcId = GetChildElementValue(parameterElement, "PTCID");
                int? parsedPtcId = ParseHexStyleId(rawPtcId);
                if (parsedPtcId.HasValue && parsedPtcId.Value > 0)
                {
                    usedIds.Add(parsedPtcId.Value);
                }
            }

            int nextId = 1;
            while (usedIds.Contains(nextId))
            {
                nextId++;
            }

            return $"#x{nextId:X8}";
        }

        private static int? ParseHexStyleId(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return null;
            }

            string normalized = rawValue.Trim();
            if (normalized.StartsWith("#x", StringComparison.OrdinalIgnoreCase))
            {
                normalized = normalized.Substring(2);
            }

            if (int.TryParse(normalized, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out int parsedValue))
            {
                return parsedValue;
            }

            return null;
        }

        private static XElement FindReferencedDataType(XDocument doc, string dataTypeName)
        {
            if (string.IsNullOrWhiteSpace(dataTypeName))
            {
                return null;
            }

            return doc.Descendants().FirstOrDefault(x =>
                x.Name.LocalName == "DataType" &&
                string.Equals(GetChildElementValue(x, "Name"), dataTypeName, StringComparison.OrdinalIgnoreCase));
        }

        private static XElement GetOrCreateChildElement(XElement parent, string childLocalName)
        {
            XElement child = parent.Elements().FirstOrDefault(x => x.Name.LocalName == childLocalName);
            if (child == null)
            {
                child = new XElement(parent.GetDefaultNamespace() + childLocalName);
                parent.Add(child);
            }

            return child;
        }

        private static void SetOrCreateChildElementValue(XElement parent, string childLocalName, string value)
        {
            XElement child = GetOrCreateChildElement(parent, childLocalName);
            child.Value = value;
        }

        private static string GetChildElementValue(XElement parent, string childLocalName)
        {
            return parent.Elements().FirstOrDefault(x => x.Name.LocalName == childLocalName)?.Value;
        }

        private static string FindPatchableModuleTmcFile(DTE dte)
        {
            string solutionDir = GetSolutionDirectory(dte);
            string projectDir = Path.Combine(solutionDir, CurrentCppProjectName);

            if (!Directory.Exists(projectDir))
            {
                return null;
            }

            var tmcFiles = SafeEnumerateFiles(projectDir, "*.tmc")
                .Where(IsProjectSourcePath)
                .OrderBy(p => p.Length)
                .ToList();

            string bestTmcPath = null;
            int bestScore = 0;

            foreach (string tmcPath in tmcFiles)
            {
                try
                {
                    int score = GetTmcPatchabilityScore(tmcPath);
                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestTmcPath = tmcPath;
                    }
                }
                catch
                {
                    // 忽略坏文件/无法解析文件
                }
            }

            if (!string.IsNullOrEmpty(bestTmcPath))
            {
                Console.WriteLine($"→ 找到候选可补丁 TMC：{bestTmcPath}（评分：{bestScore}）");
            }

            return bestTmcPath;
        }

        private static string GetProjectTmcFilePath(DTE dte)
        {
            // 1) 先找工程目录内可补丁 tmc
            string tmcFilePath = FindPatchableModuleTmcFile(dte);
            if (!string.IsNullOrEmpty(tmcFilePath))
            {
                Console.WriteLine($"→ 在工程目录中找到可补丁TMC：{tmcFilePath}");
                return tmcFilePath;
            }

            // 2) 再找 4026 常见输出位置
            tmcFilePath = TryGetProjectTmcFilePath(dte);
            if (!string.IsNullOrEmpty(tmcFilePath))
            {
                Console.WriteLine($"→ 在发布/Repository目录中找到候选TMC：{tmcFilePath}");
                return tmcFilePath;
            }

            throw new FileNotFoundException(
                "未找到可直接补丁的模块TMC文件（未发现包含 Parameters/Parameter 或可回退结构体的数据）。4026 的 Versioned C++ Project 可能尚未完成发布，或产物已落到 Repository/Publish 目录。",
                "*.tmc");
        }

        private static bool IsPatchableTmc(string tmcPath)
        {
            return GetTmcPatchabilityScore(tmcPath) > 0;
        }

        private static int GetTmcPatchabilityScore(string tmcPath)
        {
            try
            {
                XDocument doc = XDocument.Load(tmcPath);

                bool hasParametersContainer = doc.Descendants().Any(x => x.Name.LocalName == "Parameters");
                bool hasParameterNode = doc.Descendants().Any(x => x.Name.LocalName == "Parameter");
                bool hasStructuredDataType = doc.Descendants().Any(x =>
                    x.Name.LocalName == "DataType" &&
                    x.Elements().Count(e => e.Name.LocalName == "SubItem") >= 3 &&
                    (GetChildElementValue(x, "Name")?.IndexOf("Parameter", StringComparison.OrdinalIgnoreCase) >= 0));

                if (hasParametersContainer) return 3;
                if (hasParameterNode) return 2;
                if (hasStructuredDataType) return 1;
            }
            catch
            {
            }

            return 0;
        }
        private static string TryGetProjectTmcFilePath(DTE dte)
        {
            string solutionDir = GetSolutionDirectory(dte);
            string projectDir = Path.Combine(solutionDir, CurrentCppProjectName);

            // 1) 工程目录内找任意可疑 .tmc
            if (Directory.Exists(projectDir))
            {
                var localCandidates = SafeEnumerateFiles(projectDir, "*.tmc")
                    .Where(p => !p.EndsWith(".bak", StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(File.GetLastWriteTime)
                    .ToList();

                foreach (var candidate in localCandidates)
                {
                    if (IsPatchableTmc(candidate))
                        return candidate;
                }
            }

            // 2) 从发布目录找
            foreach (string root in DefaultTmcBasePaths)
            {
                try
                {
                    if (!Directory.Exists(root)) continue;

                    var candidates = Directory.EnumerateFiles(root, "*.tmc", SearchOption.AllDirectories)
                        .OrderByDescending(File.GetLastWriteTime)
                        .ToList();

                    foreach (var candidate in candidates)
                    {
                        if (candidate.IndexOf(CurrentCppProjectName, StringComparison.OrdinalIgnoreCase) >= 0 &&
                            IsPatchableTmc(candidate))
                        {
                            return candidate;
                        }
                    }
                }
                catch { }
            }

            // 3) 从 Repository 找
            foreach (string root in TwinCatRepositoryRoots)
            {
                try
                {
                    if (!Directory.Exists(root)) continue;

                    var candidates = Directory.EnumerateFiles(root, "*.tmc", SearchOption.AllDirectories)
                        .OrderByDescending(File.GetLastWriteTime)
                        .ToList();

                    foreach (var candidate in candidates)
                    {
                        if (candidate.IndexOf(CurrentCppProjectName, StringComparison.OrdinalIgnoreCase) >= 0 &&
                            IsPatchableTmc(candidate))
                        {
                            return candidate;
                        }
                    }
                }
                catch { }
            }

            return null;
        }

        private static string GetGeneratedModuleCppFilePath(DTE dte)
        {
            string cppFilePath = TryGetGeneratedModuleCppFilePath(dte);
            if (string.IsNullOrEmpty(cppFilePath))
            {
                throw new FileNotFoundException("未找到生成的模块cpp文件", CurrentModuleName + ".cpp");
            }

            return cppFilePath;
        }

        private static string TryGetGeneratedModuleCppFilePath(DTE dte)
        {
            // 先找精确文件名
            string exactPath = TryFindProjectOwnedFile(dte, CurrentModuleName + ".cpp");
            if (!string.IsNullOrEmpty(exactPath))
            {
                return exactPath;
            }

            string solutionDir = GetSolutionDirectory(dte);
            string projectDir = Path.Combine(solutionDir, CurrentCppProjectName);

            if (!Directory.Exists(projectDir))
            {
                return null;
            }

            var candidates = SafeEnumerateFiles(projectDir, "*.cpp")
                .Where(IsProjectSourcePath)
                .Where(p =>
                    !p.EndsWith("TcPch.cpp", StringComparison.OrdinalIgnoreCase) &&
                    p.IndexOf("ClassFactory", StringComparison.OrdinalIgnoreCase) < 0)
                .OrderBy(p => p.Length)
                .ToList();

            return candidates.FirstOrDefault();
        }

        private static string TryFindProjectOwnedFile(DTE dte, string fileName)
        {
            string solutionDir = GetSolutionDirectory(dte);

            return SafeEnumerateFiles(solutionDir, fileName)
                .Where(IsProjectSourcePath)
                .OrderBy(path => path.IndexOf(Path.DirectorySeparatorChar + CurrentCppProjectName + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) >= 0 ? 0 : 1)
                .ThenBy(path => path.Length)
                .FirstOrDefault();
        }

        private static string GetSolutionDirectory(DTE dte)
        {
            if (!string.IsNullOrEmpty(CurrentSolutionDirectory) && Directory.Exists(CurrentSolutionDirectory))
            {
                return CurrentSolutionDirectory;
            }

            string solutionFullName = RetryComCall(() =>
            {
                if (dte?.Solution == null || string.IsNullOrEmpty(dte.Solution.FullName))
                {
                    throw new Exception("当前未加载解决方案或解决方案尚未保存！");
                }

                return dte.Solution.FullName;
            });

            string solutionDir = Path.GetDirectoryName(solutionFullName);
            if (string.IsNullOrEmpty(solutionDir) || !Directory.Exists(solutionDir))
            {
                throw new Exception("无法定位当前解决方案目录！");
            }

            CurrentSolutionDirectory = solutionDir;
            return solutionDir;
        }

        private static IEnumerable<string> SafeEnumerateFiles(string rootPath, string searchPattern)
        {
            if (string.IsNullOrEmpty(rootPath) || !Directory.Exists(rootPath))
            {
                return Enumerable.Empty<string>();
            }

            try
            {
                return Directory.EnumerateFiles(rootPath, searchPattern, SearchOption.AllDirectories);
            }
            catch
            {
                return Enumerable.Empty<string>();
            }
        }

        private static bool IsProjectSourcePath(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return false;
            }

            string normalizedPath = filePath.Replace('/', '\\');

            if (normalizedPath.IndexOf(@"\CustomConfig\Modules\", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            if (normalizedPath.IndexOf(@"\Config\Modules\", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            if (normalizedPath.IndexOf(@"\Repository\", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            if (normalizedPath.IndexOf(@"\bin\", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            if (normalizedPath.IndexOf(@"\obj\", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            if (normalizedPath.EndsWith(".bak", StringComparison.OrdinalIgnoreCase)) return false;

            return true;
        }

        private static string GetModuleGuidFromTmcFile(string tmcFilePath)
        {
            Console.WriteLine($"→ 尝试读取TMC文件：{Path.GetFullPath(tmcFilePath)}");

            if (!File.Exists(tmcFilePath))
            {
                throw new FileNotFoundException("TMC文件不存在", tmcFilePath);
            }

            XDocument tmcDoc = XDocument.Load(tmcFilePath);

            bool hasModulesContainer = tmcDoc.Descendants().Any(x => x.Name.LocalName == "Modules");
            bool hasConcreteModuleEntries = tmcDoc.Descendants().Any(x =>
                x.Name.LocalName == "Modules" &&
                x.Elements().Any(e => e.Name.LocalName == "Module"));

            XElement moduleElement = tmcDoc.Descendants().FirstOrDefault(x => x.Name.LocalName == "Module");
            string guidValue = moduleElement?.Attributes().FirstOrDefault(x => x.Name.LocalName == "GUID")?.Value;
            if (TryNormalizeGuid(guidValue, out string normalizedFromModule))
            {
                Console.WriteLine("→ 通过 <Module GUID=...> 提取到模块GUID。");
                return normalizedFromModule;
            }

            string[] preferredAttributeNames =
            {
                "GUID", "Guid", "ClassId", "ClassID", "ModuleId", "ModuleID",
                "TypeGuid", "TypeGUID", "ObjectId", "ObjectID", "Uuid", "UUID"
            };

            string[] preferredElementNames =
            {
                "GUID", "Guid", "ClassId", "ClassID", "ModuleId", "ModuleID",
                "TypeGuid", "TypeGUID", "ObjectId", "ObjectID", "Uuid", "UUID"
            };

            foreach (XElement element in tmcDoc.Descendants())
            {
                bool moduleLike = element.Name.LocalName.IndexOf("Module", StringComparison.OrdinalIgnoreCase) >= 0
                    || element.Name.LocalName.IndexOf("Class", StringComparison.OrdinalIgnoreCase) >= 0
                    || element.Name.LocalName.IndexOf("TcCom", StringComparison.OrdinalIgnoreCase) >= 0;

                if (!moduleLike)
                {
                    continue;
                }

                foreach (string attrName in preferredAttributeNames)
                {
                    string attrValue = element.Attributes().FirstOrDefault(a =>
                        string.Equals(a.Name.LocalName, attrName, StringComparison.OrdinalIgnoreCase))?.Value;

                    if (TryNormalizeGuid(attrValue, out string normalized))
                    {
                        Console.WriteLine($"→ 通过模块相关节点属性 {attrName} 提取到模块GUID。");
                        return normalized;
                    }
                }
            }

            foreach (XElement element in tmcDoc.Descendants())
            {
                if (!preferredElementNames.Any(name => string.Equals(name, element.Name.LocalName, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                if (TryNormalizeGuid(element.Value, out string normalized))
                {
                    Console.WriteLine($"→ 通过节点 <{element.Name.LocalName}> 提取到模块GUID。");
                    return normalized;
                }
            }

            string xmlText = tmcDoc.ToString(SaveOptions.DisableFormatting);
            Match match = Regex.Match(xmlText, @"\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?");
            if (match.Success && TryNormalizeGuid(match.Value, out string normalizedFromText))
            {
                Console.WriteLine("→ 通过TMC文本回退匹配提取到GUID。");
                return normalizedFromText;
            }

            if (hasModulesContainer && !hasConcreteModuleEntries)
            {
                throw new Exception("TMC文件仅包含空的 <Modules/> 壳结构，未生成任何真实模块定义，无法提取模块GUID。");
            }

            throw new Exception("TMC文件中未找到可用的模块GUID（未匹配到 Module/GUID、ClassId、UUID 等字段）。");
        }

        private static bool TryNormalizeGuid(string rawValue, out string normalizedGuid)
        {
            normalizedGuid = null;
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return false;
            }

            string cleaned = rawValue.Trim();
            if (Guid.TryParse(cleaned, out Guid guid))
            {
                normalizedGuid = guid.ToString("B").ToUpperInvariant();
                return true;
            }

            return false;
        }

        private static string GetFinalTmcFilePath()
        {
            Console.WriteLine("\n→ 正在自动搜索已发布的 TMC 文件...");

            for (int attempt = 1; attempt <= 20; attempt++)
            {
                List<string> allCandidates = CollectPublishedTmcCandidates();

                var rankedCandidates = allCandidates
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Select(path => new
                    {
                        Path = path,
                        HasUsableModule = HasUsableModuleDefinition(path),
                        LastWriteTime = SafeGetLastWriteTime(path)
                    })
                    .OrderByDescending(x => x.HasUsableModule)
                    .ThenByDescending(x => x.LastWriteTime)
                    .ToList();

                var bestCandidate = rankedCandidates.FirstOrDefault();
                if (bestCandidate != null)
                {
                    Console.WriteLine($"→ 自动找到 TMC 文件：{bestCandidate.Path}");
                    if (!bestCandidate.HasUsableModule)
                    {
                        Console.WriteLine("× 该TMC看起来仍是工程壳文件（未发现真实模块定义），后续可能无法提取GUID。");
                    }
                    return bestCandidate.Path;
                }

                if (attempt < 20)
                {
                    System.Threading.Thread.Sleep(1000);
                }
            }

            throw new Exception("未自动定位到可用的已发布 TMC 文件（已扫描 _products/TcPublish、TwinCAT Config/Repository）。");
        }

        private static string TryGetPublishedTmcFilePath(DateTime notOlderThan)
        {
            return CollectPublishedTmcCandidates()
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(path => HasUsableModuleDefinition(path))
                .Where(path => SafeGetLastWriteTime(path) >= notOlderThan.AddSeconds(-2))
                .OrderByDescending(SafeGetLastWriteTime)
                .FirstOrDefault();
        }

        private static List<string> CollectPublishedTmcCandidates()
        {
            List<string> candidates = new List<string>();

            if (!string.IsNullOrEmpty(CurrentSolutionDirectory))
            {
                string localPublishRoot = Path.Combine(CurrentSolutionDirectory, CurrentCppProjectName, "_products", "TcPublish");
                if (Directory.Exists(localPublishRoot))
                {
                    candidates.AddRange(SafeEnumerateFiles(localPublishRoot, "*.tmc"));
                }
            }

            foreach (string root in DefaultTmcBasePaths)
            {
                try
                {
                    if (!Directory.Exists(root)) continue;
                    candidates.AddRange(Directory.EnumerateFiles(root, "*.tmc", SearchOption.AllDirectories)
                        .Where(x => x.IndexOf(CurrentCppProjectName, StringComparison.OrdinalIgnoreCase) >= 0));
                }
                catch
                {
                }
            }

            foreach (string root in TwinCatRepositoryRoots)
            {
                try
                {
                    if (!Directory.Exists(root)) continue;
                    candidates.AddRange(Directory.EnumerateFiles(root, "*.tmc", SearchOption.AllDirectories)
                        .Where(x => x.IndexOf(CurrentCppProjectName, StringComparison.OrdinalIgnoreCase) >= 0));
                }
                catch
                {
                }
            }

            return candidates;
        }

        private static bool HasUsableModuleDefinition(string tmcFilePath)
        {
            try
            {
                if (!File.Exists(tmcFilePath))
                {
                    return false;
                }

                XDocument doc = XDocument.Load(tmcFilePath);

                bool hasModuleNode = doc.Descendants().Any(x => x.Name.LocalName == "Module");
                bool hasNonEmptyModules = doc.Descendants().Any(x =>
                    x.Name.LocalName == "Modules" &&
                    x.Elements().Any(e => e.Name.LocalName == "Module"));

                bool hasGuidLikeAttribute = doc.Descendants().Any(x =>
                    x.Attributes().Any(a =>
                        a.Name.LocalName.IndexOf("Guid", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        a.Name.LocalName.IndexOf("ClassId", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        a.Name.LocalName.IndexOf("ModuleId", StringComparison.OrdinalIgnoreCase) >= 0));

                return hasNonEmptyModules || (hasModuleNode && hasGuidLikeAttribute);
            }
            catch
            {
                return false;
            }
        }

        private static DateTime SafeGetLastWriteTime(string path)
        {
            try
            {
                return File.GetLastWriteTime(path);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        // “Publish Modules” 只是让 TwinCAT 识别模块类型；
        // 真正要在当前配置树里看到可实例化对象，还需要再往 TcCOM Objects 下创建实例。
        // 这里使用最终 TMC 中解析出的 GUID，确保挂进去的是发布后的真实模块类型。
        private static bool AddTcComObject(ITcSysManager sysManager, ITcSmTreeItem cppProject)
        {
            ITcSmTreeItem tcComObjects = RetryComCall(() => sysManager.LookupTreeItem("TIRC^TcCOM Objects"));
            if (tcComObjects == null)
            {
                throw new Exception("未找到TcCOM Objects节点（TIRC^TcCOM Objects）！");
            }

            ITcSmTreeItem newTcComObject = null;
            try
            {
                string finalTmcPath = GetFinalTmcFilePath();
                string realModuleGuid = GetModuleGuidFromTmcFile(finalTmcPath);
                Console.WriteLine($"√  从TMC文件读取到模块GUID：{realModuleGuid}");

                newTcComObject = RetryComCall(() => tcComObjects.CreateChild("CustomTcComModule", 0, "", realModuleGuid));
                if (newTcComObject == null)
                {
                    throw new Exception("添加TcCOM Object失败！请确认GUID对应的模块已被TwinCAT检测到。");
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"× 添加 TcCOM Object 失败：{ex.Message}");
                return false;
            }
            finally
            {
                ReleaseComIfNeeded(newTcComObject);
                ReleaseComIfNeeded(tcComObjects);
            }
        }

        // 与其让 TwinCAT wizard 在同名冲突时给出不稳定的错误，
        // 不如先在文件系统层面生成一个确定不冲突的工程名。
        private static string GetUniqueCppProjectName(string solutionDir, string baseName)
        {
            if (string.IsNullOrWhiteSpace(solutionDir) || !Directory.Exists(solutionDir))
            {
                return baseName;
            }

            string candidate = baseName;
            int index = 1;

            while (Directory.Exists(Path.Combine(solutionDir, candidate)) ||
                   SafeEnumerateFiles(solutionDir, candidate + ".vcxproj").Any())
            {
                candidate = baseName + "_" + index;
                index++;
            }

            return candidate;
        }

        // 模块名同样提前避让，避免 .cpp / .h / .tmc 等文件撞名后把问题伪装成别的异常。
        private static string GetUniqueModuleName(string solutionDir, string cppProjectName, string baseName)
        {
            string projectDir = Path.Combine(solutionDir, cppProjectName);
            if (!Directory.Exists(projectDir))
            {
                return baseName;
            }

            string candidate = baseName;
            int index = 1;

            while (SafeEnumerateFiles(projectDir, candidate + ".cpp").Any() ||
                   SafeEnumerateFiles(projectDir, candidate + ".h").Any() ||
                   SafeEnumerateFiles(projectDir, candidate + ".hpp").Any() ||
                   SafeEnumerateFiles(projectDir, candidate + ".tmc").Any())
            {
                candidate = baseName + "_" + index;
                index++;
            }

            return candidate;
        }

        private static void RetryComCall(Action action, int maxRetryCount = 60, int retryDelayMs = 500)
        {
            RetryComCall<object>(() =>
            {
                action();
                return null;
            }, maxRetryCount, retryDelayMs);
        }

        private static T RetryComCall<T>(Func<T> func, int maxRetryCount = 60, int retryDelayMs = 500)
        {
            Exception lastException = null;

            for (int i = 0; i < maxRetryCount; i++)
            {
                try
                {
                    return func();
                }
                catch (COMException ex) when (
                    ex.ErrorCode == RPC_E_CALL_REJECTED ||
                    ex.ErrorCode == RPC_E_SERVERCALL_RETRYLATER ||
                    ex.ErrorCode == RPC_S_CALL_FAILED)
                {
                    lastException = ex;
                    System.Threading.Thread.Sleep(retryDelayMs);
                }
            }

            if (lastException != null)
            {
                throw lastException;
            }

            throw new Exception("COM调用重试失败。");
        }

        private static void ReleaseComIfNeeded(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch
                {
                }
            }
        }

        [ComImport]
        [Guid("00000016-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleMessageFilter
        {
            [PreserveSig]
            int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

            [PreserveSig]
            int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

            [PreserveSig]
            int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
        }

        private sealed class MessageFilter : IOleMessageFilter
        {
            public static void Register()
            {
                CoRegisterMessageFilter(new MessageFilter(), out IOleMessageFilter oldFilter);
            }

            public static void Revoke()
            {
                CoRegisterMessageFilter(null, out IOleMessageFilter oldFilter);
            }

            int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
            {
                return 0;
            }

            int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
            {
                if (dwRejectType == 2)
                {
                    return 250;
                }

                return -1;
            }

            int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
            {
                return 2;
            }

            [DllImport("Ole32.dll")]
            private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        #endregion
    }
}
