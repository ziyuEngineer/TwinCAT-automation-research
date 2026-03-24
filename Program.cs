
using EnvDTE;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Xml.Linq;
using TCatSysManagerLib;

namespace TwinCATCppProjectCreator
{
    internal class Program
    {
        // 限定哪些vs的版本可以被操作
        private static readonly List<string> SupportedVsDteVersions = new List<string>
        {
            "VisualStudio.DTE.15.0",  // VS2017
            "VisualStudio.DTE.16.0",  // VS2019
            "VisualStudio.DTE.17.0"   // VS2022
        };


        // 模块模板名称与对应的Wizard标识映射
        //WizardId 是 TwinCAT Automation Interface 在 CreateChild(...) 时要用于向导/模板来创建子项”的标识字符串。
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

        // C++项目模板选项
        private static readonly Dictionary<int, (string TemplateName, string WizardId)> ProjectTemplates = new Dictionary<int, (string, string)>
        {
            {1, ("Versioned C++ projects", "TcVersionedDriverWizard")},
            {2, ("Driver C++ project", "TcDriverWizard")}
        };

        private const string DEFAULT_CPP_PROJECT_NAME = "NewCppProject";
        private const string DEFAULT_MODULE_NAME = "NewModule";

        // TMC文件默认基础路径
        // TwinCAT 要把一个 TcCOM 模块“识别成可添加到系统配置里的模块”，并不只是看你项目目录里有没有源码，而是要看 发布后的模块包 是否在它能扫描到的模块目录里。 
        // Beckhoff 官方文档明确说，导出的/发布的模块文件 会被复制到 %TwinCAT3Dir%\CustomConfig\Modules\<MODULENAME>\；
        // 而在添加现有 TcCOM 模块时，TwinCAT 需要能在这个 publish directory 里检测到它。 
        // 将来这段代码想做的是“Add TcCOM Object 到 TwinCAT 配置” 
        private const string DEFAULT_TMC_BASE_PATH = @"C:\TwinCAT\3.1\CustomConfig\Modules";
        private const string TMC_RELATIVE_PATH = @"\NewCppProject\NewCppProject.tmc";
        // 超时时间常量（5秒）
        private const int DEFAULT_TIMEOUT = 5000;
        private const int CHECK_INTERVAL = 100;
        private const int DEFAULT_FILE_WAIT_TIMEOUT = 60000;

        private static string CurrentSolutionDirectory = null;
        private const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001);
        private const int RPC_E_SERVERCALL_RETRYLATER = unchecked((int)0x8001010A);

        [STAThread]
        static void Main(string[] args)
        {
            DTE selectedDte = null; //DTE 是 Visual Studio 自动化模型里的顶层对象。可以把它理解成整个 VS IDE 的自动化入口。
            ITcSysManager sysManager = null; //是 TwinCAT Automation Interface 的主接口,主工程的控制台，用来对 TwinCAT 3 XAE 做基本配置操作
            ITcSmTreeItem cppProject = null; //工程树上某个具体节点的入口,管树上的具体对象 比如 C++ 节点，I/O Devices 节点
                                             //拿到某个 ITcSmTreeItem 后，你就可以对这个节点做事，如： 查它的子节点.在它下面 CreateChild. .改属性. 导入导出配置

            MessageFilter.Register();

            try
            {
                bool tmcPatched = false;
                bool tmcCodeGenerated = false;

                // ========== 步骤1：仅绑定Visual Studio项目 ==========
                Console.WriteLine("========================================");
                Console.WriteLine("===== 绑定Visual Studio项目 =====");
                selectedDte = SelectAndBindVsProject(); //选择现在打开的vs 项目
                if (selectedDte == null)
                {
                    Console.WriteLine("× 你选择取消绑定，程序即将退出...");
                    return;
                }
                Console.WriteLine("√  项目绑定成功！");

                // 获取SysManager
                // 从当前 VS 解决方案中的 TwinCAT 项目里拿到 ITcSysManager，用来操作 TwinCAT 配置树
                sysManager = ConnectToBoundTcProject(selectedDte);
                //获取当前的目录，后面找 .tmc 和 .cpp 时会用到
                CurrentSolutionDirectory = GetSolutionDirectory(selectedDte);

                // ========== 步骤2：选择并新建C++项目（5秒默认Y） ==========
                Console.WriteLine("\n========================================");
                Console.WriteLine("是否新建TwinCAT C++项目？(Y/N)（5秒无操作默认Y）：");
                string createCppProjInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT* DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{createCppProjInput}");

                if (createCppProjInput != "Y")
                {
                    Console.WriteLine("× 你选择不创建C++项目，程序即将退出...");
                    return;
                }

                // ========== 步骤3：选择C++项目模板（5秒默认2） ==========
                Console.WriteLine("\n===== 选择C++项目模板 =====");
                foreach (var item in ProjectTemplates)
                {
                    Console.WriteLine($"{item.Key} : {item.Value.TemplateName}");
                }
                Console.WriteLine("请输入模板编号(注:4026仅支持Versioned Project)（5秒无操作默认2）：");
                int selectedProjTemplateId = GetNumberInputWithTimeout(2, ProjectTemplates.Keys, DEFAULT_TIMEOUT* DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择的模板编号：{selectedProjTemplateId}");

                var selectedProjTemplate = ProjectTemplates[selectedProjTemplateId];
                cppProject = CreateTcCppProject(sysManager, selectedProjTemplate.WizardId);//创建C++ 项目
                Console.WriteLine($"√  {selectedProjTemplate.TemplateName}创建完成！");

                // ========== 步骤4：选择并创建模块（5秒默认3） ==========
                Console.WriteLine("\n========================================");
                Console.WriteLine("===== 选择C++模块模板 =====");
                foreach (var item in ModuleTemplates)
                {
                    Console.WriteLine($"{item.Key} : {item.Value.TemplateName}");
                }
                Console.WriteLine("请输入模板编号（5秒无操作默认3）：");
                int selectedTemplateId = GetNumberInputWithTimeout(3, ModuleTemplates.Keys, DEFAULT_TIMEOUT * DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择的模板编号：{selectedTemplateId}");

                var selectedTemplate = ModuleTemplates[selectedTemplateId];
                CreateTcCppModule(cppProject, selectedTemplate.WizardId); //创建C++ 模块
                Console.WriteLine($"√  模块「{selectedTemplate.TemplateName}」创建操作已完成，请在VS中验证结果。");

                SaveAll(selectedDte);
                WaitForProjectArtifacts(selectedDte);
                System.Threading.Thread.Sleep(1000);

                // ========== 新增步骤：修改工程TMC（5秒默认Y） ==========
                Console.WriteLine("\n========================================");
                Console.WriteLine("是否修改工程TMC（保守方式：修改向导默认结构体参数为 Gain / Enable / VelocityLimit）？(Y/N)（5秒无操作默认Y）：");
                string patchTmcInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT * DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{patchTmcInput}");

                if (patchTmcInput == "Y")
                {
                    string patchedTmcPath = PatchProjectTmc(selectedDte); //修改工程TMC
                    Console.WriteLine($"√  工程TMC修改成功！→ {patchedTmcPath}");
                    tmcPatched = true;
                }
                else
                {
                    Console.WriteLine("× 跳过修改工程TMC...");
                }

                // ========== 步骤5：TMC Code Generator（5秒默认Y） ==========
                Console.WriteLine("\n========================================");
                Console.WriteLine("是否启用TMC Code Generator？(Y/N)（5秒无操作默认Y）：");
                string tmcGenerateInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT * DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{tmcGenerateInput}");

                if (tmcGenerateInput == "Y")
                {
                    ExecuteTmcCodeGenerator(cppProject); //生成TMC
                    SaveAll(selectedDte);
                    System.Threading.Thread.Sleep(1500);
                    Console.WriteLine("√  TMC Code Generator执行成功！");
                    tmcCodeGenerated = true;
                }
                else
                {
                    Console.WriteLine("× 跳过TMC Code Generator...");
                }

                // ========== 新增步骤：往生成的C++模块里写一点点代码（5秒默认Y） ==========
                Console.WriteLine("\n========================================");
                Console.WriteLine("是否往生成的C++模块里写一点点简单代码？(Y/N)（5秒无操作默认Y）：");
                string writeCppInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT * DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{writeCppInput}");

                if (writeCppInput == "Y")
                {
                    WriteSimpleCodeToGeneratedModule(selectedDte); //自动添加c++代码
                    Console.WriteLine("√  已向生成的C++模块写入简单示例代码！");
                }
                else
                {
                    Console.WriteLine("× 跳过自动写入模块源码...");
                }

                // ========== 步骤6：发布TcCOM Modules（5秒默认Y+3秒延迟） ==========
                Console.WriteLine("\n========================================");
                Console.WriteLine("是否发布TcCOM Objects？(Y/N)（5秒无操作默认Y）：");
                string publishModulesInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT * DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{publishModulesInput}");

                if (publishModulesInput == "Y")
                {
                    ExecutePublishModules(cppProject);//发布TcCOM Modules,即把当前 C++ 工程产出的模块能被TwinCAT 登记、被识别。
                    Console.WriteLine("√  TcCOM Modules发布成功！");
                }
                else
                {
                    Console.WriteLine("× 跳过发布TcCOM Objects...");
                }

                // ========== 步骤7：添加TcCOM Object（5秒默认Y） ==========
                Console.WriteLine("\n========================================");
                Console.WriteLine("是否添加TcCOM Object到配置？(Y/N)（5秒无操作默认Y）：");
                string addTcComInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT * DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{addTcComInput}");

                if (addTcComInput == "Y")
                {
                    AddTcComObject(sysManager, cppProject); //在当前 TwinCAT 配置树里，真正放进去一个模块实例。
                    Console.WriteLine("√  TcCOM Object添加成功！");
                }
                else
                {
                    Console.WriteLine("× 跳过添加TcCOM Object...");
                }

                // ========== 步骤8：是否编译当前项目（最后第二步，5秒默认Y） ==========
                Console.WriteLine("\n========================================");
                Console.WriteLine("是否编译当前项目？(Y/N)（5秒无操作默认Y）：");
                string buildProjectInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT * DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{buildProjectInput}");

                if (buildProjectInput == "Y")
                {
                    BuildCurrentSolution(selectedDte); // vs 自己的接口来SolutionBuild
                    Console.WriteLine("√  当前项目编译完成！");
                }
                else
                {
                    Console.WriteLine("× 跳过编译当前项目...");
                }

                // ========== 步骤9：是否激活 TwinCAT 配置（Active Configuration，最后一步，5秒默认Y） ==========
                Console.WriteLine("\n========================================");
                Console.WriteLine("是否激活 TwinCAT 配置（Active Configuration）？(Y/N)（5秒无操作默认Y）：");
                string activeConfigInput = GetYesNoInputWithTimeout("Y", DEFAULT_TIMEOUT * DEFAULT_TIMEOUT);
                Console.WriteLine($"→ 选择结果：{activeConfigInput}");

                if (activeConfigInput == "Y")
                {
                    ActivateTcConfiguration(selectedDte, sysManager);
                }
                else
                {
                    Console.WriteLine("× 跳过激活 TwinCAT 配置...");
                }

                // ========== 结束 ==========
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

                // 释放所有COM资源
                if (cppProject != null) Marshal.ReleaseComObject(cppProject);
                if (sysManager != null) Marshal.ReleaseComObject(sysManager);
                if (selectedDte != null) Marshal.ReleaseComObject(selectedDte);
            }

            Console.WriteLine("\n按任意键关闭窗口...");
            Console.ReadKey();
        }

        #region 通用超时输入处理方法
        /// <summary>
        /// 通用Y/N输入处理，支持超时默认值
        /// </summary>
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

        /// <summary>
        /// 通用数字选择处理，支持超时默认值和合法值校验
        /// </summary>
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
                    else
                    {
                        Console.WriteLine($"× 输入无效，使用默认值：{defaultValue}");
                        return defaultValue;
                    }
                }
                System.Threading.Thread.Sleep(CHECK_INTERVAL);
                remainingTime -= CHECK_INTERVAL;
            }
            return defaultValue;
        }

        /// <summary>
        /// 清理控制台输入缓冲区，避免回车等残留按键影响下一步选择
        /// </summary>
        private static void ClearConsoleInputBuffer()
        {
            while (Console.KeyAvailable)
            {
                Console.ReadKey(true);
            }
        }
        #endregion

        #region 核心方法实现
        /// <summary>
        /// 仅选择并绑定运行中的Visual Studio实例（排除TcXaeShell）
        /// </summary>
        private static DTE SelectAndBindVsProject()
        {
            var dteInstances = GetRunningVsInstances(); //调 Windows API 把当前所有运行中的 VS DTE 对象找出来。
            if (dteInstances.Count == 0)
            {
                throw new Exception("当前没有运行的Visual Studio实例！");
            }

            Console.WriteLine("请选择项目编号（0=取消）：");
            foreach (var item in dteInstances)
            {
                string projectName = RetryComCall(() => item.Value.Solution != null && item.Value.Solution.Projects.Count > 0
                    ? item.Value.Solution.Projects.Item(1).Name
                    : "未加载项目");
                Console.WriteLine($"{item.Key} : {projectName} (Visual Studio实例)");
            }

            Console.Write("请输入编号：");
            string input = Console.ReadLine()?.Trim();
            if (input == "0") return null;
            if (!int.TryParse(input, out int selectedIndex) || !dteInstances.ContainsKey(selectedIndex))
            {
                throw new Exception("输入的编号无效！");
            }

            return dteInstances[selectedIndex];  //返回用户选中的那个 DTE。
        }

        /// <summary>
        /// 获取所有运行的Visual Studio实例（仅匹配DTE_PROG_ID，排除TcXaeShell）
        /// </summary>
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

            IntPtr fetchedPtr = System.Runtime.InteropServices.Marshal.AllocHGlobal(System.Runtime.InteropServices.Marshal.SizeOf(typeof(int)));
            try
            {
                while (enumMoniker.Next(1, monikers, fetchedPtr) == 0)
                {
                    int instanceCount = System.Runtime.InteropServices.Marshal.ReadInt32(fetchedPtr);
                    if (instanceCount != 1) break;

                    IBindCtx bindCtx;
                    CreateBindCtx(0, out bindCtx);
                    string displayName;
                    monikers[0].GetDisplayName(bindCtx, null, out displayName);

                    // 关键修改：匹配VS2017(15.0)/2019(16.0)/2022(17.0)，排除TcXaeShell
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

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(monikers[0]);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(bindCtx);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FreeHGlobal(fetchedPtr);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(enumMoniker);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rot);
            return dteInstances;
        }

        /// <summary>
        /// 连接到绑定的TwinCAT项目
        /// </summary>
        private static ITcSysManager ConnectToBoundTcProject(DTE dte)
        {
            Console.WriteLine($"√  已连接到Visual Studio实例：{RetryComCall(() => dte.Name)} v{RetryComCall(() => dte.Version)}");

            if (RetryComCall(() => dte.Solution) == null || RetryComCall(() => dte.Solution.Projects.Count) == 0)
            {
                throw new Exception("当前实例未加载任何项目！");
            }

            Project tcProject = RetryComCall(() => dte.Solution.Projects.Item(1));
            Console.WriteLine($"√  已定位到TwinCAT项目：{tcProject.Name}");

            // Project.Object 返回的是该项目背后的 COM 自动化对象；对 TwinCAT 项目来说，这个对象可以转成 ITcSysManager
            return (ITcSysManager)tcProject.Object;
        }

        /// <summary>
        /// 编译当前解决方案（等待编译结束）
        /// </summary>
        private static void BuildCurrentSolution(DTE dte)
        {
            if (dte?.Solution == null)
            {
                throw new Exception("当前未加载解决方案！");
            }
            //vs 自己的接口来SolutionBuild
            SolutionBuild sb = RetryComCall(() => dte.Solution.SolutionBuild);
            RetryComCall(() => sb.Build(true));

            // 等待编译结束（轮询 BuildState）
            const int buildWaitTimeoutMs = 300000; // 最多等 5 分钟
            int waited = 0;
            while (waited < buildWaitTimeoutMs)
            {
                if (RetryComCall(() => sb.BuildState) != vsBuildState.vsBuildStateInProgress)
                    break;
                System.Threading.Thread.Sleep(500);
                waited += 500;
            }

            bool buildOk = (RetryComCall(() => sb.LastBuildInfo) == 0 && RetryComCall(() => sb.BuildState) == vsBuildState.vsBuildStateDone);
            if (buildOk)
            {
                Console.WriteLine("→ 编译成功。");
            }
            else
            {
                Console.WriteLine("→ 编译已结束，若失败请查看 Visual Studio 错误列表（Error List）。");
            }
        }

        /// <summary>
        /// 激活 TwinCAT 配置：先保存当前配置，再通过 DTE 执行“激活配置”命令或 COM ActivateConfiguration
        /// </summary>
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
                // 1) 先保存当前配置，否则激活的可能是旧状态
                Console.WriteLine("→ 正在保存当前 TwinCAT 配置...");
                sysManager.SaveConfiguration(configPath);
                Console.WriteLine($"→ 配置已保存：{configPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"× 保存配置时提示：{ex.Message}（继续尝试激活）");
            }

            bool activated = false;

            // 2) 优先用 DTE 执行 TwinCAT 菜单“激活配置”，与 IDE 里点击效果一致
            string[] tryCommands = { "TwinCAT.ActivateConfiguration", "TcXaeShell.TwinCAT.ActivateConfiguration", "Build.ActivateConfiguration" };
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
                    /* 命令不存在或不可用时忽略，尝试下一个 */
                }
            }

            // 3) 若 DTE 命令都不可用，再调用 COM ActivateConfiguration
            if (!activated)
            {
                try
                {
                    Console.WriteLine("→ 正在通过 System Manager 激活配置...");
                    sysManager.ActivateConfiguration();
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

        /// <summary>
        /// 显示当前 VS Solution Active Configuration 并可选切换（供需要时使用）
        /// </summary>
        private static void ShowAndSelectActiveConfig(DTE dte)
        {
            if (dte?.Solution == null)
            {
                throw new Exception("当前未加载解决方案！");
            }

            SolutionBuild sb = dte.Solution.SolutionBuild;
            SolutionConfiguration activeConfig = sb.ActiveConfiguration;
            string currentName = activeConfig?.Name ?? "未知";
            Console.WriteLine($"→ 当前 Active Configuration：{currentName}");

            SolutionConfigurations configs = sb.SolutionConfigurations;
            if (configs == null || configs.Count == 0)
            {
                Console.WriteLine("× 未找到可用的 Solution Configuration。");
                return;
            }

            var configList = new List<string>();
            for (int i = 1; i <= configs.Count; i++)
            {
                try
                {
                    SolutionConfiguration c = configs.Item(i);
                    if (c != null)
                    {
                        string name = c.Name;
                        configList.Add(name);
                        string mark = (name == currentName) ? " [当前]" : "";
                        Console.WriteLine($"  {i} : {name}{mark}");
                    }
                }
                catch { /* 忽略无效项 */ }
            }

            if (configList.Count == 0)
            {
                Console.WriteLine("× 无法枚举 Configuration。");
                return;
            }

            Console.WriteLine("请输入要切换的编号（0=不切换，保持当前）：");
            string input = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(input) || input == "0")
            {
                Console.WriteLine("× 未切换 Active Configuration。");
                return;
            }

            if (!int.TryParse(input, out int idx) || idx < 1 || idx > configList.Count)
            {
                Console.WriteLine("× 编号无效，未切换。");
                return;
            }

            string targetName = configList[idx - 1];
            if (targetName == currentName)
            {
                Console.WriteLine("→ 已是当前配置，无需切换。");
                return;
            }

            try
            {
                SolutionConfiguration targetConfig = configs.Item(idx);
                if (targetConfig != null)
                {
                    // SolutionConfiguration2.Activate() 需 EnvDTE80，此处用 dynamic 兼容
                    ((dynamic)targetConfig).Activate();
                    Console.WriteLine($"√  已切换 Active Configuration 为：{targetName}");
                }
                else
                {
                    Console.WriteLine("× 无法激活该 Configuration。");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"× 切换失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 创建指定模板的TwinCAT C++项目
        /// </summary>
        private static ITcSmTreeItem CreateTcCppProject(ITcSysManager sysManager, string templateId)
        {
            // 定位TIXC（C++节点）
            ITcSmTreeItem cppNode = sysManager.LookupTreeItem("TIXC"); //去 TwinCAT 配置树里找到 C++ 节点
            //LookupTreeItem 按路径名找树节点,返回 ITcSmTreeItem
            //为什么这样拿：想在 C++ 节点下面创建新项目
            //拿到后做什么：对这个节点调用 CreateChild

            if (cppNode == null)
            {
                throw new Exception("未找到TwinCAT C++节点（TIXC），请确认项目支持C++开发！");
            }

            // 创建C++项目（根据选择的模板）
            ITcSmTreeItem cppProject = cppNode.CreateChild(DEFAULT_CPP_PROJECT_NAME, 0, "", templateId);
            //子项名字："NewCppProject"
            //nSubtype=0：创建一个 C++ project 级别的子项
            //templateId 来自前面用户选的模板： var selectedProjTemplate 

            return cppProject;
        }

        /// <summary>
        /// 优化模块创建判断：不依赖返回值，仅捕获异常
        /// </summary>
        private static void CreateTcCppModule(ITcSmTreeItem cppProject, string wizardId)
        {
            try
            {
                // 执行创建模块操作
                ITcSmTreeItem cppModule = cppProject.CreateChild(DEFAULT_MODULE_NAME, 1, "", wizardId);
                //父节点 cppProject：上一步创建 C++ 项目时返回的 ITcSmTreeItem
                //子项名字 "NewModule"
                //subtype = 1: 创建一个模块子项
                //wizardId：来自用户选的ModuleTemplates 字典

            }
            catch (Exception ex)
            {
                throw new Exception($"创建模块时发生异常（模板：{wizardId}）：{ex.Message}");
            }
        }

        /// <summary>
        /// 对这个 C++ 项目执行 “StartTmcCodeGenerator” 这个方法。
        /// </summary>
        private static void ExecuteTmcCodeGenerator(ITcSmTreeItem cppProject)
        {
            // 这段 XML 不是普通配置数据，而是发给 TwinCAT 项目节点的一条“内部命令”。
            // 这里的意思是：对当前 C++ 项目执行 StartTmcCodeGenerator 方法，
            // 让 TwinCAT 自己启动 TMC Code Generator，生成/刷新模块描述相关内容。
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

            // TwinCAT Automation Interface 对很多高级操作并没有直接暴露成普通 C# 方法，
            // 而是要求通过 ConsumeXml(...) 喂一段约定格式的 XML，让节点自行解析并执行。
            // 所以这里本质上是在“命令 TwinCAT 项目节点执行 TMC 生成器”，
            // 而不是把 XML 当作普通文本写进某个文件。
            cppProject.ConsumeXml(tmcGeneratorXml);
        }

        /// <summary>
        /// 执行Publish Modules命令（含3秒延迟）
        /// </summary>
        /// 模块一开始在你的VS项目目录，而Publish Modules 做的事情 编译模块 生成 dll tmc， 复制到模块库 C:\TwinCAT\3.1\Config\Modules
        /// 这样TwinCAT 才能知道这个 GUID， 才能在模块库里发现这个模块。
        private static void ExecutePublishModules(ITcSmTreeItem cppProject)
        {
            // 这段 XML 同样是一条发给 TwinCAT 项目节点的内部命令。
            // 这里要求 TwinCAT 对当前 C++ 项目执行 PublishModules 方法，
            // 也就是把生成好的模块发布到 TwinCAT 可识别的模块目录中。
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

            // 通过 ConsumeXml(...) 让 ITcSmTreeItem 自己执行“发布模块”动作；
            cppProject.ConsumeXml(publishModulesXml);
            Console.WriteLine("→ 等待3秒，确保发布文件同步完成...");
            System.Threading.Thread.Sleep(3000);
        }

        /// <summary>
        /// 保存当前Solution里的文档，尽量让向导生成的文件先落盘
        /// </summary>
        private static void SaveAll(DTE dte)
        {
            if (dte == null)
            {
                return;
            }

            try
            {
                RetryComCall(() => dte.ExecuteCommand("File.SaveAll"));
            }
            catch
            {
                // 某些情况下命令不可用时忽略
            }

            try
            {
                RetryComCall(() =>
                {
                    if (dte.Documents != null)
                    {
                        dte.Documents.SaveAll();
                    }
                });
            }
            catch
            {
                // 忽略保存异常
            }
        }

        /// <summary>
        /// 等待工程TMC和模块源码真正生成到磁盘
        /// </summary>
        private static void WaitForProjectArtifacts(DTE dte)
        {
            DateTime endTime = DateTime.Now.AddMilliseconds(DEFAULT_FILE_WAIT_TIMEOUT);

            while (DateTime.Now < endTime)
            {
                string tmcFilePath = TryGetProjectTmcFilePath(dte);
                string cppFilePath = TryGetGeneratedModuleCppFilePath(dte);

                if (!string.IsNullOrEmpty(tmcFilePath) && !string.IsNullOrEmpty(cppFilePath))
                {
                    System.Threading.Thread.Sleep(1000);
                    return;
                }

                System.Threading.Thread.Sleep(500);
            }

            throw new Exception("等待工程TMC/模块源码生成超时，请确认Visual Studio已完成向导文件落盘。");
        }

        // 找到工程里的 xml 文件，按 XML 结构去定位参数相关节点，然后修改参数
        private static string PatchProjectTmc(DTE dte)
        {
            // 先定位当前工程目录下的 TMC 文件
            string tmcFilePath = GetProjectTmcFilePath(dte);
            Console.WriteLine($"→ 准备修改工程TMC：{tmcFilePath}");

            // 如果工程里连 TMC 文件都没有，就没法继续做参数补丁
            if (!File.Exists(tmcFilePath))
            {
                throw new FileNotFoundException("工程TMC文件不存在", tmcFilePath);
            }

            // 先做一个 .bak 备份，避免后面修改失败时把原始 TMC 覆盖掉
            string backupPath = tmcFilePath + ".bak";
            File.Copy(tmcFilePath, backupPath, true);

            // 把 TMC 当作 XML 文档读进来；保留原有空白，尽量少破坏文件原样
            XDocument tmcDoc = XDocument.Load(tmcFilePath, LoadOptions.PreserveWhitespace);

            // 找模块定义节点；后面的参数区通常挂在 <Module> 下面
            XElement moduleElement = tmcDoc.Descendants().FirstOrDefault(x => x.Name.LocalName == "Module");
            if (moduleElement == null)
            {
                throw new Exception("工程TMC中未找到<Module>节点");
            }

            // 找模块的参数容器 <Parameters>；我们后面要在这里查找/修改/新增参数
            XElement parametersElement = moduleElement.Descendants().FirstOrDefault(x => x.Name.LocalName == "Parameters");
            if (parametersElement == null)
            {
                throw new Exception("工程TMC中未找到<Parameters>节点");
            }

            // 先把当前模块下已有的所有 <Parameter> 节点取出来，便于后面做判断
            List<XElement> parameterElements = parametersElement.Elements()
                .Where(x => x.Name.LocalName == "Parameter")
                .ToList();

            // 打印当前已有参数名，方便调试时快速判断模板里本来就带了什么参数
            string parameterNames = string.Join(", ", parameterElements.Select(x => GetChildElementValue(x, "Name")).Where(x => !string.IsNullOrEmpty(x)));
            Console.WriteLine($"→ 当前工程TMC里的参数：{parameterNames}");

            // patched 用来标记：是否已经成功找到一种可行方式完成修改
            bool patched = false;

            // 第一优先级：看看模块参数里是否已经有“内联结构体参数”
            // 判断标准：不是 TraceLevelMax，并且它下面至少有 3 个 SubItem
            // 如果有，就直接把这个现成结构体的前三个字段改成 Gain / Enable / VelocityLimit
            XElement inlineStructuredParameter = parameterElements.FirstOrDefault(x =>
                !string.Equals(GetChildElementValue(x, "Name"), "TraceLevelMax", StringComparison.OrdinalIgnoreCase) &&
                x.Elements().Count(e => e.Name.LocalName == "SubItem") >= 3);

            if (inlineStructuredParameter != null)
            {
                // 直接改这个现成参数节点内部的结构体字段，属于“最少改动”的做法
                PatchStructuredTypeElement(inlineStructuredParameter);
                patched = true;
                Console.WriteLine("→ 已找到模块里现成的结构体参数，直接改成 Gain / Enable / VelocityLimit。");
            }

            if (!patched)
            {
                // 第二优先级：如果参数节点本身不是结构体，就看看它引用的 Type 是否对应一个结构体 DataType
                XElement editableParameterElement = parameterElements.FirstOrDefault(x =>
                    !string.Equals(GetChildElementValue(x, "Name"), "TraceLevelMax", StringComparison.OrdinalIgnoreCase));

                if (editableParameterElement != null)
                {
                    // 先取这个参数声明里写的类型名，再去 DataType 区里找真正的结构体定义
                    string parameterTypeName = GetChildElementValue(editableParameterElement, "Type");
                    XElement referencedDataType = FindReferencedDataType(tmcDoc, parameterTypeName);

                    // 如果引用到的 DataType 的确是一个至少有 3 个 SubItem 的结构体，
                    // 那就改这个结构体定义本身，并把参数名统一改成 Parameter
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
                // 第三优先级：前两种都不行时，再在整个 TMC 的 DataType 区里兜底找一个候选结构体
                // 条件：名字里带 Parameter，且下面至少有 3 个 SubItem
                XElement fallbackStructuredDataType = tmcDoc.Descendants()
                    .FirstOrDefault(x =>
                        x.Name.LocalName == "DataType" &&
                        x.Elements().Count(e => e.Name.LocalName == "SubItem") >= 3 &&
                        (GetChildElementValue(x, "Name")?.IndexOf("Parameter", StringComparison.OrdinalIgnoreCase) >= 0));

                if (fallbackStructuredDataType != null)
                {
                    // 如果找到这种候选结构体，也直接把它改成目标字段结构
                    PatchStructuredTypeElement(fallbackStructuredDataType);
                    patched = true;
                    Console.WriteLine("→ 已找到候选结构体类型，并完成修改。");
                }
            }

            if (!patched)
            {
                // 最后的兜底方案：如果模板里根本没有现成结构体可改，
                // 那就克隆一个已有的 Parameter 节点作为模板，再补出 3 个简单标量参数
                XElement templateParameterElement = parameterElements.FirstOrDefault();
                if (templateParameterElement == null)
                {
                    throw new Exception("工程TMC中没有任何现成的Parameter节点，无法克隆模板参数。");
                }

                Console.WriteLine("→ 当前模板里没有默认结构体参数，改为克隆一个现有有效参数模板来新增三个标量参数...");

                // 依次补出三个常用标量参数
                UpsertScalarParameterByCloningTemplate(parametersElement, templateParameterElement, "Gain", "LREAL", "自动添加参数：Gain");
                UpsertScalarParameterByCloningTemplate(parametersElement, templateParameterElement, "Enable", "BOOL", "自动添加参数：Enable");
                UpsertScalarParameterByCloningTemplate(parametersElement, templateParameterElement, "VelocityLimit", "LREAL", "自动添加参数：VelocityLimit");
                patched = true;
            }

            // 理论上 patched 到这里应该已经为 true；否则说明前面的所有策略都没成功
            if (!patched)
            {
                throw new Exception("工程TMC修改失败。");
            }

            // 把修改后的 XML 写回原文件；这里关闭自动格式化，尽量减少无关 diff
            tmcDoc.Save(tmcFilePath, SaveOptions.DisableFormatting);

            // 通知 VS 保存一下，确保工程视图里看到的是最新状态
            SaveAll(dte);

            Console.WriteLine($"→ 已备份原始工程TMC：{backupPath}");
            return tmcFilePath;
        }

        // 修改结构体参数/结构体类型的前三个字段
        private static void PatchStructuredTypeElement(XElement structuredElement)
        {
            // 把当前结构体节点下面的所有 SubItem 取出来。
            // 这里假设一个结构体参数会由多个 SubItem 组成，每个 SubItem 对应一个字段。
            List<XElement> subItems = structuredElement.Elements()
                .Where(x => x.Name.LocalName == "SubItem")
                .ToList();

            // 我们要把这个结构体改造成 3 个字段：
            // Gain / Enable / VelocityLimit
            // 因此至少需要 3 个现成的 SubItem 可供复用；否则就不安全，直接报错。
            if (subItems.Count < 3)
            {
                throw new Exception("默认结构体参数的SubItem数量不足3个，无法安全改成 Gain / Enable / VelocityLimit。");
            }

            // 如果当前传进来的就是一个 <Parameter> 节点本身，
            // 那顺手把它的名字统一改成 Parameter，避免沿用模板里原来的旧名字。
            if (structuredElement.Name.LocalName == "Parameter")
            {
                SetOrCreateChildElementValue(structuredElement, "Name", "Parameter");
            }

            // 把前 3 个字段依次改成目标结构：
            // 1) Gain : LREAL，占 64 bit，从 bit offset 0 开始
            // 2) Enable : BOOL，占 8 bit，从 bit offset 64 开始
            // 3) VelocityLimit : LREAL，占 64 bit，从 bit offset 128 开始
            //
            // 这里本质上是在“复用原来的 3 个字段壳子”，
            // 把字段名、类型、位宽和位偏移改成我们需要的定义。
            PatchStructuredSubItem(subItems[0], "Gain", "LREAL", 64, 0);
            PatchStructuredSubItem(subItems[1], "Enable", "BOOL", 8, 64);
            PatchStructuredSubItem(subItems[2], "VelocityLimit", "LREAL", 64, 128);

            // 如果这个结构体刚好只有 3 个字段，那总位宽就可以明确写成 192 bit。
            // 64 + 8 + 64 = 136，但这里用 192 的意图更接近“按槽位/对齐后的整体结构大小”。
            // 这部分其实带一点模板假设：认为保守写成 192 更符合原始结构体布局预期。
            if (subItems.Count == 3)
            {
                SetOrCreateChildElementValue(structuredElement, "BitSize", "192");
            }
        }

        /// <summary>
        /// 修改单个 SubItem：把它重写成一个目标字段（字段名 / 类型 / 位宽 / 位偏移）
        /// </summary>
        private static void PatchStructuredSubItem(XElement subItem, string fieldName, string typeName, int bitSize, int bitOffs)
        {
            // 先把这个字段的名字改掉，例如 Gain / Enable / VelocityLimit
            SetOrCreateChildElementValue(subItem, "Name", fieldName);

            // 读取或新建 <Type> 节点，并把类型改成目标类型。
            // 这里会清掉旧属性，避免把模板里原来类型相关的属性残留下来。
            XElement typeElement = GetOrCreateChildElement(subItem, "Type");
            typeElement.RemoveAttributes();
            typeElement.Value = typeName;

            // 写入位宽和位偏移。
            // BitSize 表示这个字段占多少 bit；
            // BitOffs 表示这个字段在整个结构体里的起始 bit 位置。
            SetOrCreateChildElementValue(subItem, "BitSize", bitSize.ToString());
            SetOrCreateChildElementValue(subItem, "BitOffs", bitOffs.ToString());

            // 如果原字段里有默认值定义，就先删掉。
            // 原因是我们现在是在“借壳重写”字段，
            // 继续保留模板里的旧默认值可能会和新类型/新语义不一致。
            XElement defaultElement = subItem.Elements().FirstOrDefault(x => x.Name.LocalName == "Default");
            if (defaultElement != null)
            {
                defaultElement.Remove();
            }
        }

        /// <summary>
        /// 往生成的 C++ 模块源码里插入一小段非常简单的示例代码，
        /// </summary>
        private static void WriteSimpleCodeToGeneratedModule(DTE dte)
        {
            // 先定位生成出来的模块 cpp 文件路径
            string cppFilePath = GetGeneratedModuleCppFilePath(dte);

            // 读出整个源文件内容，后面在内存里做字符串插入，再写回磁盘
            string sourceCode = File.ReadAllText(cppFilePath);

            // 这是我们自己插入代码时打的标记；
            // 用它来避免重复执行时把同一段代码插进去很多次
            const string marker = "// AUTO_WRITTEN_BY_TOOL";

            // 如果已经存在这个标记，说明之前已经自动插入过一次，直接跳过
            if (sourceCode.Contains(marker))
            {
                Console.WriteLine("→ 已检测到自动写入标记，跳过重复写入。");
                return;
            }

            // 这是准备插入到模块函数体里的简单示例代码：
            string injectCode =
                "    // AUTO_WRITTEN_BY_TOOL" + Environment.NewLine +
                "    static int s_autoCounter = 0;" + Environment.NewLine +
                "    ++s_autoCounter;" + Environment.NewLine + Environment.NewLine;

            // 优先寻找模板里常见的 TODO 注释位置；
            // 如果找到了，就把代码插在那个 TODO 前面，尽量贴近模板原本给用户留的示例区。
            int insertPos = sourceCode.IndexOf("// TODO: Replace the sample with your cyclic code", StringComparison.Ordinal);

            if (insertPos < 0)
            {
                // 如果没有找到 TODO 注释，就退一步：
                // 先找 CycleUpdate 函数，再把代码插到函数体开头的大括号后面。
                int cyclePos = sourceCode.IndexOf("CycleUpdate", StringComparison.Ordinal);
                if (cyclePos < 0)
                {
                    throw new Exception($"在 {cppFilePath} 中未找到 CycleUpdate，无法自动写入简单代码。");
                }

                // 找到 CycleUpdate 之后，再找它方法体起始的左大括号
                int bracePos = sourceCode.IndexOf('{', cyclePos);
                if (bracePos < 0)
                {
                    throw new Exception($"在 {cppFilePath} 中未找到 CycleUpdate 的方法体起始大括号。");
                }

                // 真正插入的位置是大括号后面一位，也就是函数体开头
                insertPos = bracePos + 1;
                injectCode = Environment.NewLine + injectCode;
            }

            // 在写回源码前先做一个 .bak 备份，避免自动修改后不方便恢复
            File.Copy(cppFilePath, cppFilePath + ".bak", true);

            // 把准备好的示例代码插入到目标位置
            sourceCode = sourceCode.Insert(insertPos, injectCode);

            // 把修改后的源码写回原 cpp 文件
            File.WriteAllText(cppFilePath, sourceCode);

            Console.WriteLine($"→ 已自动修改模块源码：{cppFilePath}");
            Console.WriteLine($"→ 已自动备份原始文件：{cppFilePath}.bak");
        }

        /// <summary>
        /// 通过克隆一个现有的有效参数模板来新增/更新标量参数
        /// </summary>
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

            // 关键修复 1：新参数不要继承模板里的隐藏属性
            XAttribute hideAttr = targetParameterElement.Attributes()
                .FirstOrDefault(a => a.Name.LocalName == "HideParameter");
            hideAttr?.Remove();

            // 关键修复 2：删除模板参数中与普通标量参数不相干的子项，避免 TMC 编辑器标红
            List<XElement> childElementsToRemove = targetParameterElement.Elements()
                .Where(e =>
                {
                    string localName = e.Name.LocalName;
                    return localName == "SubItem" ||
                           localName == "EnumInfo" ||
                           localName == "ArrayInfo" ||
                           localName == "Type" ||      // 后面统一由 SetParameterTypeInfo 重建/修正
                           localName == "BaseType";    // 后面统一由 SetParameterTypeInfo 重建/修正
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

        /// <summary>
        /// 设置参数类型信息
        /// </summary>
        private static void SetParameterTypeInfo(XElement parameterElement, string typeName)
        {
            // 优先使用 BaseType，并清掉克隆模板里遗留的 GUID/其它属性，避免
            // 出现 <BaseType GUID="原TcTraceLevel的GUID">LREAL</BaseType> 这种不一致情况
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

            // 如果同时存在 Type 和 BaseType，只保留一个，避免编辑器解析不一致
            if (baseTypeElement != null && typeElement != null)
            {
                typeElement.Remove();
            }
        }

        /// <summary>
        /// 设置参数大小信息
        /// </summary>
        private static void SetParameterSizeInfo(XElement parameterElement, string typeName)
        {
            int bitSize = GetBitSizeForSimpleType(typeName);
            int byteSize = Math.Max(1, bitSize / 8);

            SetOrCreateChildElementValue(parameterElement, "BitSize", bitSize.ToString());
            SetOrCreateChildElementValue(parameterElement, "Size", byteSize.ToString());
            SetOrCreateChildElementValue(parameterElement, "SizeX64", byteSize.ToString());
        }

        /// <summary>
        /// 设置参数默认值
        /// </summary>
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

        /// <summary>
        /// 设置参数常量名，避免克隆TraceLevelMax后残留同名常量
        /// </summary>
        private static void SetParameterConstantName(XElement parameterElement, string parameterName)
        {
            XElement constantNameElement = parameterElement.Elements().FirstOrDefault(x => x.Name.LocalName == "ConstantName");
            if (constantNameElement != null)
            {
                constantNameElement.Value = "PID_" + parameterName;
            }
        }

        /// <summary>
        /// 去掉TraceLevelMax模板里可能遗留的特定元数据
        /// </summary>
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

        /// <summary>
        /// 获取简单类型对应的位宽
        /// </summary>
        private static int GetBitSizeForSimpleType(string typeName)
        {
            switch ((typeName ?? string.Empty).ToUpperInvariant())
            {
                case "BOOL":
                    return 8;
                case "LREAL":
                    return 64;
                case "REAL":
                    return 32;
                case "ULINT":
                case "LINT":
                    return 64;
                case "UDINT":
                case "DINT":
                    return 32;
                case "UINT":
                case "INT":
                    return 16;
                case "USINT":
                case "SINT":
                    return 8;
                default:
                    return 32;
            }
        }

        /// <summary>
        /// 生成下一个可用的用户自定义PTCID
        /// </summary>
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

        /// <summary>
        /// 解析#x00000001风格的十六进制ID
        /// </summary>
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

        /// <summary>
        /// 通过参数类型名定位DataType节点
        /// </summary>
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

        /// <summary>
        /// 获取或创建子元素
        /// </summary>
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

        /// <summary>
        /// 设置或创建子元素的值
        /// </summary>
        private static void SetOrCreateChildElementValue(XElement parent, string childLocalName, string value)
        {
            XElement child = GetOrCreateChildElement(parent, childLocalName);
            child.Value = value;
        }

        /// <summary>
        /// 读取子元素的值
        /// </summary>
        private static string GetChildElementValue(XElement parent, string childLocalName)
        {
            return parent.Elements().FirstOrDefault(x => x.Name.LocalName == childLocalName)?.Value;
        }

        /// <summary>
        /// 查找工程里的TMC文件路径
        /// </summary>
        private static string GetProjectTmcFilePath(DTE dte)
        {
            string tmcFilePath = TryGetProjectTmcFilePath(dte);
            if (string.IsNullOrEmpty(tmcFilePath))
            {
                throw new FileNotFoundException("未找到工程TMC文件", DEFAULT_CPP_PROJECT_NAME + ".tmc");
            }

            return tmcFilePath;
        }

        /// <summary>
        /// 尝试查找工程里的TMC文件路径
        /// </summary>
        private static string TryGetProjectTmcFilePath(DTE dte)
        {
            return TryFindProjectOwnedFile(dte, DEFAULT_CPP_PROJECT_NAME + ".tmc");
        }

        /// <summary>
        /// 查找生成的模块cpp路径
        /// </summary>
        private static string GetGeneratedModuleCppFilePath(DTE dte)
        {
            string cppFilePath = TryGetGeneratedModuleCppFilePath(dte);
            if (string.IsNullOrEmpty(cppFilePath))
            {
                throw new FileNotFoundException("未找到生成的模块cpp文件", DEFAULT_MODULE_NAME + ".cpp");
            }

            return cppFilePath;
        }

        /// <summary>
        /// 尝试查找生成的模块cpp路径
        /// </summary>
        private static string TryGetGeneratedModuleCppFilePath(DTE dte)
        {
            return TryFindProjectOwnedFile(dte, DEFAULT_MODULE_NAME + ".cpp");
        }

        /// <summary>
        /// 在当前Solution目录下查找指定文件（排除发布目录/中间目录）
        /// </summary>
        private static string TryFindProjectOwnedFile(DTE dte, string fileName)
        {
            string solutionDir = GetSolutionDirectory(dte);

            return SafeEnumerateFiles(solutionDir, fileName)
                .Where(IsProjectSourcePath)
                .OrderBy(path => path.IndexOf(Path.DirectorySeparatorChar + DEFAULT_CPP_PROJECT_NAME + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) >= 0 ? 0 : 1)
                .ThenBy(path => path.Length)
                .FirstOrDefault();
        }

        /// <summary>
        /// 获取当前Solution目录
        /// </summary>
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

        /// <summary>
        /// 安全枚举文件
        /// </summary>
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

        /// <summary>
        /// 判断文件是否属于工程源文件，而不是发布目录/中间目录
        /// </summary>
        private static bool IsProjectSourcePath(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return false;
            }

            string normalizedPath = filePath.Replace('/', '\\');

            if (normalizedPath.IndexOf(@"\CustomConfig\Modules\", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            if (normalizedPath.IndexOf(@"\Repository\", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            if (normalizedPath.IndexOf(@"\bin\", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            if (normalizedPath.IndexOf(@"\obj\", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            if (normalizedPath.EndsWith(".bak", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 从指定TMC文件中读取模块GUID,因为TcCOM模块的 GUID 可以从 TMC 文件中获取
        /// </summary>
        private static string GetModuleGuidFromTmcFile(string tmcFilePath)
        {
            // 打印路径用于调试
            Console.WriteLine($"→ 尝试读取TMC文件：{Path.GetFullPath(tmcFilePath)}");

            if (!File.Exists(tmcFilePath))
            {
                throw new FileNotFoundException("TMC文件不存在", tmcFilePath);
            }

            // 加载XML文档
            XDocument tmcDoc = XDocument.Load(tmcFilePath);

            // 查找第一个<Module>节点并提取GUID属性
            XElement moduleElement = tmcDoc.Descendants().FirstOrDefault(x => x.Name.LocalName == "Module");
            if (moduleElement == null)
            {
                throw new Exception("TMC文件中未找到<Module>节点");
            }

            string guidValue = moduleElement.Attributes().FirstOrDefault(x => x.Name.LocalName == "GUID")?.Value;
            if (string.IsNullOrEmpty(guidValue))
            {
                throw new Exception("<Module>节点未包含GUID属性");
            }

            return guidValue;
        }

        /// <summary>
        /// 5秒内可自定义TMC基础路径，返回最终完整TMC路径
        /// </summary>
        private static string GetFinalTmcFilePath()
        {
            string tmcBasePath = DEFAULT_TMC_BASE_PATH;

            // 提示用户5秒内可修改路径
            Console.WriteLine($"\n→ TMC文件默认基础路径：{DEFAULT_TMC_BASE_PATH}");
            Console.WriteLine("→ 如需修改，请在5秒内按下任意键；否则将使用默认路径...");

            int timeout = DEFAULT_TIMEOUT; // 5秒超时
            bool isKeyPressed = false;

            while (timeout > 0 && !isKeyPressed)
            {
                if (Console.KeyAvailable)
                {
                    Console.ReadKey(true); // 读取按键但不显示
                    isKeyPressed = true;
                    break;
                }
                System.Threading.Thread.Sleep(CHECK_INTERVAL);
                timeout -= CHECK_INTERVAL;
            }

            // 如果用户按下按键，自定义基础路径
            if (isKeyPressed)
            {
                Console.Write("\n请输入新的TMC基础路径（如C:\\TwinCAT\\3.2\\CustomConfig\\Modules）：");
                string newBasePath = Console.ReadLine()?.Trim();

                // 校验输入的路径是否合法
                if (!string.IsNullOrEmpty(newBasePath))
                {
                    // 处理路径末尾的反斜杠，避免重复拼接
                    tmcBasePath = newBasePath.TrimEnd('\\');
                    Console.WriteLine($"→ 已使用自定义基础路径：{tmcBasePath}");
                }
                else
                {
                    Console.WriteLine("→ 输入为空，使用默认基础路径...");
                }
            }
            else
            {
                Console.WriteLine("→ 5秒超时，使用默认基础路径...");
            }

            // 拼接完整TMC路径
            string finalTmcPath = Path.Combine(tmcBasePath, TMC_RELATIVE_PATH.TrimStart('\\'));
            Console.WriteLine($"→ 最终TMC文件路径：{finalTmcPath}");

            return finalTmcPath;
        }

        /// <summary>
        /// 添加TcCOM Object（支持自定义TMC路径+读取真实GUID）
        /// </summary>
        private static void AddTcComObject(ITcSysManager sysManager, ITcSmTreeItem cppProject)
        {
            // 从tree里定位TcCOM Objects节点
            ITcSmTreeItem tcComObjects = sysManager.LookupTreeItem("TIRC^TcCOM Objects");
            if (tcComObjects == null)
            {
                throw new Exception("未找到TcCOM Objects节点（TIRC^TcCOM Objects）！");
            }

            try
            {
                // 因为我们前一步发布了TcCOM Modules，所以现在可以直接读取TMC文件中的TcCOM 模块的 GUID，
                //  TwinCAT 创建实例必须先知道“类型”， 而TwinCAT 是根据 GUID 找模块类型 的。
                string finalTmcPath = GetFinalTmcFilePath();
                string realModuleGuid = GetModuleGuidFromTmcFile(finalTmcPath);
                Console.WriteLine($"√  从TMC文件读取到模块GUID：{realModuleGuid}");

                // 使用真实GUID创建TcCOM对象，也就是创建实例. 即TwinCAT 会在 TcCOM Objects 节点下，新增一个该模块类型的实例。
                ITcSmTreeItem newTcComObject = tcComObjects.CreateChild("CustomTcComModule", 0, "", realModuleGuid);
                //父节点：tcComObject
                //新实例名称："CustomTcComModule"

                if (newTcComObject == null)
                {
                    throw new Exception("添加TcCOM Object失败！请确认GUID对应的模块已被TwinCAT检测到。");
                }

                Console.WriteLine($"√  TcCOM Object添加成功！");
            }
            catch (Exception ex)
            {
                throw new Exception($"添加TcCOM Object失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 对Visual Studio / TwinCAT的COM调用做简单重试，避免“被呼叫方拒绝接收呼叫”
        /// </summary>
        private static void RetryComCall(Action action, int maxRetryCount = 60, int retryDelayMs = 500)
        {
            RetryComCall<object>(() =>
            {
                action();
                return null;
            }, maxRetryCount, retryDelayMs);
        }

        /// <summary>
        /// 对Visual Studio / TwinCAT的COM调用做简单重试，避免“被呼叫方拒绝接收呼叫”
        /// </summary>
        private static T RetryComCall<T>(Func<T> func, int maxRetryCount = 60, int retryDelayMs = 500)
        {
            Exception lastException = null;

            for (int i = 0; i < maxRetryCount; i++)
            {
                try
                {
                    return func();
                }
                catch (COMException ex) when (ex.ErrorCode == RPC_E_CALL_REJECTED || ex.ErrorCode == RPC_E_SERVERCALL_RETRYLATER)
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

        /// <summary>
        /// 让当前控制台进程在VS繁忙时自动重试COM调用
        /// </summary>
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

        // P/Invoke声明
        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);
        #endregion
    }
}
