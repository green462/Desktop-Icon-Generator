// =====================================================================
//  桌面快捷方式生成器(AppShortcutCreator)
//  -----------------------------------------------------------------
//  为「桌面上没有图标的应用」一键创建桌面快捷方式。
//  原理等价于手动流程:Win+R -> shell:AppsFolder -> 右键 -> 创建快捷方式。
//
//  - UWP/商店应用:通过 SHParseDisplayName 把 AUMID 解析为 PIDL,
//    再用 IShellLinkW.SetIDList 生成快捷方式(与系统手动创建效果一致,
//    双击可启动、图标自动继承)。
//  - Win32 桌面应用:IShellLinkW.SetPath/SetWorkingDirectory/SetIconLocation。
//
//  编译:运行 build.bat(使用系统自带 csc.exe,.NET Framework 4.x,零第三方依赖)。
// =====================================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32.SafeHandles;

namespace DesktopShortcutTool
{
    // =================================================================
    // 一、数据模型
    // =================================================================

    /// <summary>应用类型。</summary>
    internal enum AppKind
    {
        /// <summary>传统 Win32 桌面应用(Path 为 exe 全路径)。</summary>
        Desktop,

        /// <summary>UWP/商店应用(Path 为 AUMID,如 Microsoft.WindowsCalculator_8wekyb3d8bbwe!App)。</summary>
        Uwp,

        /// <summary>系统未提供路径信息的特殊项,暂不支持创建。</summary>
        Unsupported
    }

    /// <summary>应用列表中的一项。</summary>
    internal sealed class AppInfo
    {
        /// <summary>应用显示名称。</summary>
        public readonly string Name;

        /// <summary>Win32 应用为 exe 全路径;UWP 应用为 AUMID;无路径的项为 null。</summary>
        public readonly string Path;

        /// <summary>应用类型。</summary>
        public readonly AppKind Kind;

        public AppInfo(string name, string path, AppKind kind)
        {
            Name = name;
            Path = path;
            Kind = kind;
        }

        /// <summary>类型的中文显示文本。</summary>
        public string KindText
        {
            get
            {
                switch (Kind)
                {
                    case AppKind.Desktop: return "桌面应用";
                    case AppKind.Uwp: return "UWP 应用";
                    default: return "暂不支持";
                }
            }
        }
    }

    /// <summary>单个快捷方式的创建结果。</summary>
    internal sealed class CreateResult
    {
        public readonly bool Success;
        public readonly string Message;
        public readonly string LnkPath;

        private CreateResult(bool success, string message, string lnkPath)
        {
            Success = success;
            Message = message;
            LnkPath = lnkPath;
        }

        public static CreateResult Ok(string lnkPath, string message)
        {
            return new CreateResult(true, message, lnkPath);
        }

        public static CreateResult Fail(string message)
        {
            return new CreateResult(false, message, null);
        }
    }

    // =================================================================
    // 二、Win32 / COM 互操作定义
    // =================================================================

    /// <summary>ShellLink coclass(IShellLinkW 的系统实现,即 .lnk 的 COM 对象)。</summary>
    [ComImport]
    [Guid("00021401-0000-0000-C000-000000000046")]
    [ComDefaultInterface(typeof(IShellLinkW))]
    internal class ShellLinkClass
    {
    }

    /// <summary>IShellLinkW(Unicode 版),用于设置快捷方式的各种属性。</summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("000214F9-0000-0000-C000-000000000046")]
    internal interface IShellLinkW
    {
        void GetPath([MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, IntPtr pfd, uint fFlags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string lpszName);
        void GetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string lpszDir);
        void GetArguments([MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxArgs);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string lpszArgs);
        void GetHotkey(out short pwHotkey);
        void SetHotkey(short wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation([MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string lpszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string lpszPathRel, uint dwReserved);
        void Resolve(IntPtr hWnd, uint fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string lpszFile);
    }

    /// <summary>IPersistFile,用于把内存中的快捷方式保存为 .lnk 文件。</summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("0000010B-0000-0000-C000-000000000046")]
    internal interface IPersistFile
    {
        // ---- IPersist ----
        void GetClassID(out Guid pClassID);

        // ---- IPersistFile ----
        void IsDirty();
        void Load([MarshalAs(UnmanagedType.LPWStr)] string lpszFileName, uint dwMode);
        void Save([MarshalAs(UnmanagedType.LPWStr)] string lpszFileName, bool fRemember);
        void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string lpszFileName);
        void GetCurFile([MarshalAs(UnmanagedType.LPWStr)] out string lplpszFileName);
    }

    /// <summary>shell32 原生函数(快捷方式相关)。</summary>
    internal static class Native
    {
        /// <summary>把 "shell:AppsFolder\{AUMID}" 之类的路径解析为 PIDL。失败时抛 COMException。</summary>
        [DllImport("shell32.dll", PreserveSig = false)]
        public static extern void SHParseDisplayName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszName,
            IntPtr pbc,
            out IntPtr ppidl,
            uint sfgaoIn,
            out uint psfgaoOut);

        /// <summary>释放 SHParseDisplayName 返回的 PIDL。</summary>
        [DllImport("shell32.dll")]
        public static extern void ILFree(IntPtr pidl);
    }

    /// <summary>
    /// 控制台互操作。本程序编译为 winexe,默认没有控制台;
    /// 从命令行窗口直接运行时,通过附着父进程控制台让 --list / --create 可以输出文本。
    /// </summary>
    internal static class NativeConsole
    {
        private const uint ATTACH_PARENT_PROCESS = 0xFFFFFFFFu;
        private const int STD_OUTPUT_HANDLE = -11;
        private const int STD_ERROR_HANDLE = -12;
        private const uint GENERIC_WRITE = 0x40000000u;
        private const uint FILE_SHARE_READ = 0x00000001u;
        private const uint FILE_SHARE_WRITE = 0x00000002u;
        private const uint OPEN_EXISTING = 3u;

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AttachConsole(uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetStdHandle(int nStdHandle);

        [DllImport("kernel32.dll")]
        private static extern uint GetFileType(IntPtr hFile);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateFileW(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        /// <summary>
        /// 若标准输出已被重定向(文件/管道),则句柄有效,直接输出即可,不作任何处理;
        /// 否则附着父进程控制台,并把 Console.Out / Console.Error 重新绑定到 CONOUT$。
        /// </summary>
        public static void TryAttach()
        {
            if (HasValidStdHandle(STD_OUTPUT_HANDLE)) return;
            try
            {
                if (!AttachConsole(ATTACH_PARENT_PROCESS)) return;
                RebindToConsoleOutput(STD_ERROR_HANDLE);
                RebindToConsoleOutput(STD_OUTPUT_HANDLE);
            }
            catch
            {
                // 附着失败(例如双击启动、无父控制台)时静默忽略,不影响 GUI 模式。
            }
        }

        /// <summary>刷新输出缓冲并脱离控制台,必须在 CLI 进程退出前调用。</summary>
        public static void FlushAndFree()
        {
            try { Console.Out.Flush(); } catch { }
            try { Console.Error.Flush(); } catch { }
            try { FreeConsole(); } catch { }
        }

        private static bool HasValidStdHandle(int stdHandleId)
        {
            IntPtr handle = GetStdHandle(stdHandleId);
            if (handle == IntPtr.Zero || handle == (IntPtr)(-1)) return false;
            return GetFileType(handle) != 0; // 返回 0 表示句柄无效
        }

        private static void RebindToConsoleOutput(int stdHandleId)
        {
            IntPtr handle = CreateFileW(
                "CONOUT$", GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE,
                IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
            if (handle == IntPtr.Zero || handle == (IntPtr)(-1)) return;
            var writer = new StreamWriter(
                new FileStream(new SafeFileHandle(handle, false), FileAccess.Write),
                GetOutputEncoding()) { AutoFlush = true };
            if (stdHandleId == STD_ERROR_HANDLE) Console.SetError(writer);
            else Console.SetOut(writer);
        }

        private static Encoding GetOutputEncoding()
        {
            try { return Console.OutputEncoding; }
            catch { return Encoding.Default; }
        }
    }

    // =================================================================
    // 三、核心逻辑:枚举应用 + 创建快捷方式
    // =================================================================

    internal static class ShortcutTool
    {
        /// <summary>
        /// 枚举系统里所有可识别的应用(来源:shell:AppsFolder,即"开始菜单里的所有应用")。
        /// 分隔符等无效项直接跳过;Path 为空的项标记为「暂不支持」并保留展示。
        /// </summary>
        public static List<AppInfo> EnumerateApps()
        {
            var apps = new List<AppInfo>();

            // 通过 Shell.Application COM 对象访问 AppsFolder,用 dynamic 简化后期绑定。
            dynamic shell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
            if (shell == null) throw new InvalidOperationException("无法创建 Shell.Application COM 对象。");
            dynamic folder = shell.Namespace("shell:AppsFolder");
            if (folder == null) throw new InvalidOperationException("无法打开 shell:AppsFolder。");

            foreach (dynamic item in folder.Items())
            {
                string name;
                string path;
                try
                {
                    name = (string)item.Name;
                    path = (string)item.Path;
                }
                catch
                {
                    continue; // 个别项属性读取失败,按无效项跳过
                }
                if (string.IsNullOrWhiteSpace(name)) continue; // 分隔符等无名称项

                AppKind kind;
                if (string.IsNullOrEmpty(path))
                {
                    kind = AppKind.Unsupported; // 系统未提供路径,保留展示但不可创建
                }
                else if (path.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) && path.IndexOf('\\') >= 0)
                {
                    kind = AppKind.Desktop; // 真实的 exe 全路径 -> Win32 桌面应用
                }
                else
                {
                    kind = AppKind.Uwp; // 形如 PackageFamily!AppId 的 AUMID -> UWP 应用
                }
                apps.Add(new AppInfo(name, path, kind));
            }
            return apps;
        }

        /// <summary>
        /// 为指定应用创建桌面快捷方式。同名 .lnk 已存在时先删除旧文件再覆盖保存。
        /// </summary>
        public static CreateResult CreateShortcut(AppInfo app)
        {
            if (app == null) return CreateResult.Fail("应用信息为空。");
            if (app.Kind == AppKind.Unsupported)
                return CreateResult.Fail("暂不支持(系统未提供该应用的路径信息)。");

            // 用 DesktopDirectory 而非逻辑桌面,确保拿到真实文件系统路径(兼容 OneDrive 重定向)。
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            if (string.IsNullOrEmpty(desktop)) return CreateResult.Fail("无法定位桌面目录。");

            string lnkPath = System.IO.Path.Combine(desktop, SanitizeFileName(app.Name) + ".lnk");
            bool existed = File.Exists(lnkPath);
            try
            {
                if (existed) File.Delete(lnkPath); // Save 前先删旧文件,确保覆盖成功

                IShellLinkW link = (IShellLinkW)new ShellLinkClass();
                if (app.Kind == AppKind.Desktop)
                {
                    // Win32 应用:直接指向 exe 本体
                    link.SetPath(app.Path);
                    string dir = System.IO.Path.GetDirectoryName(app.Path);
                    link.SetWorkingDirectory(string.IsNullOrEmpty(dir) ? desktop : dir);
                    link.SetIconLocation(app.Path, 0);
                }
                else
                {
                    // UWP 应用:把 "shell:AppsFolder\{AUMID}" 解析成 PIDL,再挂到快捷方式上。
                    // 这与系统手动"创建快捷方式"的产物完全一致。
                    uint attributes;
                    IntPtr pidl;
                    Native.SHParseDisplayName("shell:AppsFolder\\" + app.Path, IntPtr.Zero, out pidl, 0, out attributes);
                    try { link.SetIDList(pidl); }
                    finally { if (pidl != IntPtr.Zero) Native.ILFree(pidl); }
                }
                link.SetDescription(app.Name);
                ((IPersistFile)link).Save(lnkPath, true);
                return CreateResult.Ok(lnkPath, existed ? "已存在,已覆盖" : "创建成功");
            }
            catch (Exception ex)
            {
                return CreateResult.Fail(ex.Message);
            }
        }

        /// <summary>去掉应用名中不能用于文件名的字符,得到 .lnk 文件名主体。</summary>
        public static string SanitizeFileName(string name)
        {
            char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (char c in name)
            {
                if (Array.IndexOf(invalidChars, c) < 0) sb.Append(c);
            }
            string clean = sb.ToString().Trim();
            return clean.Length > 0 ? clean : "快捷方式";
        }
    }

    // =================================================================
    // 四、CLI 模式(--list / --create)
    // =================================================================

    internal static class CliRunner
    {
        /// <summary>执行 CLI 命令,返回进程退出码(0 成功 / 1 失败)。</summary>
        public static int Run(string[] args)
        {
            string command = args[0].ToLowerInvariant();
            if (command == "--list") return RunList();
            if (command == "--create") return RunCreate(args.Length > 1 ? args[1] : null);
            PrintUsage();
            return 1;
        }

        /// <summary>--list:按「序号|名称|类型|路径或AUMID」打印所有应用。</summary>
        private static int RunList()
        {
            List<AppInfo> apps;
            try
            {
                apps = ShortcutTool.EnumerateApps();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("枚举应用失败: " + ex.Message);
                return 1;
            }

            int index = 1;
            foreach (AppInfo app in apps)
            {
                Console.WriteLine(string.Format("{0}|{1}|{2}|{3}", index++, app.Name, app.KindText, app.Path ?? ""));
            }
            int unsupported = 0;
            foreach (AppInfo a in apps)
            {
                if (a.Kind == AppKind.Unsupported) unsupported++;
            }
            Console.WriteLine(string.Format("# 共 {0} 个应用(其中 {1} 个因缺少路径信息暂不支持创建)", apps.Count, unsupported));
            return 0;
        }

        /// <summary>--create:按名称模糊匹配第一个命中项并创建桌面快捷方式。</summary>
        private static int RunCreate(string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                Console.Error.WriteLine("用法: AppShortcutCreator.exe --create <应用名称>");
                return 1;
            }

            List<AppInfo> apps;
            try
            {
                apps = ShortcutTool.EnumerateApps();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("枚举应用失败: " + ex.Message);
                return 1;
            }

            AppInfo hit = null;
            foreach (AppInfo app in apps)
            {
                if (app.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    hit = app;
                    break;
                }
            }
            if (hit == null)
            {
                Console.Error.WriteLine(string.Format("未找到名称包含 \"{0}\" 的应用。可先运行 --list 查看全部应用。", query));
                return 1;
            }

            CreateResult result = ShortcutTool.CreateShortcut(hit);
            if (result.Success)
            {
                Console.WriteLine(string.Format("[成功] {0} -> {1}({2})", hit.Name, result.LnkPath, result.Message));
                return 0;
            }
            Console.Error.WriteLine(string.Format("[失败] {0}: {1}", hit.Name, result.Message));
            return 1;
        }

        private static void PrintUsage()
        {
            Console.WriteLine("桌面快捷方式生成器");
            Console.WriteLine("用法:");
            Console.WriteLine("  AppShortcutCreator.exe                 打开图形界面");
            Console.WriteLine("  AppShortcutCreator.exe --list          列出所有可识别的应用");
            Console.WriteLine("  AppShortcutCreator.exe --create <名称>  按名称模糊匹配并创建桌面快捷方式");
        }
    }

    // =================================================================
    // 五、GUI 模式(WinForms,全中文界面)
    // =================================================================

    internal sealed class MainForm : Form
    {
        private readonly TextBox _searchBox = new TextBox();
        private readonly Label _searchPlaceholder = new Label();
        private readonly ListView _listView = new ListView();
        private readonly Button _btnSelectAll = new Button();
        private readonly Button _btnCreate = new Button();
        private readonly Label _statusLabel = new Label();
        private readonly List<AppInfo> _apps = new List<AppInfo>();

        public MainForm()
        {
            SetupWindow();
            SetupListView();     // Dock.Fill 必须最先加入
            SetupBottomPanel();  // Dock.Bottom
            SetupSearchPanel();  // Dock.Top
            LoadApps();
        }

        // ---------- 界面搭建 ----------

        private void SetupWindow()
        {
            Text = "桌面快捷方式生成器";
            ClientSize = new Size(640, 480);
            MinimumSize = new Size(560, 420);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9F);
        }

        /// <summary>顶部搜索框 + 占位文本。</summary>
        private void SetupSearchPanel()
        {
            var panel = new Panel { Dock = DockStyle.Top, Height = 36, Padding = new Padding(8, 8, 8, 0) };

            _searchBox.Dock = DockStyle.Fill;
            _searchBox.TextChanged += delegate { OnSearchTextChanged(); };

            _searchPlaceholder.Text = "搜索应用…";
            _searchPlaceholder.ForeColor = Color.Gray;
            _searchPlaceholder.BackColor = Color.Transparent;
            _searchPlaceholder.AutoSize = true;
            _searchPlaceholder.Cursor = Cursors.IBeam;
            _searchPlaceholder.Location = new Point(12, 10);
            _searchPlaceholder.Click += delegate { _searchBox.Focus(); };

            panel.Controls.Add(_searchBox);
            panel.Controls.Add(_searchPlaceholder);
            Controls.Add(panel);
        }

        /// <summary>中部应用列表(可勾选,双击直接创建)。</summary>
        private void SetupListView()
        {
            _listView.View = View.Details;
            _listView.CheckBoxes = true;
            _listView.FullRowSelect = true;
            _listView.MultiSelect = true;
            _listView.HideSelection = false;
            _listView.Dock = DockStyle.Fill;
            _listView.Columns.Add("应用名称", 170);
            _listView.Columns.Add("类型", 90);
            _listView.Columns.Add("路径或 AUMID", 340);
            _listView.MouseDoubleClick += OnListDoubleClick;
            _listView.Resize += delegate { AdjustLastColumn(); };
            Controls.Add(_listView);
        }

        /// <summary>底部操作按钮 + 多行状态栏。</summary>
        private void SetupBottomPanel()
        {
            var panel = new Panel { Dock = DockStyle.Bottom, Height = 132, Padding = new Padding(8) };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 34,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            _btnCreate.Text = "创建到桌面";
            _btnCreate.Size = new Size(120, 28);
            _btnCreate.Margin = new Padding(0, 3, 8, 3);
            _btnCreate.Click += OnCreateClicked;
            _btnSelectAll.Text = "全选";
            _btnSelectAll.Size = new Size(80, 28);
            _btnSelectAll.Margin = new Padding(0, 3, 8, 3);
            _btnSelectAll.Click += OnSelectAllClicked;
            buttons.Controls.Add(_btnCreate);
            buttons.Controls.Add(_btnSelectAll);

            _statusLabel.Dock = DockStyle.Fill;
            _statusLabel.ForeColor = SystemColors.ControlText;

            panel.Controls.Add(_statusLabel); // Dock.Fill 必须先于按钮面板加入
            panel.Controls.Add(buttons);
            Controls.Add(panel);
        }

        // ---------- 数据加载与列表填充 ----------

        private void LoadApps()
        {
            List<AppInfo> apps;
            Cursor = Cursors.WaitCursor;
            try
            {
                apps = ShortcutTool.EnumerateApps();
            }
            catch (Exception ex)
            {
                _statusLabel.Text = "加载应用列表失败: " + ex.Message;
                return;
            }
            finally
            {
                Cursor = Cursors.Default;
            }

            _apps.Clear();
            _apps.AddRange(apps);
            RepopulateList(_searchBox.Text);

            int unsupported = 0;
            foreach (AppInfo a in _apps)
            {
                if (a.Kind == AppKind.Unsupported) unsupported++;
            }
            string tip = string.Format("共加载 {0} 个应用。勾选后点“创建到桌面”,或双击某一项直接创建。", _apps.Count);
            if (unsupported > 0) tip += string.Format("(另有 {0} 项系统未提供路径,暂不支持。)", unsupported);
            _statusLabel.Text = tip;
        }

        /// <summary>按过滤条件重新填充列表(即时过滤)。</summary>
        private void RepopulateList(string filter)
        {
            _listView.BeginUpdate();
            _listView.Items.Clear();
            foreach (AppInfo app in _apps)
            {
                if (!MatchesFilter(app, filter)) continue;
                var item = new ListViewItem(app.Name);
                item.SubItems.Add(app.KindText);
                item.SubItems.Add(app.Path ?? "");
                item.Tag = app;
                _listView.Items.Add(item);
            }
            _listView.EndUpdate();
        }

        private static bool MatchesFilter(AppInfo app, string filter)
        {
            if (string.IsNullOrWhiteSpace(filter)) return true;
            return (app.Name != null && app.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                || (app.Path != null && app.Path.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        /// <summary>窗口缩放时让第三列(路径)占满剩余宽度。</summary>
        private void AdjustLastColumn()
        {
            if (_listView.Columns.Count < 3) return;
            int fixedWidth = _listView.Columns[0].Width + _listView.Columns[1].Width;
            int width = _listView.ClientSize.Width - fixedWidth - 4;
            _listView.Columns[2].Width = width > 220 ? width : 220;
        }

        // ---------- 事件处理 ----------

        private void OnSearchTextChanged()
        {
            _searchPlaceholder.Visible = _searchBox.TextLength == 0;
            RepopulateList(_searchBox.Text);
        }

        /// <summary>双击列表某一项 = 直接为该项创建快捷方式。</summary>
        private void OnListDoubleClick(object sender, MouseEventArgs e)
        {
            if (_listView.SelectedItems.Count == 0) return;
            var app = _listView.SelectedItems[0].Tag as AppInfo;
            if (app == null) return;
            CreateForApps(new AppInfo[] { app });
        }

        /// <summary>全选 / 取消全选(自动切换)。</summary>
        private void OnSelectAllClicked(object sender, EventArgs e)
        {
            if (_listView.Items.Count == 0) return;
            bool anyUnchecked = false;
            foreach (ListViewItem item in _listView.Items)
            {
                if (!item.Checked) { anyUnchecked = true; break; }
            }
            foreach (ListViewItem item in _listView.Items)
            {
                item.Checked = anyUnchecked;
            }
            _btnSelectAll.Text = anyUnchecked ? "取消全选" : "全选";
        }

        /// <summary>把所有勾选项创建为桌面快捷方式。</summary>
        private void OnCreateClicked(object sender, EventArgs e)
        {
            var selected = new List<AppInfo>();
            foreach (ListViewItem item in _listView.CheckedItems)
            {
                var app = item.Tag as AppInfo;
                if (app != null) selected.Add(app);
            }
            if (selected.Count == 0)
            {
                _statusLabel.Text = "请先勾选至少一个应用,或直接双击某一项。";
                return;
            }
            CreateForApps(selected);
        }

        /// <summary>为若干应用创建快捷方式,并把每个结果写入状态栏。数量不大,同步执行 + 等待光标即可。</summary>
        private void CreateForApps(IEnumerable<AppInfo> apps)
        {
            var list = new List<AppInfo>(apps);
            int success = 0;
            var lines = new List<string>();
            Cursor = Cursors.WaitCursor;
            try
            {
                foreach (AppInfo app in list)
                {
                    CreateResult result = ShortcutTool.CreateShortcut(app);
                    if (result.Success) success++;
                    lines.Add(string.Format("{0} {1}: {2}", result.Success ? "√" : "×", app.Name, result.Message));
                }
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            string summary = string.Format("本次操作:成功 {0} 个,失败 {1} 个。", success, list.Count - success);
            _statusLabel.Text = summary + Environment.NewLine + string.Join(Environment.NewLine, lines.ToArray());
        }
    }

    // =================================================================
    // 六、程序入口
    // =================================================================

    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (IsCliRequest(args))
            {
                // winexe 没有控制台:先附着父进程控制台,执行完刷新并脱离,最后返回退出码。
                NativeConsole.TryAttach();
                int exitCode = CliRunner.Run(args);
                NativeConsole.FlushAndFree();
                return exitCode;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            return 0;
        }

        private static bool IsCliRequest(string[] args)
        {
            if (args == null || args.Length == 0) return false;
            return string.Equals(args[0], "--list", StringComparison.OrdinalIgnoreCase)
                || string.Equals(args[0], "--create", StringComparison.OrdinalIgnoreCase);
        }
    }
}
