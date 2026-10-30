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
using System.Drawing.Drawing2D;
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

        /// <summary>声明本进程为 DPI 感知,避免高分屏上整窗被系统位图拉伸而发虚。</summary>
        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();
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
    // 五、GUI 模式(雾霭蓝调 · Pages 风格自绘界面)
    //     只负责呈现与交互,枚举/创建逻辑仍在 ShortcutTool 中,未作改动。
    // =================================================================

    /// <summary>主题令牌:改配色只需改这里。</summary>
    internal static class UiTheme
    {
        public static readonly Color Background = Color.FromArgb(0x6E, 0x8B, 0xB8);
        public static readonly Color BlobA = Color.FromArgb(40, 0x87, 0xA0, 0xC6);
        public static readonly Color BlobB = Color.FromArgb(58, 0xB9, 0xCB, 0xE2);
        public static readonly Color Card = Color.White;
        public static readonly Color Primary = Color.FromArgb(0x4E, 0x6E, 0x9E);
        public static readonly Color PrimaryHover = Color.FromArgb(0x44, 0x61, 0x8C);
        public static readonly Color PrimaryDisabled = Color.FromArgb(0xB9, 0xCB, 0xE2);
        public static readonly Color Tint = Color.FromArgb(0xEE, 0xF3, 0xFA);
        public static readonly Color Ink = Color.FromArgb(0x10, 0x18, 0x26);
        public static readonly Color Muted = Color.FromArgb(0x6B, 0x7C, 0x93);
        public static readonly Color Hint = Color.FromArgb(0x8A, 0x9A, 0xAE);
        public static readonly Color Divider = Color.FromArgb(0xE6, 0xED, 0xF5);
        public static readonly Color RowHover = Color.FromArgb(0xF5, 0xF8, 0xFC);
        public static readonly Color RowLine = Color.FromArgb(0xF0, 0xF4, 0xF9);
        public static readonly Color Success = Color.FromArgb(0x3F, 0x7D, 0x5C);
        public static readonly Color Warning = Color.FromArgb(0xA9, 0x77, 0x2B);
        public static readonly Color Danger = Color.FromArgb(0xB4, 0x42, 0x3E);
        public static readonly Color ChipDesktopBack = Color.FromArgb(0xF1, 0xEF, 0xE8);
        public static readonly Color ChipDesktopFore = Color.FromArgb(0x5F, 0x5E, 0x5A);
        public static readonly Color CheckBorder = Color.FromArgb(0xC3, 0xD0, 0xE0);
        public static readonly Color TitleBarText = Color.FromArgb(217, 255, 255, 255);

        public const string Family = "Microsoft YaHei UI";

        // 字体以 pt 为单位,系统 DPI 变化时由 GDI+ 自动换算像素,故这里不做手动缩放。
        public static readonly Font TitleFont = new Font(Family, 14F, FontStyle.Bold);
        public static readonly Font BodyFont = new Font(Family, 9F);
        public static readonly Font SmallFont = new Font(Family, 8.25F);
        public static readonly Font TinyFont = new Font(Family, 7.9F);
        public static readonly Font SearchFont = new Font(Family, 9.5F);

        /// <summary>读取系统 DPI 相对 96 的倍率(100% → 1.0,150% → 1.5),用于几何尺寸缩放。</summary>
        public static float DetectScale()
        {
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                float s = g.DpiX / 96F;
                return s <= 0 ? 1F : s;
            }
        }

        /// <summary>圆角矩形路径(radius 为圆角半径)。</summary>
        public static GraphicsPath RoundRect(Rectangle bounds, int radius)
        {
            var path = new GraphicsPath();
            if (radius <= 1 || bounds.Width <= 2 || bounds.Height <= 2)
            {
                path.AddRectangle(bounds);
                return path;
            }
            int d = Math.Min(radius * 2, Math.Min(bounds.Width, bounds.Height));
            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        public static void FillRound(Graphics g, Rectangle bounds, int radius, Color color)
        {
            using (GraphicsPath path = RoundRect(bounds, radius))
            using (var brush = new SolidBrush(color))
            {
                g.FillPath(brush, path);
            }
        }

        public static void DrawRound(Graphics g, Rectangle bounds, int radius, Color color, float width)
        {
            using (GraphicsPath path = RoundRect(bounds, radius))
            using (var pen = new Pen(color, width))
            {
                g.DrawPath(pen, path);
            }
        }
    }

    /// <summary>无边框胶囊主按钮。</summary>
    internal sealed class PillButton : Control
    {
        private bool _hover;

        public PillButton()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.UserPaint | ControlStyles.ResizeRedraw |
                     ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Cursor = Cursors.Hand;
            Font = UiTheme.SmallFont;
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            _hover = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _hover = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            Cursor = Enabled ? Cursors.Hand : Cursors.Default;
            Invalidate();
            base.OnEnabledChanged(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Color fill = !Enabled ? UiTheme.PrimaryDisabled : (_hover ? UiTheme.PrimaryHover : UiTheme.Primary);
            var bounds = new Rectangle(0, 0, Width - 1, Height - 1);
            UiTheme.FillRound(g, bounds, Height / 2, fill);
            TextRenderer.DrawText(g, Text, Font, bounds, Color.White,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }
    }

    /// <summary>无边框文字按钮(用于「全选」「刷新」)。</summary>
    internal sealed class TextButton : Control
    {
        private bool _hover;

        public TextButton()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.UserPaint | ControlStyles.ResizeRedraw |
                     ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Cursor = Cursors.Hand;
            Font = UiTheme.SmallFont;
            ForeColor = UiTheme.Primary;
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            _hover = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _hover = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnTextChanged(EventArgs e)
        {
            Invalidate();
            base.OnTextChanged(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var bounds = new Rectangle(0, 0, Width - 1, Height - 1);
            if (_hover) UiTheme.FillRound(g, bounds, 6, UiTheme.Tint);
            TextRenderer.DrawText(g, Text, Font, bounds, ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }
    }

    /// <summary>
    /// 自绘应用列表:表头 + 行(复选框 / 名称 / 类型胶囊 / 路径),支持滚轮滚动、悬停、双击直建。
    /// </summary>
    internal sealed class AppListView : Control
    {
        private readonly List<AppInfo> _rows = new List<AppInfo>();
        private readonly HashSet<AppInfo> _checked = new HashSet<AppInfo>();
        private int _scroll;
        private int _hover = -1;
        private float _scale = 1F;

        public event EventHandler<AppInfo> CreateRequested;
        public event EventHandler CheckedChanged;

        public AppListView()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            BackColor = UiTheme.Card;
        }

        public void SetScale(float scale)
        {
            _scale = scale <= 0 ? 1F : scale;
            Invalidate();
        }

        private int P(int value)
        {
            return (int)Math.Round(value * _scale);
        }

        private int RowHeight { get { return P(37); } }
        private int HeaderHeight { get { return P(26); } }
        private int VisibleRows { get { int h = Height - HeaderHeight; return h > 0 ? h / RowHeight : 0; } }

        private int NameX { get { return P(32); } }
        private int ChipX { get { return (int)(Width * 0.46); } }
        private int PathX { get { return (int)(Width * 0.60); } }

        public int RowCount { get { return _rows.Count; } }
        public int CheckedCount { get { return _checked.Count; } }

        public bool AllChecked
        {
            get
            {
                if (_rows.Count == 0) return false;
                foreach (AppInfo app in _rows)
                {
                    if (!_checked.Contains(app)) return false;
                }
                return true;
            }
        }

        public void SetRows(IEnumerable<AppInfo> rows)
        {
            _rows.Clear();
            _rows.AddRange(rows);
            _scroll = 0;
            _hover = -1;
            Invalidate();
        }

        public void ToggleAll(bool check)
        {
            if (check)
            {
                foreach (AppInfo app in _rows) _checked.Add(app);
            }
            else
            {
                foreach (AppInfo app in _rows) _checked.Remove(app);
            }
            Invalidate();
            RaiseCheckedChanged();
        }

        public List<AppInfo> GetChecked()
        {
            var result = new List<AppInfo>();
            foreach (AppInfo app in _rows)
            {
                if (_checked.Contains(app)) result.Add(app);
            }
            return result;
        }

        private void RaiseCheckedChanged()
        {
            if (CheckedChanged != null) CheckedChanged(this, EventArgs.Empty);
        }

        private int HitRow(int y)
        {
            if (y < HeaderHeight) return -1;
            int index = _scroll + (y - HeaderHeight) / RowHeight;
            if (index < 0 || index >= _rows.Count) return -1;
            return index;
        }

        /// <summary>取某个可视行号的矩形区域,用于局部重绘,避免整窗闪烁。</summary>
        private Rectangle RowRect(int visualIndex)
        {
            return new Rectangle(0, HeaderHeight + visualIndex * RowHeight, Width, RowHeight);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            int index = HitRow(e.Y);
            if (index != _hover)
            {
                int old = _hover;
                _hover = index;
                if (old != -1)
                {
                    int oldVisual = old - _scroll;
                    if (oldVisual >= 0 && oldVisual < VisibleRows) Invalidate(RowRect(oldVisual));
                }
                if (index != -1) Invalidate(RowRect(index - _scroll));
            }
            base.OnMouseMove(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            if (_hover != -1)
            {
                int oldVisual = _hover - _scroll;
                _hover = -1;
                if (oldVisual >= 0 && oldVisual < VisibleRows) Invalidate(RowRect(oldVisual));
            }
            base.OnMouseLeave(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            int index = HitRow(e.Y);
            if (index >= 0 && e.X <= P(30))
            {
                AppInfo app = _rows[index];
                if (!_checked.Remove(app)) _checked.Add(app);
                Invalidate(RowRect(index - _scroll));
                RaiseCheckedChanged();
            }
            base.OnMouseDown(e);
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            int index = HitRow(e.Y);
            if (index >= 0 && CreateRequested != null)
            {
                CreateRequested(this, _rows[index]);
            }
            base.OnMouseDoubleClick(e);
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            if (_rows.Count > VisibleRows)
            {
                int step = e.Delta > 0 ? -3 : 3;
                int max = _rows.Count - VisibleRows;
                _scroll = Math.Max(0, Math.Min(max, _scroll + step));
                Invalidate();
            }
            base.OnMouseWheel(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            using (var background = new SolidBrush(UiTheme.Card))
            {
                g.FillRectangle(background, ClientRectangle);
            }

            // 表头
            var headerFlags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding;
            TextRenderer.DrawText(g, "应用名称", UiTheme.SmallFont, new Rectangle(NameX, 0, ChipX - NameX - P(8), HeaderHeight), UiTheme.Hint, headerFlags);
            TextRenderer.DrawText(g, "类型", UiTheme.SmallFont, new Rectangle(ChipX, 0, PathX - ChipX - P(8), HeaderHeight), UiTheme.Hint, headerFlags);
            TextRenderer.DrawText(g, "路径 / AUMID", UiTheme.SmallFont, new Rectangle(PathX, 0, Width - PathX - P(8), HeaderHeight), UiTheme.Hint, headerFlags);
            using (var line = new Pen(UiTheme.Divider, 1F))
            {
                g.DrawLine(line, 0, HeaderHeight - 1, Width, HeaderHeight - 1);
            }

            int rowH = RowHeight;
            int visible = VisibleRows;
            int maxScroll = Math.Max(0, _rows.Count - visible);
            if (_scroll > maxScroll) _scroll = maxScroll;

            // 空结果提示
            if (_rows.Count == 0)
            {
                var emptyRect = new Rectangle(0, HeaderHeight, Width, Math.Max(P(80), Height - HeaderHeight));
                TextRenderer.DrawText(g, "未找到匹配的应用，换个关键词试试", UiTheme.BodyFont, emptyRect, UiTheme.Hint,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            }

            for (int i = 0; i <= visible && _scroll + i < _rows.Count; i++)
            {
                int index = _scroll + i;
                AppInfo app = _rows[index];
                var row = new Rectangle(0, HeaderHeight + i * rowH, Width - (maxScroll > 0 ? P(6) : 0), rowH);
                bool isChecked = _checked.Contains(app);
                bool isHover = index == _hover;

                if (isChecked)
                {
                    UiTheme.FillRound(g, new Rectangle(row.Left + P(4), row.Top + P(2), row.Width - P(8), rowH - P(4)), P(8), UiTheme.RowHover);
                    UiTheme.FillRound(g, new Rectangle(row.Left + P(4), row.Top + P(9), P(3), rowH - P(18)), P(2), UiTheme.Primary);
                }
                else if (isHover)
                {
                    UiTheme.FillRound(g, new Rectangle(row.Left + P(4), row.Top + P(2), row.Width - P(8), rowH - P(4)), P(8), UiTheme.RowHover);
                }

                // 复选框
                int boxSize = P(14);
                var box = new Rectangle(P(12), row.Top + (rowH - boxSize) / 2, boxSize, boxSize);
                if (isChecked)
                {
                    UiTheme.FillRound(g, box, P(4), UiTheme.Primary);
                    using (var check = new Pen(Color.White, Math.Max(1.2F, P(2) * 0.8F)))
                    {
                        g.DrawLines(check, new Point[]
                        {
                            new Point(box.Left + boxSize / 5, box.Top + boxSize / 2),
                            new Point(box.Left + boxSize * 2 / 5, box.Top + boxSize * 3 / 4),
                            new Point(box.Left + boxSize * 4 / 5, box.Top + boxSize / 4)
                        });
                    }
                }
                else
                {
                    UiTheme.DrawRound(g, box, P(4), UiTheme.CheckBorder, 1.2F);
                }

                // 名称
                var nameRect = new Rectangle(NameX, row.Top, ChipX - NameX - P(10), rowH);
                TextRenderer.DrawText(g, app.Name, UiTheme.BodyFont, nameRect, UiTheme.Ink,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis);

                // 类型胶囊
                string chipText;
                Color chipBack;
                Color chipFore;
                if (app.Kind == AppKind.Uwp)
                {
                    chipText = "UWP";
                    chipBack = UiTheme.Tint;
                    chipFore = UiTheme.Primary;
                }
                else if (app.Kind == AppKind.Desktop)
                {
                    chipText = "桌面";
                    chipBack = UiTheme.ChipDesktopBack;
                    chipFore = UiTheme.ChipDesktopFore;
                }
                else
                {
                    chipText = "不支持";
                    chipBack = UiTheme.ChipDesktopBack;
                    chipFore = UiTheme.Hint;
                }
                Size textSize = TextRenderer.MeasureText(g, chipText, UiTheme.TinyFont);
                int chipW = textSize.Width + P(14);
                int chipH = P(19);
                var chip = new Rectangle(ChipX, row.Top + (rowH - chipH) / 2, chipW, chipH);
                UiTheme.FillRound(g, chip, chipH / 2, chipBack);
                TextRenderer.DrawText(g, chipText, UiTheme.TinyFont, chip, chipFore,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);

                // 路径
                var pathRect = new Rectangle(PathX, row.Top, row.Width - PathX - P(8), rowH);
                TextRenderer.DrawText(g, app.Path ?? "", UiTheme.TinyFont, pathRect, UiTheme.Muted,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis);

                // 行分隔线
                if (!isChecked)
                {
                    using (var line = new Pen(UiTheme.RowLine, 1F))
                    {
                        g.DrawLine(line, P(12), row.Bottom - 1, row.Right - P(8), row.Bottom - 1);
                    }
                }
            }

            // 滚动条
            if (maxScroll > 0)
            {
                int trackTop = HeaderHeight;
                int trackHeight = Height - HeaderHeight;
                UiTheme.FillRound(g, new Rectangle(Width - P(4), trackTop + P(4), P(3), trackHeight - P(8)), P(2), UiTheme.Tint);
                int thumbH = Math.Max(P(24), trackHeight * visible / _rows.Count);
                int thumbY = trackTop + P(4) + (trackHeight - P(8) - thumbH) * _scroll / Math.Max(1, maxScroll);
                UiTheme.FillRound(g, new Rectangle(Width - P(4), thumbY, P(3), thumbH), P(2), UiTheme.PrimaryDisabled);
            }
        }
    }

    /// <summary>主窗口:自绘圆角窗口 + 白色卡片 + 搜索 + 列表 + 主按钮。</summary>
    internal sealed class MainForm : Form
    {
        private readonly List<AppInfo> _apps = new List<AppInfo>();
        private readonly TextBox _search = new TextBox();
        private readonly Label _searchHint = new Label();
        private readonly AppListView _list = new AppListView();
        private readonly PillButton _btnCreate = new PillButton();
        private readonly TextButton _btnSelectAll = new TextButton();
        private readonly TextButton _btnRefresh = new TextButton();

        private string _statusText = "";
        private Color _statusColor = UiTheme.Muted;
        private float _scale = 1F;

        private Rectangle _titleBar;
        private Rectangle _card;
        private Rectangle _searchRect;
        private Rectangle _btnMin;
        private Rectangle _btnClose;
        private int _hoverChrome = -1;

        /// <summary>「已选 N 项 · 共 M 项」文字区域,OnPaint 与局部重绘共用,避免坐标漂移。</summary>
        private Rectangle SelectionTextRect
        {
            get { return new Rectangle(_card.Left + _card.Width / 2, _searchRect.Bottom + P(4), _card.Width / 2 - P(18), P(24)); }
        }

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        public MainForm()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                     ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
            SetupWindow();
            BuildUi();
            LoadApps();
        }

        private int P(int value)
        {
            return (int)Math.Round(value * _scale);
        }

        private void SetupWindow()
        {
            Text = "桌面快捷方式生成器";
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            // 高 DPI 下按倍率放大整窗,所有坐标仍以 720×520 为设计稿基准。
            _scale = UiTheme.DetectScale();
            ClientSize = new Size(P(720), P(520));
            MinimumSize = new Size(P(720), P(520));
            MaximumSize = new Size(P(720), P(520));
            AutoScaleMode = AutoScaleMode.None;
            Font = UiTheme.BodyFont;
            KeyPreview = true;
            DoubleBuffered = true;
            ShowIcon = false;
        }

        private void BuildUi()
        {
            _scale = ClientSize.Width / 720F;
            _titleBar = new Rectangle(0, 0, ClientSize.Width, P(34));
            _card = new Rectangle(P(14), P(34), ClientSize.Width - P(28), ClientSize.Height - P(48));
            _btnMin = new Rectangle(ClientSize.Width - P(74), P(8), P(26), P(20));
            _btnClose = new Rectangle(ClientSize.Width - P(40), P(8), P(26), P(20));

            int pad = P(18);
            _searchRect = new Rectangle(_card.Left + pad, _card.Top + P(70), _card.Width - pad * 2, P(34));

            _search.BorderStyle = BorderStyle.None;
            _search.BackColor = UiTheme.Tint;
            _search.ForeColor = UiTheme.Ink;
            _search.Font = UiTheme.SearchFont;
            _search.Location = new Point(_searchRect.Left + P(32), _searchRect.Top + (_searchRect.Height - _search.PreferredHeight) / 2);
            _search.Width = _searchRect.Width - P(46);
            _search.TextChanged += OnSearchChanged;

            _searchHint.Text = "搜索应用（名称或路径）";
            _searchHint.AutoSize = false;
            _searchHint.BackColor = UiTheme.Tint;
            _searchHint.ForeColor = UiTheme.Hint;
            _searchHint.Font = UiTheme.SearchFont;
            _searchHint.TextAlign = ContentAlignment.MiddleLeft;
            _searchHint.Bounds = new Rectangle(_searchRect.Left + P(32), _searchRect.Top + P(2), _searchRect.Width - P(44), _searchRect.Height - P(4));
            _searchHint.Cursor = Cursors.IBeam;
            _searchHint.Click += delegate { _search.Focus(); };

            int toolbarY = _searchRect.Bottom + P(6);
            _btnSelectAll.Text = "全选";
            _btnSelectAll.Bounds = new Rectangle(_card.Left + pad - P(6), toolbarY, P(54), P(24));
            _btnSelectAll.Click += OnSelectAllClicked;

            _btnRefresh.Text = "刷新";
            _btnRefresh.Bounds = new Rectangle(_btnSelectAll.Right + P(4), toolbarY, P(54), P(24));
            _btnRefresh.Click += delegate { LoadApps(); };

            int listTop = toolbarY + P(28);
            int listBottom = _card.Bottom - P(46);
            _list.Bounds = new Rectangle(_card.Left + pad - P(8), listTop, _card.Width - (pad - P(8)) * 2, Math.Max(P(80), listBottom - listTop));
            _list.SetScale(_scale);
            _list.CreateRequested += delegate(object sender, AppInfo app) { CreateForApps(new AppInfo[] { app }); };
            _list.CheckedChanged += delegate { UpdateCounts(); };

            _btnCreate.Text = "创建到桌面";
            _btnCreate.Bounds = new Rectangle(_card.Right - pad - P(112), _card.Bottom - P(44), P(112), P(32));
            _btnCreate.Click += OnCreateClicked;

            // 加入顺序决定 z 序:先加的先在上层,占位提示需要盖住输入框
            Controls.Add(_searchHint);
            Controls.Add(_search);
            Controls.Add(_list);
            Controls.Add(_btnSelectAll);
            Controls.Add(_btnRefresh);
            Controls.Add(_btnCreate);

            Region = new Region(UiTheme.RoundRect(new Rectangle(0, 0, Width, Height), P(14)));
        }

        // ---------- 数据 ----------

        private void LoadApps()
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                List<AppInfo> apps = ShortcutTool.EnumerateApps();
                _apps.Clear();
                _apps.AddRange(apps);
                RebuildRows();
                _statusColor = UiTheme.Muted;
                _statusText = string.Format("共加载 {0} 个应用。勾选后点「创建到桌面」，或双击某一项直接创建。", _apps.Count);
            }
            catch (Exception ex)
            {
                _statusColor = UiTheme.Danger;
                _statusText = "加载应用列表失败：" + ex.Message;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            UpdateCounts();
            Invalidate();
        }

        private void RebuildRows()
        {
            string filter = _search.Text;
            var visible = new List<AppInfo>();
            foreach (AppInfo app in _apps)
            {
                if (MatchesFilter(app, filter)) visible.Add(app);
            }
            _list.SetRows(visible);
            if (_searchHint != null) _searchHint.Visible = _search.TextLength == 0;
            UpdateCounts();
        }

        private static bool MatchesFilter(AppInfo app, string filter)
        {
            if (string.IsNullOrWhiteSpace(filter)) return true;
            return (app.Name != null && app.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                || (app.Path != null && app.Path.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void UpdateCounts()
        {
            _btnSelectAll.Text = _list.AllChecked ? "取消全选" : "全选";
            Invalidate(SelectionTextRect);
        }

        // ---------- 事件 ----------

        private void OnSearchChanged(object sender, EventArgs e)
        {
            RebuildRows();
        }

        private void OnSelectAllClicked(object sender, EventArgs e)
        {
            _list.ToggleAll(!_list.AllChecked);
            UpdateCounts();
        }

        private void OnCreateClicked(object sender, EventArgs e)
        {
            List<AppInfo> selected = _list.GetChecked();
            if (selected.Count == 0)
            {
                _statusColor = UiTheme.Warning;
                _statusText = "请先勾选至少一个应用，或直接双击某一项。";
                Invalidate();
                return;
            }
            CreateForApps(selected);
        }

        /// <summary>为若干应用创建快捷方式。数量不大，同步执行 + 等待光标即可。</summary>
        private void CreateForApps(IEnumerable<AppInfo> apps)
        {
            var list = new List<AppInfo>(apps);
            int success = 0;
            string last = "";
            _btnCreate.Enabled = false;
            Cursor = Cursors.WaitCursor;
            try
            {
                foreach (AppInfo app in list)
                {
                    CreateResult result = ShortcutTool.CreateShortcut(app);
                    if (result.Success) success++;
                    last = string.Format("{0} {1}：{2}", result.Success ? "[成功]" : "[失败]", app.Name, result.Message);
                }
            }
            finally
            {
                _btnCreate.Enabled = true;
                Cursor = Cursors.Default;
            }
            int failed = list.Count - success;
            _statusColor = failed == 0 ? UiTheme.Success : (success == 0 ? UiTheme.Danger : UiTheme.Warning);
            _statusText = string.Format("成功 {0} 个，失败 {1} 个 · {2}", success, failed, last);
            Invalidate();
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape && _search.TextLength > 0)
            {
                _search.Clear();
                e.Handled = true;
            }
            base.OnKeyDown(e);
        }

        // ---------- 窗口拖动与自绘标题栏按钮 ----------

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (_btnClose.Contains(e.Location))
            {
                Close();
                return;
            }
            if (_btnMin.Contains(e.Location))
            {
                WindowState = FormWindowState.Minimized;
                return;
            }
            if (_titleBar.Contains(e.Location))
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            int hover = _btnClose.Contains(e.Location) ? 2 : (_btnMin.Contains(e.Location) ? 1 : 0);
            if (hover != _hoverChrome)
            {
                _hoverChrome = hover;
                Invalidate(new Rectangle(ClientSize.Width - P(84), 0, P(84), _titleBar.Bottom));
            }
            base.OnMouseMove(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            if (_hoverChrome != -1)
            {
                _hoverChrome = -1;
                Invalidate(new Rectangle(ClientSize.Width - P(84), 0, P(84), _titleBar.Bottom));
            }
            base.OnMouseLeave(e);
        }

        // ---------- 绘制 ----------

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            using (var brush = new SolidBrush(UiTheme.Background))
            {
                g.FillRectangle(brush, ClientRectangle);
            }
            using (var brush = new SolidBrush(UiTheme.BlobA))
            {
                g.FillEllipse(brush, new Rectangle(-P(90), P(130), P(320), P(240)));
            }
            using (var brush = new SolidBrush(UiTheme.BlobB))
            {
                g.FillEllipse(brush, new Rectangle(ClientSize.Width - P(250), P(30), P(320), P(260)));
            }

            UiTheme.FillRound(g, _card, P(16), UiTheme.Card);

            int pad = P(18);
            var titleRect = new Rectangle(_card.Left + pad, _card.Top + P(14), _card.Width - pad * 2 - P(96), P(28));
            TextRenderer.DrawText(g, "桌面快捷方式生成器", UiTheme.TitleFont, titleRect, UiTheme.Ink,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);

            var subtitleRect = new Rectangle(_card.Left + pad, _card.Top + P(42), _card.Width - pad * 2 - P(96), P(18));
            TextRenderer.DrawText(g, "为没有桌面图标的应用创建快捷方式", UiTheme.SmallFont, subtitleRect, UiTheme.Muted,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);

            // 应用总数胶囊
            string countText = string.Format("{0} 个应用", _apps.Count);
            Size countSize = TextRenderer.MeasureText(g, countText, UiTheme.SmallFont);
            int pillWidth = countSize.Width + P(20);
            int pillHeight = P(24);
            var countPill = new Rectangle(_card.Right - pad - pillWidth, _card.Top + P(18), pillWidth, pillHeight);
            UiTheme.FillRound(g, countPill, pillHeight / 2, UiTheme.Tint);
            TextRenderer.DrawText(g, countText, UiTheme.SmallFont, countPill, UiTheme.Primary,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);

            // 搜索框底 + 放大镜
            UiTheme.FillRound(g, _searchRect, _searchRect.Height / 2, UiTheme.Tint);
            DrawMagnifier(g, _searchRect.Left + P(14), _searchRect.Top + _searchRect.Height / 2, P(5));

            // 右侧已选/总数
            string selection = string.Format("已选 {0} 项 · 共 {1} 项", _list.CheckedCount, _apps.Count);
            var selectionRect = SelectionTextRect;
            TextRenderer.DrawText(g, selection, UiTheme.SmallFont, selectionRect, UiTheme.Muted,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);

            // 状态文字(单行省略)
            int statusWidth = _btnCreate.Left - _card.Left - pad - P(12);
            var statusRect = new Rectangle(_card.Left + pad, _card.Bottom - P(40), Math.Max(P(120), statusWidth), P(26));
            TextRenderer.DrawText(g, _statusText, UiTheme.SmallFont, statusRect, _statusColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis);

            // 自绘标题栏
            var titleBarText = new Rectangle(P(14), 0, ClientSize.Width - P(120), _titleBar.Bottom);
            TextRenderer.DrawText(g, "桌面快捷方式生成器", UiTheme.SmallFont, titleBarText, UiTheme.TitleBarText,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);

            DrawChromeButton(g, _btnMin, false, _hoverChrome == 1);
            DrawChromeButton(g, _btnClose, true, _hoverChrome == 2);
        }

        private void DrawChromeButton(Graphics g, Rectangle bounds, bool close, bool hover)
        {
            if (hover)
            {
                Color fill = close ? Color.FromArgb(70, 180, 66, 62) : Color.FromArgb(60, 255, 255, 255);
                UiTheme.FillRound(g, bounds, P(6), fill);
            }
            Color glyph = UiTheme.TitleBarText;
            using (var pen = new Pen(glyph, 1.3F))
            {
                int cx = bounds.Left + bounds.Width / 2;
                int cy = bounds.Top + bounds.Height / 2;
                if (close)
                {
                    int r = P(4);
                    g.DrawLine(pen, cx - r, cy - r, cx + r, cy + r);
                    g.DrawLine(pen, cx - r, cy + r, cx + r, cy - r);
                }
                else
                {
                    int r = P(5);
                    g.DrawLine(pen, cx - r, cy + P(1), cx + r, cy + P(1));
                }
            }
        }

        private static void DrawMagnifier(Graphics g, int centerX, int centerY, int radius)
        {
            using (var pen = new Pen(UiTheme.Hint, 1.4F))
            {
                g.DrawEllipse(pen, centerX - radius, centerY - radius, radius * 2, radius * 2);
                g.DrawLine(pen, centerX + radius - 1, centerY + radius - 1, centerX + radius + 3, centerY + radius + 3);
            }
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
            Native.SetProcessDPIAware();
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
