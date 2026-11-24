// ClipAngelDb — SQLite read/write helper for ClipAngel's live SQLite DB.
// Must be compiled x86 (System.Data.SQLite.dll from ClipAngel is 32-bit).
// Password is ClipAngel's hardcoded SQLite key (Main.cs Reputation = "Magic67234784").
// Exit 0 = ok, 1 = bad args, 2 = DB unavailable / query failure.

using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

internal static class Program
{
    const string Password = "Magic67234784";
    const int DefaultWaitPollMs = 50;

    delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    static extern int GetDlgCtrlID(IntPtr hWndCtl);

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    static extern IntPtr GetDlgItem(IntPtr hDlg, int nIDDlgItem);

    [DllImport("user32.dll")]
    static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll")]
    static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    static extern bool SetDllDirectory(string lpPathName);

    static int Main(string[] args)
    {
        if (args == null || args.Length < 1)
        {
            Console.Error.WriteLine("usage: ClipAngelDb <maxid|waitnew|newest|isfav|mergepayload|setfav|favnewest|unfavall|deletenonfav|deletenonfav-staged|confirmdelete> ...");
            return 1;
        }

        TryConfigureClipAngelDllDirectory();

        string cmd = args[0].Trim().ToLowerInvariant();
        try
        {
            switch (cmd)
            {
                case "maxid":
                    return CmdMaxId();
                case "waitnew":
                    if (args.Length < 3)
                    {
                        Console.Error.WriteLine("usage: ClipAngelDb waitnew <prevId> <timeoutMs>");
                        return 1;
                    }
                    return CmdWaitNew(ParseLong(args[1]), ParseInt(args[2]));
                case "newest":
                    return CmdNewest();
                case "isfav":
                    if (args.Length < 2)
                    {
                        Console.Error.WriteLine("usage: ClipAngelDb isfav <id>");
                        return 1;
                    }
                    return CmdIsFav(ParseLong(args[1]));
                case "mergepayload":
                    return CmdMergePayload();
                case "setfav":
                    if (args.Length < 3)
                    {
                        Console.Error.WriteLine("usage: ClipAngelDb setfav <id> <0|1>");
                        return 1;
                    }
                    return CmdSetFav(ParseLong(args[1]), ParseInt(args[2]) != 0 ? 1 : 0);
                case "favnewest":
                    return CmdFavNewest(args.Length >= 2 ? (ParseInt(args[1]) != 0 ? 1 : 0) : 1);
                case "unfavall":
                    return CmdUnfavAll();
                case "deletenonfav":
                    return CmdDeleteNonFav();
                case "deletenonfav-staged":
                    return CmdDeleteNonFavStaged();
                case "confirmdelete":
                    return CmdConfirmDelete(args.Length >= 2 ? ParseInt(args[1]) : 4000);
                default:
                    Console.Error.WriteLine("unknown command: " + cmd);
                    return 1;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 2;
        }
    }

    static long ParseLong(string s)
    {
        long v;
        if (!long.TryParse(s, out v))
            throw new ArgumentException("invalid integer: " + s);
        return v;
    }

    static int ParseInt(string s)
    {
        int v;
        if (!int.TryParse(s, out v))
            throw new ArgumentException("invalid integer: " + s);
        return v;
    }

    static string GetRunningClipAngelDir()
    {
        try
        {
            Process[] procs = Process.GetProcessesByName("ClipAngel");
            foreach (var p in procs)
            {
                try
                {
                    string exePath = p.MainModule.FileName;
                    if (!string.IsNullOrEmpty(exePath))
                        return Path.GetDirectoryName(exePath);
                }
                catch
                {
                }
            }
        }
        catch
        {
        }
        return "";
    }

    static void TryConfigureClipAngelDllDirectory()
    {
        try
        {
            string dir = GetRunningClipAngelDir();
            if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                SetDllDirectory(dir);
        }
        catch
        {
        }
    }

    // Finds and confirms any open ClipAngel #32770 confirmation dialog (OKCancel -> IDOK=1, YesNo -> IDYES=6, or Yes/Sim/OK text).
    static bool ConfirmAnyClipAngelDeleteDialogs()
    {
        var clipAngelPids = new HashSet<uint>();
        try
        {
            foreach (var p in Process.GetProcessesByName("ClipAngel"))
                clipAngelPids.Add((uint)p.Id);
        }
        catch
        {
        }
        if (clipAngelPids.Count == 0)
            return false;

        var dialogs = new List<IntPtr>();
        EnumWindows((hWnd, lParam) =>
        {
            uint pid;
            GetWindowThreadProcessId(hWnd, out pid);
            if (!clipAngelPids.Contains(pid))
                return true;
            var sb = new StringBuilder(64);
            if (GetClassName(hWnd, sb, sb.Capacity) > 0 && sb.ToString() == "#32770")
                dialogs.Add(hWnd);
            return true;
        }, IntPtr.Zero);

        bool closedAny = false;
        foreach (IntPtr dlg in dialogs)
        {
            int origExStyle = 0;
            bool madeLayered = false;
            try
            {
                origExStyle = GetWindowLong(dlg, -20);
                if ((origExStyle & 0x00080000) == 0)
                {
                    SetWindowLong(dlg, -20, origExStyle | 0x00080000); // WS_EX_LAYERED
                    SetLayeredWindowAttributes(dlg, 0, 0, 2);           // Opacity 0
                    madeLayered = true;
                }
            }
            catch
            {
            }

            IntPtr targetBtn = IntPtr.Zero;
            int targetCtrlId = 0;
            EnumChildWindows(dlg, (child, lParam) =>
            {
                var cls = new StringBuilder(32);
                if (GetClassName(child, cls, cls.Capacity) <= 0 || cls.ToString() != "Button")
                    return true;
                int id = GetDlgCtrlID(child);
                var txtSb = new StringBuilder(64);
                GetWindowText(child, txtSb, txtSb.Capacity);
                string txt = txtSb.ToString().Replace("&", "").Trim().ToLowerInvariant();
                if (txt == "no" || txt == "não" || txt == "nao" || txt == "cancel" || txt == "cancelar")
                    return true;
                if (id == 6 || id == 1 || txt == "yes" || txt == "sim" || txt == "ok")
                {
                    targetBtn = child;
                    targetCtrlId = id;
                    return false;
                }
                return true;
            }, IntPtr.Zero);

            IntPtr unused;
            SendMessageTimeout(dlg, 0x0006, new IntPtr(1), IntPtr.Zero, 2, 50, out unused); // WM_ACTIVATE (WA_ACTIVE)

            if (targetBtn != IntPtr.Zero)
            {
                SendMessageTimeout(targetBtn, 0x00F5, IntPtr.Zero, IntPtr.Zero, 2, 100, out unused); // BM_CLICK
                SendMessageTimeout(dlg, 0x0111, new IntPtr(targetCtrlId & 0xFFFF), targetBtn, 2, 100, out unused); // WM_COMMAND BN_CLICKED
            }
            else
            {
                IntPtr btnOk = GetDlgItem(dlg, 1);
                IntPtr btnYes = GetDlgItem(dlg, 6);
                if (btnYes != IntPtr.Zero)
                    SendMessageTimeout(btnYes, 0x00F5, IntPtr.Zero, IntPtr.Zero, 2, 100, out unused);
                if (btnOk != IntPtr.Zero)
                    SendMessageTimeout(btnOk, 0x00F5, IntPtr.Zero, IntPtr.Zero, 2, 100, out unused);
                PostMessage(dlg, 0x0111, new IntPtr(6), btnYes);
                PostMessage(dlg, 0x0111, new IntPtr(1), btnOk);
            }

            if (WaitDialogClosed(dlg, 120))
            {
                closedAny = true;
                continue;
            }

            // Fallback: activate briefly and send Enter / BM_CLICK
            try
            {
                SetForegroundWindow(dlg);
                if (targetBtn != IntPtr.Zero)
                    SendMessageTimeout(targetBtn, 0x00F5, IntPtr.Zero, IntPtr.Zero, 2, 100, out unused);
                PostMessage(dlg, 0x0100, new IntPtr(0x0D), IntPtr.Zero); // WM_KEYDOWN VK_RETURN
                PostMessage(dlg, 0x0101, new IntPtr(0x0D), IntPtr.Zero); // WM_KEYUP VK_RETURN
            }
            catch
            {
            }

            if (WaitDialogClosed(dlg, 120))
            {
                closedAny = true;
            }
            else if (madeLayered && IsWindow(dlg))
            {
                // Never leave a stuck modal dialog invisible
                try { SetWindowLong(dlg, -20, origExStyle); }
                catch { }
            }
        }
        return closedAny;
    }

    static bool WaitDialogClosed(IntPtr dlg, int timeoutMs)
    {
        int waitEnd = Environment.TickCount + timeoutMs;
        while (unchecked(Environment.TickCount - waitEnd) < 0)
        {
            if (!IsWindow(dlg))
                return true;
            Thread.Sleep(15);
        }
        return !IsWindow(dlg);
    }

    static int CmdConfirmDelete(int timeoutMs)
    {
        if (timeoutMs < 50)
            timeoutMs = 50;
        int deadline = Environment.TickCount + timeoutMs;
        while (true)
        {
            if (ConfirmAnyClipAngelDeleteDialogs())
            {
                Console.Write("1");
                return 0;
            }
            if (unchecked(Environment.TickCount - deadline) >= 0)
                return 2;
            Thread.Sleep(20);
        }
    }

    static List<string> ResolveDbPaths()
    {
        var paths = new List<string>();
        string env = Environment.GetEnvironmentVariable("CLIPANGEL_DB_PATH");
        if (!string.IsNullOrWhiteSpace(env))
        {
            string p = env.Trim();
            if (File.Exists(p))
                paths.Add(p);
        }

        string caDir = GetRunningClipAngelDir();
        if (!string.IsNullOrEmpty(caDir))
        {
            string portableDb = Path.Combine(caDir, "db.sqlite");
            if (File.Exists(portableDb) && !paths.Contains(portableDb))
                paths.Add(portableDb);
        }

        string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string localDb = Path.Combine(local, "ClipAngel", "db.sqlite");
        if (File.Exists(localDb) && !paths.Contains(localDb))
            paths.Add(localDb);

        paths.Sort((a, b) => File.GetLastWriteTimeUtc(b).CompareTo(File.GetLastWriteTimeUtc(a)));
        if (paths.Count == 0)
            paths.Add(localDb);
        return paths;
    }

    static string ResolveDbPath()
    {
        return ResolveDbPaths()[0];
    }

    static SQLiteConnection OpenReadOnly(string explicitPath = null)
    {
        string path = string.IsNullOrEmpty(explicitPath) ? ResolveDbPath() : explicitPath;
        if (!File.Exists(path))
            throw new FileNotFoundException("ClipAngel DB not found: " + path);
        // Read Only=True avoids write locks; Password matches ClipAngel Main.cs.
        string cs = "data source=" + path + ";Password=" + Password + ";Read Only=True;BusyTimeout=2000;";
        var con = new SQLiteConnection(cs);
        con.Open();
        return con;
    }

    static SQLiteConnection OpenWrite(string explicitPath = null)
    {
        string path = string.IsNullOrEmpty(explicitPath) ? ResolveDbPath() : explicitPath;
        if (!File.Exists(path))
            throw new FileNotFoundException("ClipAngel DB not found: " + path);
        string cs = "data source=" + path + ";Password=" + Password + ";Read Only=False;Pooling=False;BusyTimeout=500;";
        var con = new SQLiteConnection(cs);
        con.DefaultTimeout = 3;
        con.Open();
        ExecuteNonQuery(con, "PRAGMA busy_timeout = 500;");
        return con;
    }

    static int ExecuteNonQuery(SQLiteConnection con, string sql, params SQLiteParameter[] parms)
    {
        using (var cmd = con.CreateCommand())
        {
            cmd.CommandText = sql;
            if (parms != null)
            {
                foreach (var p in parms)
                    cmd.Parameters.Add(p);
            }
            return cmd.ExecuteNonQuery();
        }
    }

    static object Scalar(SQLiteConnection con, string sql, params SQLiteParameter[] parms)
    {
        using (var cmd = con.CreateCommand())
        {
            cmd.CommandText = sql;
            if (parms != null)
            {
                foreach (var p in parms)
                    cmd.Parameters.Add(p);
            }
            return cmd.ExecuteScalar();
        }
    }

    static int CmdMaxId()
    {
        using (var con = OpenReadOnly())
        {
            object o = Scalar(con, "SELECT MAX(Id) FROM Clips");
            long id = (o == null || o == DBNull.Value) ? 0L : Convert.ToInt64(o);
            Console.Write(id.ToString());
            return 0;
        }
    }

    static int CmdWaitNew(long prevId, int timeoutMs)
    {
        if (timeoutMs < 0)
            timeoutMs = 0;
        int deadline = Environment.TickCount + timeoutMs;
        using (var con = OpenReadOnly())
        {
            while (true)
            {
                object o = Scalar(con, "SELECT MAX(Id) FROM Clips");
                long id = (o == null || o == DBNull.Value) ? 0L : Convert.ToInt64(o);
                if (id > prevId)
                {
                    Console.Write(id.ToString());
                    return 0;
                }
                int now = Environment.TickCount;
                if (unchecked(now - deadline) >= 0)
                {
                    Console.Write(id.ToString());
                    return 2;
                }
                Thread.Sleep(DefaultWaitPollMs);
            }
        }
    }

    static int CmdNewest()
    {
        using (var con = OpenReadOnly())
        {
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText =
                    "SELECT Id, Type, Favorite, COALESCE(NULLIF(Title,''), SUBSTR(COALESCE(Text,''),1,120)) " +
                    "FROM Clips ORDER BY Id DESC LIMIT 1";
                using (var r = cmd.ExecuteReader())
                {
                    if (!r.Read())
                    {
                        Console.Write("0\t\t0\t");
                        return 0;
                    }
                    long id = Convert.ToInt64(r.GetValue(0));
                    string type = r.IsDBNull(1) ? "" : Convert.ToString(r.GetValue(1));
                    bool fav = !r.IsDBNull(2) && Convert.ToBoolean(r.GetValue(2));
                    string title = r.IsDBNull(3) ? "" : Convert.ToString(r.GetValue(3));
                    title = SanitizeOneLine(title);
                    Console.Write(id + "\t" + type + "\t" + (fav ? "1" : "0") + "\t" + title);
                    return 0;
                }
            }
        }
    }

    static int CmdIsFav(long id)
    {
        using (var con = OpenReadOnly())
        {
            object o = Scalar(con, "SELECT Favorite FROM Clips WHERE Id = @id",
                new SQLiteParameter("@id", id));
            if (o == null || o == DBNull.Value)
            {
                Console.Write("0");
                return 2;
            }
            bool fav = Convert.ToBoolean(o);
            Console.Write(fav ? "1" : "0");
            return 0;
        }
    }

    static int CmdMergePayload()
    {
        using (var con = OpenReadOnly())
        {
            object boundaryObj = Scalar(con, "SELECT MAX(Id) FROM Clips WHERE Favorite = 1");
            long boundaryId = (boundaryObj == null || boundaryObj == DBNull.Value)
                ? 0L : Convert.ToInt64(boundaryObj);

            var texts = new List<string>();
            int total = 0;
            int skipped = 0;

            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText =
                    "SELECT Id, Type, Text, HtmlText FROM Clips WHERE Id > @b ORDER BY Id ASC";
                cmd.Parameters.Add(new SQLiteParameter("@b", boundaryId));
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        total++;
                        string type = r.IsDBNull(1) ? "" : Convert.ToString(r.GetValue(1));
                        type = (type ?? "").Trim().ToLowerInvariant();
                        string text = r.IsDBNull(2) ? "" : Convert.ToString(r.GetValue(2));
                        string html = r.IsDBNull(3) ? "" : Convert.ToString(r.GetValue(3));

                        if (type == "img" || type == "file" || type == "image")
                        {
                            skipped++;
                            continue;
                        }

                        string body = text ?? "";
                        if (string.IsNullOrEmpty(body) && !string.IsNullOrEmpty(html))
                            body = StripHtmlRough(html);
                        if (string.IsNullOrEmpty(body))
                        {
                            skipped++;
                            continue;
                        }
                        texts.Add(body);
                    }
                }
            }

            // Header: boundaryId, totalCount (clips above favorite), skippedNonText
            Console.WriteLine(boundaryId + "\t" + total + "\t" + skipped);
            Console.Write(string.Join("\n", texts.ToArray()));
            return 0;
        }
    }

    static int CmdSetFav(long id, int fav)
    {
        using (var con = OpenWrite())
        {
            int rows = ExecuteNonQuery(con,
                "UPDATE Clips SET Favorite = @fav WHERE Id = @id",
                new SQLiteParameter("@fav", fav != 0),
                new SQLiteParameter("@id", id));
            Console.Write(rows + "\t" + id + "\t" + fav);
            return rows > 0 ? 0 : 2;
        }
    }

    static int CmdFavNewest(int fav)
    {
        using (var con = OpenWrite())
        {
            using (var tx = con.BeginTransaction())
            {
                long id = 0;
                string title = "";
                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText =
                        "SELECT Id, COALESCE(NULLIF(Title,''), SUBSTR(COALESCE(Text,''),1,120)) " +
                        "FROM Clips ORDER BY Id DESC LIMIT 1";
                    using (var r = cmd.ExecuteReader())
                    {
                        if (!r.Read())
                        {
                            tx.Rollback();
                            Console.Write("0\t0\t");
                            return 2;
                        }
                        id = Convert.ToInt64(r.GetValue(0));
                        title = r.IsDBNull(1) ? "" : SanitizeOneLine(Convert.ToString(r.GetValue(1)));
                    }
                }
                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = "UPDATE Clips SET Favorite = @fav WHERE Id = @id";
                    cmd.Parameters.Add(new SQLiteParameter("@fav", fav != 0));
                    cmd.Parameters.Add(new SQLiteParameter("@id", id));
                    cmd.ExecuteNonQuery();
                }
                tx.Commit();
                Console.Write(id + "\t" + fav + "\t" + title);
                return 0;
            }
        }
    }

    static int CmdUnfavAll()
    {
        using (var con = OpenWrite())
        {
            int rows = ExecuteNonQuery(con,
                "UPDATE Clips SET Favorite = 0 WHERE Favorite = 1");
            Console.Write(rows.ToString());
            return 0;
        }
    }

    static int CmdDeleteNonFav()
    {
        ConfirmAnyClipAngelDeleteDialogs();
        List<string> dbPaths = ResolveDbPaths();
        int totalDeleted = 0;
        long totalRemaining = 0;

        foreach (string dbPath in dbPaths)
        {
            if (!File.Exists(dbPath))
                continue;

            // Fast read-only pre-check: skip write lock if there are 0 non-favorite clips.
            try
            {
                using (var ro = OpenReadOnly(dbPath))
                {
                    object cntObj = Scalar(ro, "SELECT COUNT(*) FROM Clips WHERE NOT Favorite OR Favorite IS NULL");
                    long nonFavCount = (cntObj == null || cntObj == DBNull.Value) ? 0L : Convert.ToInt64(cntObj);
                    if (nonFavCount == 0)
                        continue;
                }
            }
            catch
            {
                // Fall through to write attempt.
            }

            for (int attempt = 0; attempt < 6; attempt++)
            {
                try
                {
                    ConfirmAnyClipAngelDeleteDialogs();
                    using (var con = OpenWrite(dbPath))
                    using (var tx = con.BeginTransaction())
                    {
                        int rows = 0;
                        using (var cmd = con.CreateCommand())
                        {
                            cmd.Transaction = tx;
                            cmd.CommandText = "DELETE FROM Clips WHERE NOT Favorite OR Favorite IS NULL";
                            rows = cmd.ExecuteNonQuery();
                        }

                        long remaining = 0;
                        using (var cmd = con.CreateCommand())
                        {
                            cmd.Transaction = tx;
                            cmd.CommandText = "SELECT COUNT(*) FROM Clips WHERE NOT Favorite OR Favorite IS NULL";
                            object o = cmd.ExecuteScalar();
                            remaining = (o == null || o == DBNull.Value) ? 0L : Convert.ToInt64(o);
                        }

                        tx.Commit();
                        totalDeleted += rows;
                        totalRemaining += remaining;
                        break;
                    }
                }
                catch (SQLiteException)
                {
                    if (attempt == 5)
                        throw;
                    Thread.Sleep(50 + (attempt * 50));
                }
            }
        }

        Console.Write(totalDeleted + "\t" + totalRemaining);
        return totalRemaining == 0 ? 0 : 2;
    }

    static int CmdDeleteNonFavStaged()
    {
        long preCount = -1;
        // Unstick any leftover #32770 confirmation dialog and pre-warm SQLite without holding a write lock.
        try
        {
            ConfirmAnyClipAngelDeleteDialogs();
            using (var con = OpenReadOnly())
            {
                object o = Scalar(con, "SELECT COUNT(*) FROM Clips WHERE NOT Favorite OR Favorite IS NULL");
                preCount = (o == null || o == DBNull.Value) ? 0L : Convert.ToInt64(o);
            }
        }
        catch
        {
            // Proceed anyway; CmdDeleteNonFav will retry or report busy.
        }

        Console.WriteLine("READY\t" + preCount);
        Console.Out.Flush();

        string line = Console.In.ReadLine();
        if (line == null || line.Trim() != "GO")
            return 2;

        return CmdDeleteNonFav();
    }

    static string SanitizeOneLine(string s)
    {
        if (string.IsNullOrEmpty(s))
            return "";
        s = s.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");
        return s;
    }

    static string StripHtmlRough(string html)
    {
        if (string.IsNullOrEmpty(html))
            return "";
        string s = System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", " ");
        s = System.Net.WebUtility.HtmlDecode(s);
        return s.Trim();
    }
}
