// ClipAngelDb — read-only queries against ClipAngel's live SQLite DB.
// Must be compiled x86 (System.Data.SQLite.dll from ClipAngel is 32-bit).
// Password is ClipAngel's hardcoded SQLite key (Main.cs Reputation = "Magic67234784").
// Exit 0 = ok, 1 = bad args, 2 = DB unavailable / query failure.

using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Text;
using System.Threading;

internal static class Program
{
    const string Password = "Magic67234784";
    const int DefaultWaitPollMs = 50;

    static int Main(string[] args)
    {
        if (args == null || args.Length < 1)
        {
            Console.Error.WriteLine("usage: ClipAngelDb <maxid|waitnew|newest|isfav|mergepayload> ...");
            return 1;
        }

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

    static string ResolveDbPath()
    {
        string env = Environment.GetEnvironmentVariable("CLIPANGEL_DB_PATH");
        if (!string.IsNullOrWhiteSpace(env))
            return env.Trim();
        string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return Path.Combine(local, "ClipAngel", "db.sqlite");
    }

    static SQLiteConnection OpenReadOnly()
    {
        string path = ResolveDbPath();
        if (!File.Exists(path))
            throw new FileNotFoundException("ClipAngel DB not found: " + path);
        // Read Only=True avoids write locks; Password matches ClipAngel Main.cs.
        string cs = "data source=" + path + ";Password=" + Password + ";Read Only=True;BusyTimeout=2000;";
        var con = new SQLiteConnection(cs);
        con.Open();
        return con;
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
