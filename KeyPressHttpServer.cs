using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;
using System.Resources;
using System.Collections.Generic;

public class KeyPressHttpServer : ApplicationContext
{
    [DllImport("user32.dll", SetLastError = true)]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        private const uint KEYEVENTF_KEYUP = 0x0002;

    private readonly HashSet<Keys> activeModifiers = new();
    private static readonly Dictionary<string, Keys> modifierMap = new()
    {
        { "Ctrl", Keys.ControlKey },
        { "Shift", Keys.ShiftKey },
        { "Alt", Keys.Menu },
        { "LWin", Keys.LWin },
        { "RWin", Keys.RWin }
    };


    private readonly HttpListener _listener = new HttpListener();
    private const int Port = 4664;

    // CHANGED: now defines multiple specific addresses instead of +
    private readonly string[] ServerUrls = new string[]
    {
        $"http://192.168.1.101:{Port}/cmd/",
        $"http://localhost:{Port}/cmd/",
        $"http://127.0.0.1:{Port}/cmd/"
    };

    private bool _isListening = false;

    private NotifyIcon trayIcon;

    public KeyPressHttpServer()
    {
        string icoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "icon.ico");
        Icon icon;
        try
        {
            var assembly = Assembly.GetExecutingAssembly();
            using Stream iconStream = assembly.GetManifestResourceStream("KeySenderApp.icon.ico");
            icon = new Icon(iconStream);
        }
        catch
        {
            icon = SystemIcons.Application;
        }

        trayIcon = new NotifyIcon
        {
            Icon = icon,
            ContextMenuStrip = new ContextMenuStrip(),
            Visible = true,
            Text = $"Listening on port {Port}"
        };

        // Updated to show all specific URLs instead of +
        trayIcon.ContextMenuStrip.Items.Add("Listening on:\n" + string.Join("\n", ServerUrls), null, ShowListeningInfo);
        trayIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
        trayIcon.ContextMenuStrip.Items.Add("Exit", null, ExitApplication);

        Task.Run(StartListener);
    }

    private void ShowListeningInfo(object sender, EventArgs e)
    {
        MessageBox.Show(
            "The HTTP Key Sender is running and listening for requests on:\n\n" +
            string.Join("\n", ServerUrls) +
            $"\n\nExample: http://localhost:{Port}/cmd?key=VKC1&mod=S",
            "Key Sender Service Information",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }

    private void ExitApplication(object sender, EventArgs e)
    {
        if (_isListening)
        {
            try { _listener.Stop(); } catch { }
            try { _listener.Close(); } catch { }
            _isListening = false;
        }

        if (trayIcon != null)
        {
            trayIcon.Visible = false;
            trayIcon.Dispose();
            trayIcon = null;
        }

        Application.Exit();
    }

    private void StartListener()
    {
        try
        {
            // CHANGED: add all desired prefixes
            foreach (var url in ServerUrls)
                _listener.Prefixes.Add(url);

            _listener.Start();
            _isListening = true;
            Console.WriteLine("Listening for requests on:");
            foreach (var url in ServerUrls) Console.WriteLine(url);

            while (_listener.IsListening)
            {
                HttpListenerContext context = null;
                try { context = _listener.GetContext(); }
                catch (HttpListenerException) { break; }
                catch (ObjectDisposedException) { break; }

                if (context != null)
                {
                    Task.Run(() => ProcessRequest(context));
                }
            }
        }
        catch (HttpListenerException ex) when (ex.ErrorCode == 5)
        {
            ShowError($"Failed to start HTTP listener on port {Port}. You may need to run this program as an administrator or configure URL reservations. Error: {ex.Message}");
        }
        catch (Exception ex)
        {
            if (_isListening)
            {
                ShowError($"An error occurred in the listener: {ex.Message}");
            }
        }
        finally
        {
            _isListening = false;
            try { _listener.Close(); } catch { }
        }
    }

private void ExecuteMacro(string rawMacro)
{
    string decoded = WebUtility.UrlDecode(rawMacro);
    var tokens = System.Text.RegularExpressions.Regex.Matches(decoded, @"\{.*?\}|.");

    foreach (System.Text.RegularExpressions.Match token in tokens)
    {
        string part = token.Value;

        if (part.StartsWith("{") && part.EndsWith("}"))
        {
            string content = part.Substring(1, part.Length - 2);
            string[] pieces = content.Split(' ');

            if (pieces.Length == 2 && (string.Equals(pieces[1], "Down", StringComparison.OrdinalIgnoreCase) || string.Equals(pieces[1], "Up", StringComparison.OrdinalIgnoreCase)))
            {
                string modName = pieces[0];
                if (modifierMap.TryGetValue(modName, out Keys modKey))
                {
                    bool isDown = string.Equals(pieces[1], "Down", StringComparison.OrdinalIgnoreCase);
                    keybd_event((byte)modKey, 0, isDown ? 0 : KEYEVENTF_KEYUP, UIntPtr.Zero);

                    if (isDown) activeModifiers.Add(modKey);
                    else activeModifiers.Remove(modKey);
                }
                else
                {
                    Keys key = ParseKey(modName);
                    bool isDown = string.Equals(pieces[1], "Down", StringComparison.OrdinalIgnoreCase);
                    keybd_event((byte)key, 0, isDown ? 0 : KEYEVENTF_KEYUP, UIntPtr.Zero);
                }
            }
            else
            {
                Keys key = ParseKey(content);
                foreach (var mod in activeModifiers)
                    keybd_event((byte)mod, 0, 0, UIntPtr.Zero); // Modifier down

                keybd_event((byte)key, 0, 0, UIntPtr.Zero);
                keybd_event((byte)key, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);

                foreach (var mod in activeModifiers)
                    keybd_event((byte)mod, 0, KEYEVENTF_KEYUP, UIntPtr.Zero); // Modifier up
            }
        }
        else
        {
            SendKeys.SendWait(part);
        }
    }

    activeModifiers.Clear(); // Clean up
}


    private Keys ParseKey(string name)
    {
        return Enum.TryParse(name, true, out Keys key) ? key : Keys.None;
    }

    private void ProcessRequest(HttpListenerContext context)
    {
        try
        {
            var query = context.Request.QueryString;
            string keyName = query["key"];
            string mod = query["mod"]?.ToUpperInvariant();
            string macro = query["macro"];
            string typeText = query["type"];
            ushort? modifierVK = GetModifierVK(mod);
            if (!string.IsNullOrEmpty(macro))
            {
                ExecuteMacro(macro);
                SendResponse(context, HttpStatusCode.OK, $"Macro executed: {macro}");
                Console.WriteLine($"Executed macro: {macro}");
                return;
            }

            if (!string.IsNullOrEmpty(typeText))
            {
                SendKeys.SendWait(typeText);
                SendResponse(context, HttpStatusCode.OK, $"Typed text: {typeText}");
                Console.WriteLine($"Typed text: {typeText}");
                return;
            }

            if (string.IsNullOrWhiteSpace(keyName))
            {
                SendResponse(context, HttpStatusCode.BadRequest, "Missing 'key' parameter.");
                return;
            }

            if (keyName.StartsWith("VK", StringComparison.OrdinalIgnoreCase))
            {
                string hexPart = keyName.Substring(2);
                if (ushort.TryParse(hexPart, System.Globalization.NumberStyles.HexNumber, null, out ushort vkCode))
                {
                    SimulateKeyPress(vkCode, modifierVK);
                    SendResponse(context, HttpStatusCode.OK, $"Raw VK code {keyName} simulated with modifier {mod ?? "none"}.");
                    Console.WriteLine($"Simulated raw VK: {keyName} ({vkCode}) with modifier {mod ?? "none"}");
                }
                else
                {
                    SendResponse(context, HttpStatusCode.BadRequest, $"Invalid VK hex code: {keyName}");
                }
            }
            else if (Enum.TryParse(keyName, true, out Keys virtualKey))
            {
                SimulateKeyPress((ushort)virtualKey, modifierVK);
                SendResponse(context, HttpStatusCode.OK, $"Named key {keyName} simulated with modifier {mod ?? "none"}.");
                Console.WriteLine($"Simulated named key: {keyName} ({virtualKey}) with modifier {mod ?? "none"}");
            }
            else
            {
                SendResponse(context, HttpStatusCode.BadRequest, $"Unrecognized key: {keyName}");
            }
        }
        catch (Exception ex)
        {
            SendResponse(context, HttpStatusCode.InternalServerError, $"Server processing error: {ex.Message}");
            ShowError($"Request processing error: {ex.Message}");
        }
    }

    private void SendResponse(HttpListenerContext context, HttpStatusCode statusCode, string message)
    {
        context.Response.StatusCode = (int)statusCode;
        if (!string.IsNullOrEmpty(message) && statusCode != HttpStatusCode.NoContent)
        {
            using (var writer = new StreamWriter(context.Response.OutputStream))
            {
                writer.Write(message);
            }
        }
        context.Response.Close();
    }

    private void ShowError(string message)
    {
        if (trayIcon?.ContextMenuStrip != null && trayIcon.ContextMenuStrip.InvokeRequired)
        {
            trayIcon.ContextMenuStrip.Invoke((System.Windows.Forms.MethodInvoker)delegate
            {
                MessageBox.Show(message, "Key Sender Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            });
        }
        else
        {
            MessageBox.Show(message, "Key Sender Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void SimulateKeyPress(ushort vkCode, ushort? modifierVK = null)
    {
        if (modifierVK.HasValue)
        {
            keybd_event((byte)modifierVK.Value, 0, 0, UIntPtr.Zero); // Modifier down
        }

        keybd_event((byte)vkCode, 0, 0, UIntPtr.Zero); // Key down
        keybd_event((byte)vkCode, 0, KEYEVENTF_KEYUP, UIntPtr.Zero); // Key up

        if (modifierVK.HasValue)
        {
            keybd_event((byte)modifierVK.Value, 0, KEYEVENTF_KEYUP, UIntPtr.Zero); // Modifier up
        }
    }

    private ushort? GetModifierVK(string mod)
    {
        return mod switch
        {
            "S" => 0x10, // VK_SHIFT
            "C" => 0x11, // VK_CONTROL
            "A" => 0x12, // VK_MENU (Alt)
            "W" => 0x5B, // VK_LWIN
            _ => null
        };
    }
}
