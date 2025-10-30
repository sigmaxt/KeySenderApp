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
using System.Threading;

public class KeyPressHttpServer : ApplicationContext
{
    [DllImport("user32.dll", SetLastError = true)]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    static extern uint MapVirtualKey(uint uCode, uint uMapType);

    [DllImport("user32.dll", SetLastError = true)]
    static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [StructLayout(LayoutKind.Sequential)]
    struct INPUT
    {
        public uint type;
        public InputUnion u;
    }

    [StructLayout(LayoutKind.Explicit)]
    struct InputUnion
    {
        [FieldOffset(0)] public KEYBDINPUT ki;
        [FieldOffset(0)] public MOUSEINPUT mi;
        [FieldOffset(0)] public HARDWAREINPUT hi;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct HARDWAREINPUT
    {
        public uint uMsg;
        public ushort wParamL;
        public ushort wParamH;
    }

    private const uint INPUT_KEYBOARD = 1;

    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const uint KEYEVENTF_SCANCODE = 0x0008;

    private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;

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
        string icoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "VKsender.icon.ico");
        Icon icon;
        try
        {
            var assembly = Assembly.GetExecutingAssembly();
            using Stream iconStream = assembly.GetManifestResourceStream("VKsender.icon.ico");
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
        "The HTTP Key Sender is running and listening for requests on:\n" +
        string.Join("\n", ServerUrls) +
        $"\n\nExamples:\n" +
        $"  http://localhost:{Port}/cmd?key=VKC1&mod=S\n" +
        $"  http://localhost:{Port}/cmd?macro={{Shift Down}}{{Ctrl Down}}{{Escape}}{{Shift Up}}{{Ctrl Up}}\n" +
        $"  http://localhost:{Port}/cmd?macro=This{{Space}}is{{Space}}Text",
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
        var tokens = System.Text.RegularExpressions.Regex.Matches(decoded, @"\{.*?\}|

\[.*?\]

|.");

        HashSet<Keys> heldModifiers = new();

        foreach (System.Text.RegularExpressions.Match token in tokens)
        {
            string part = token.Value;

            if (part.StartsWith("{") && part.EndsWith("}"))
            {
                string content = part.Substring(1, part.Length - 2);
                string[] pieces = content.Split(' ');

                if (pieces.Length == 2 &&
                    (string.Equals(pieces[1], "Down", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(pieces[1], "Up", StringComparison.OrdinalIgnoreCase)))
                {
                    string modName = pieces[0];
                    bool isDown = string.Equals(pieces[1], "Down", StringComparison.OrdinalIgnoreCase);

                    if (modifierMap.TryGetValue(modName, out Keys modKey))
                    {
                        QueueKey(modKey, isDown);
                        if (isDown) heldModifiers.Add(modKey);
                        else heldModifiers.Remove(modKey);
                    }
                    else
                    {
                        Keys parsedKey = ParseKey(modName);
                        QueueKey(parsedKey, isDown);
                    }
                }
                else
                {
                    Keys parsedKey = ParseKey(content);
                    QueueKey(parsedKey, true);
                    QueueKey(parsedKey, false);
                }
            }
            else if (part.StartsWith("[") && part.EndsWith("]"))
            {
                // Send literal block through the SendInput queue so spaces are handled correctly.
                string literal = part.Substring(1, part.Length - 2);

                foreach (char ch in literal)
                {
                    if (ch == ' ')
                    {
                        QueueKey(Keys.Space, true);
                        QueueKey(Keys.Space, false);
                        continue;
                    }

                    // Attempt to map printable character to a Keys value
                    string s = ch.ToString();
                    Keys parsedKey = ParseKey(s);

                    // If ParseKey failed, try upper-case mapping for letters
                    if (parsedKey == Keys.None && char.IsLetter(ch))
                    {
                        parsedKey = ParseKey(s.ToUpperInvariant());
                    }

                    // If still None, fall back to ASCII-to-key mapping for common characters
                    if (parsedKey == Keys.None)
                    {
                        switch (ch)
                        {
                            case '.': parsedKey = Keys.OemPeriod; break;
                            case ',': parsedKey = Keys.Oemcomma; break;
                            case ';': parsedKey = Keys.Oem1; break;
                            case ':': parsedKey = Keys.Oem1; break;
                            case '/': parsedKey = Keys.Oem2; break;
                            case '?': parsedKey = Keys.Oem2; break;
                            case '\\': parsedKey = Keys.Oem5; break;
                            case '-': parsedKey = Keys.OemMinus; break;
                            case '_': parsedKey = Keys.OemMinus; break;
                            case '=': parsedKey = Keys.Oemplus; break;
                            case '+': parsedKey = Keys.Oemplus; break;
                            case '\'': parsedKey = Keys.Oem7; break;
                            case '"': parsedKey = Keys.Oem7; break;
                            case '[': parsedKey = Keys.Oem4; break;
                            case ']': parsedKey = Keys.Oem6; break;
                            case '`': parsedKey = Keys.Oem3; break;
                            case '~': parsedKey = Keys.Oem3; break;
                            case '!': parsedKey = Keys.D1; break;
                            case '@': parsedKey = Keys.D2; break;
                            case '#': parsedKey = Keys.D3; break;
                            case '$': parsedKey = Keys.D4; break;
                            case '%': parsedKey = Keys.D5; break;
                            case '^': parsedKey = Keys.D6; break;
                            case '&': parsedKey = Keys.D7; break;
                            case '*': parsedKey = Keys.D8; break;
                            case '(': parsedKey = Keys.D9; break;
                            case ')': parsedKey = Keys.D0; break;
                            case '0': parsedKey = Keys.D0; break;
                            case '1': parsedKey = Keys.D1; break;
                            case '2': parsedKey = Keys.D2; break;
                            case '3': parsedKey = Keys.D3; break;
                            case '4': parsedKey = Keys.D4; break;
                            case '5': parsedKey = Keys.D5; break;
                            case '6': parsedKey = Keys.D6; break;
                            case '7': parsedKey = Keys.D7; break;
                            case '8': parsedKey = Keys.D8; break;
                            case '9': parsedKey = Keys.D9; break;
                        }
                    }

                    if (parsedKey == Keys.None)
                    {
                        // As a last resort, skip unknown characters to avoid sending invalid codes.
                        Console.WriteLine($"[ExecuteMacro] Skipping unsupported literal character: '{ch}' (0x{((int)ch):X})");
                        continue;
                    }

                    // For letters and shifted characters you may need to send shift down/up around the key,
                    // but if the target expects case-sensitive input you can rely on existing modifier logic.
                    QueueKey(parsedKey, true);
                    QueueKey(parsedKey, false);
                }
            }
            else
            {
                Keys parsedKey = ParseKey(part);
                QueueKey(parsedKey, true);
                QueueKey(parsedKey, false);
            }
        }

        FlushInputQueue();
    }



    private List<INPUT> inputQueue = new();

    private void QueueKey(Keys key, bool isDown)
    {
        ushort scanCode = (ushort)MapVirtualKey((uint)key, 0);

        uint flags = KEYEVENTF_SCANCODE | (isDown ? 0 : KEYEVENTF_KEYUP);

        // Add extended flag for arrow keys and others
        if (key == Keys.Up || key == Keys.Down || key == Keys.Left || key == Keys.Right ||
            key == Keys.Home || key == Keys.End || key == Keys.Insert || key == Keys.Delete ||
            key == Keys.PageUp || key == Keys.PageDown)
        {
            flags |= KEYEVENTF_EXTENDEDKEY;
        }

        INPUT input = new INPUT
        {
            type = INPUT_KEYBOARD,
            u = new InputUnion
            {
                ki = new KEYBDINPUT
                {
                    wVk = 0,
                    wScan = scanCode,
                    dwFlags = flags,
                    time = 0,
                    dwExtraInfo = UIntPtr.Zero
                }
            }
        };

        inputQueue.Add(input);
    }



    private void FlushInputQueue()
    {
        if (inputQueue.Count > 0)
        {
            uint result = SendInput((uint)inputQueue.Count, inputQueue.ToArray(), Marshal.SizeOf(typeof(INPUT)));
            Console.WriteLine($"[FlushInputQueue] Sent {inputQueue.Count} events, result: {result}");
            inputQueue.Clear();
        }
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
