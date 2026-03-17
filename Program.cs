using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading;
using VM.Core;
using VM.PlatformSDKCS;
using IMVSCalibBoardCalibModuCs;
using ImageSourceModuleCs;

namespace VmCalibHelper
{
    class Program
    {
        private const int Port = 9101;
        private static bool _solutionLoaded;
        private static string _lastError = "";

        private static readonly string VmAppDir =
            @"C:\Program Files\VisionMaster4.4.0\Applications";
        private static readonly string[] VmLibDirs = new[]
        {
            @"C:\Program Files\VisionMaster4.4.0\Applications\myLibs",
            @"C:\Program Files\VisionMaster4.4.0\Applications",
            @"C:\Program Files\VisionMaster4.4.0\Applications\PublicFile\x64",
            @"C:\Program Files\VisionMaster4.4.0\Development\V4.x\ComControls\Assembly",
        };

        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;

            if (Directory.Exists(VmAppDir))
            {
                Environment.CurrentDirectory = VmAppDir;
                Log($"Working directory set to: {VmAppDir}");
            }
            else
            {
                Log($"WARNING: VisionMaster app dir not found: {VmAppDir}");
            }

            string solPath = args.Length > 0 ? args[0] : FindDefaultSolution();
            Log($"VmCalibHelper starting on port {Port}");

            if (!string.IsNullOrEmpty(solPath) && File.Exists(solPath))
            {
                var vmProcs = System.Diagnostics.Process.GetProcessesByName("VisionMaster");
                if (vmProcs.Length > 0)
                {
                    _lastError = "VisionMaster is running (PID " + vmProcs[0].Id +
                        "). Close VisionMaster before starting VmCalibHelper.";
                    Log($"ERROR: {_lastError}");
                }
                else
                {
                    try
                    {
                        Log($"Loading VM solution: {solPath}");
                        VmSolution.Load(solPath, "");
                        _solutionLoaded = true;
                        Log("VM solution loaded successfully.");
                    }
                    catch (TypeInitializationException ex)
                    {
                        string inner = ex.InnerException?.Message ?? ex.Message;
                        var deepInner = ex.InnerException?.InnerException;
                        _lastError = $"VM.Core init failed: {inner}";
                        Log($"Failed to load solution: {_lastError}");
                        if (deepInner != null)
                            Log($"  Deep inner: {deepInner.GetType().Name}: {deepInner.Message}");
                        Log("Hint: close VisionMaster and any other VM SDK process, then restart VmCalibHelper.");
                    }
                    catch (Exception ex)
                    {
                        _lastError = ex.Message;
                        Log($"Failed to load solution: {_lastError}");
                        if (ex.InnerException != null)
                            Log($"  Inner: {ex.InnerException.Message}");
                    }
                }
            }
            else
            {
                _lastError = solPath == null
                    ? "No .sol file found at default search paths."
                    : $"Sol file not found: {solPath}";
                Log($"ERROR: {_lastError}");
                Log("Expected at: C:\\Users\\Dennis WS\\Gantry Robot\\VMSolutions\\checkerboard_calib.sol");
            }

            RunTcpServer();
        }

        static string FindDefaultSolution()
        {
            string[] searchPaths = new[]
            {
                @"C:\Users\Dennis WS\Gantry Robot\VMSolutions\checkerboard_calib.sol",
                @"C:\Users\Dennis WS\Gantry Robot\VMSolutions\calib.sol",
            };
            foreach (var p in searchPaths)
            {
                if (File.Exists(p)) return p;
            }
            return null;
        }

        static void RunTcpServer()
        {
            var listener = new TcpListener(IPAddress.Loopback, Port);
            listener.Start();
            Log($"TCP server listening on 127.0.0.1:{Port}");

            while (true)
            {
                try
                {
                    var client = listener.AcceptTcpClient();
                    client.NoDelay = true;
                    Log("Client connected.");
                    ThreadPool.QueueUserWorkItem(_ => HandleClient(client));
                }
                catch (Exception ex)
                {
                    Log($"Accept error: {ex.Message}");
                }
            }
        }

        static void HandleClient(TcpClient client)
        {
            try
            {
                var stream = client.GetStream();
                stream.ReadTimeout = System.Threading.Timeout.Infinite;
                var reader = new StreamReader(stream, Encoding.UTF8);
                var writer = new StreamWriter(stream, Encoding.UTF8) { AutoFlush = true };

                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    Log($"Received: {line.Trim()}");
                    string response = ProcessCommand(line.Trim());
                    writer.WriteLine(response);
                    Log($"Sent: {(response.Length > 120 ? response.Substring(0, 120) + "..." : response)}");
                }
            }
            catch (Exception ex)
            {
                Log($"Client error: {ex.Message}");
            }
            finally
            {
                try { client.Close(); } catch { }
                Log("Client disconnected.");
            }
        }

        static string ProcessCommand(string command)
        {
            if (string.IsNullOrEmpty(command)) return "ERROR:EMPTY";

            string[] parts = command.Split(new[] { ':' }, 2);
            string cmd = parts[0].ToUpperInvariant();

            switch (cmd)
            {
                case "PING":
                    return "PONG";

                case "STATUS":
                    return _solutionLoaded ? "OK|LOADED" : $"ERROR|NOT_LOADED|{_lastError}";

                case "DETECT_GRID":
                    return HandleDetectGrid(parts.Length > 1 ? parts[1] : "");

                case "QUIT":
                    Log("QUIT received — shutting down.");
                    Environment.Exit(0);
                    return "OK";

                default:
                    return $"ERROR:UNKNOWN_COMMAND:{cmd}";
            }
        }

        static string HandleDetectGrid(string paramStr)
        {
            if (!_solutionLoaded)
                return "ERROR:VM solution not loaded";

            string[] p = paramStr.Split(',');
            if (p.Length < 3)
                return "ERROR:FORMAT — DETECT_GRID:imagePath,cols,rows[,sizeMm]";

            string imagePath = p[0].Trim();
            if (!int.TryParse(p[1].Trim(), out int cols) || !int.TryParse(p[2].Trim(), out int rows))
                return "ERROR:INVALID_DIMENSIONS";
            double sizeMm = (p.Length >= 4 && double.TryParse(p[3].Trim(), out double sz)) ? sz : 1.0;

            if (!File.Exists(imagePath))
                return $"ERROR:Image file not found: {imagePath}";

            try
            {
                // -- Step 1: Get the typed ImageSource module --
                var imgSrc = FindModule<ImageSourceModuleTool>("ImageSource1", "图像源1", "Image Source1");
                if (imgSrc == null)
                    return "ERROR:ImageSource module not found in solution";

                // Set LocalImage mode and path
                SetImageSourcePath(imgSrc, imagePath);

                // -- Step 2: Get the CalibBoard module and set params --
                var calibTool = FindModule<IMVSCalibBoardCalibModuTool>("CalibBoardCalib1", "标定板标定1", "CalibBoard Calib1");
                if (calibTool == null)
                    return "ERROR:CalibBoard module not found";

                var param = calibTool.ModuParams;
                param.PhysicalSize = sizeMm;
                param.CalibBoardType = CalibBoardCalibParam.CalibBoardTypeEnum.TypeChecker;
                param.FilterStatus = CalibBoardCalibParam.FilterStatusEnum.FilterStateTure;
                calibTool.ModuParams = param;

                // -- Step 3: Run the procedure (ImageSource feeds CalibBoard via subscription) --
                var proc = FindProcedure("流程1", "Procedure1", "Process1", "Flow1");
                if (proc != null)
                {
                    Log($"Running procedure '{proc}'...");
                    proc.Run();
                }
                else
                {
                    // Fallback: run modules individually
                    Log("No procedure found — running modules individually.");
                    imgSrc.Run();
                    // Set InputImage directly from ImageSource result
                    var imgData = imgSrc.ModuResult?.ImageData;
                    if (imgData != null)
                    {
                        var p2 = calibTool.ModuParams;
                        p2.InputImage = imgData;
                        calibTool.ModuParams = p2;
                        Log($"Set InputImage directly: {imgData.GetType().Name}");
                    }
                    calibTool.Run();
                }

                // -- Step 4: Read results --
                var result = calibTool.ModuResult;
                Log($"CalibBoard status={result.ModuStatus}, corners={result.CalibrationPoint?.Count ?? 0}, " +
                    $"RMS={result.EstimationError:F4}, Scale={result.Scale:F2}");

                // VM ModuStatus: 0=not run / init, 1=success, 2=fail
                // We ran it, so 1=OK, anything else=error
                if (result.ModuStatus == 2)
                    return $"ERROR:CalibBoard detection failed (status={result.ModuStatus})";

                var corners = result.CalibrationPoint;
                int count = corners != null ? corners.Count : 0;

                Log($"Detected {count} corners, RMS={result.EstimationError:F4}mm, Scale={result.Scale:F2}px/mm");
                Log($"  TranslateX={result.TranslateX:F4} TranslateY={result.TranslateY:F4}");
                Log($"  Rotate={result.Rotate:F6} Skew={result.Skew:F6} Aspect={result.Aspect:F6}");
                Log($"  Origin=({result.CalibrationOrigin.X:F3},{result.CalibrationOrigin.Y:F3})");
                Log($"  PosXVec=({result.PosXVector.X:F6},{result.PosXVector.Y:F6})");
                Log($"  PosYVec=({result.PosYVector.X:F6},{result.PosYVector.Y:F6})");
                Log($"  PixelPrecision={result.PixelPrecision:F4}");

                // Log V range to diagnose row clustering
                if (corners != null && corners.Count > 0)
                {
                    double vMin2 = double.MaxValue, vMax2 = double.MinValue;
                    foreach (var pt in corners) { if (pt.Y < vMin2) vMin2 = pt.Y; if (pt.Y > vMax2) vMax2 = pt.Y; }
                    Log($"Corner V range: {vMin2:F1} to {vMax2:F1} (span={vMax2-vMin2:F1}px)");
                }

                var sb = new StringBuilder();
                sb.Append($"OK|{count}|{result.EstimationError:F6}|{result.Scale:F6}");
                sb.Append($"|TX:{result.TranslateX:F6}|TY:{result.TranslateY:F6}");
                sb.Append($"|ROT:{result.Rotate:F8}|SKEW:{result.Skew:F8}|ASPECT:{result.Aspect:F8}");
                sb.Append($"|OX:{result.CalibrationOrigin.X:F4}|OY:{result.CalibrationOrigin.Y:F4}");
                sb.Append($"|XXV:{result.PosXVector.X:F8}|XYV:{result.PosXVector.Y:F8}");
                sb.Append($"|YXV:{result.PosYVector.X:F8}|YYV:{result.PosYVector.Y:F8}");
                sb.Append($"|PP:{result.PixelPrecision:F6}");
                if (corners != null)
                    foreach (var pt in corners)
                        sb.Append($"|{pt.X:F3},{pt.Y:F3}");
                return sb.ToString();
            }
            catch (VmException ex)
            {
                Log($"VM error: 0x{ex.errorCode:X} {ex.Message}");
                return $"ERROR:VM exception 0x{ex.errorCode:X}: {ex.Message}";
            }
            catch (Exception ex)
            {
                Log($"Error: {ex.Message}");
                if (ex.InnerException != null)
                    Log($"  Inner: {ex.InnerException.Message}");
                return $"ERROR:{ex.Message}";
            }
        }

        static void DumpSolutionContents()
        {
            try
            {
                var inst = VmSolution.Instance;
                string[] procNames = { "Procedure1", "流程1", "Process1", "Flow1", "流程 1", "Procedure 1" };
                foreach (var name in procNames)
                {
                    try
                    {
                        var obj = inst[name];
                        if (obj != null)
                        {
                            Log($"Found solution item '{name}': {obj.GetType().Name}");
                            string[] modNames = {
                                "ImageSource1", "图像源1", "Image Source1",
                                "CalibBoardCalib1", "标定板标定1", "CalibBoard Calib1",
                                name + ".ImageSource1", name + ".图像源1",
                                name + ".CalibBoardCalib1", name + ".标定板标定1"
                            };
                            foreach (var mn in modNames)
                            {
                                try { var m = inst[mn]; if (m != null) Log($"  Module '{mn}': {m.GetType().Name}"); }
                                catch { }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Log($"DumpSolution error: {ex.Message}"); }
        }

        static T FindModule<T>(params string[] names) where T : class
        {
            foreach (var name in names)
            {
                try
                {
                    // Try with procedure prefix
                    string[] prefixes = { "流程1.", "Procedure1.", "Process1.", "" };
                    foreach (var prefix in prefixes)
                    {
                        try
                        {
                            var obj = VmSolution.Instance[prefix + name];
                            if (obj is T typed) return typed;
                            if (obj != null) return obj as T;
                        }
                        catch { }
                    }
                }
                catch { }
            }
            return null;
        }

        static VmProcedure FindProcedure(params string[] names)
        {
            foreach (var name in names)
            {
                try
                {
                    var obj = VmSolution.Instance[name] as VmProcedure;
                    if (obj != null) return obj;
                }
                catch { }
            }
            return null;
        }

        static void SetImageSourcePath(object imgSrcObj, string path)
        {
            // Use typed API when possible
            var imgSrc = imgSrcObj as ImageSourceModuleTool;
            if (imgSrc != null)
            {
                var p = imgSrc.ModuParams;
                // Try setting typed ImageSourceType to LocalImage
                try
                {
                    p.ImageSourceType = ImageSourceParam.ImageSourceTypeEnum.LocalImage;
                    imgSrc.ModuParams = p;
                    Log("Set ImageSourceType = LocalImage (typed)");
                }
                catch (Exception ex) { Log($"ImageSourceType set failed: {ex.Message}"); }

                // Use typed SetImagePath method
                imgSrc.SetImagePath(path);
                Log($"Called SetImagePath({path}) (typed)");
                return;
            }

            // Reflection fallback
            var type = imgSrcObj.GetType();
            try
            {
                var paramsProp = type.GetProperty("ModuParams");
                if (paramsProp != null)
                {
                    var moduParams = paramsProp.GetValue(imgSrcObj);
                    var srcTypeProp = moduParams?.GetType().GetProperty("ImageSourceType");
                    if (srcTypeProp != null)
                    {
                        var enumType = srcTypeProp.PropertyType;
                        srcTypeProp.SetValue(moduParams, Enum.Parse(enumType, "LocalImage"));
                        paramsProp.SetValue(imgSrcObj, moduParams);
                        Log("Set ImageSourceType = LocalImage (reflection)");
                    }
                }
            }
            catch (Exception ex) { Log($"ImageSourceType reflection failed: {ex.Message}"); }

            var setPathMethod = type.GetMethod("SetImagePath", new[] { typeof(string) });
            if (setPathMethod != null)
            {
                setPathMethod.Invoke(imgSrcObj, new object[] { path });
                Log($"Called SetImagePath({path}) (reflection)");
            }
            else
            {
                Log("WARNING: SetImagePath not found.");
            }
        }

        static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {
            string name = new System.Reflection.AssemblyName(args.Name).Name + ".dll";
            foreach (var dir in VmLibDirs)
            {
                string path = Path.Combine(dir, name);
                if (File.Exists(path))
                {
                    Log($"Resolved assembly: {name} from {dir}");
                    return Assembly.LoadFrom(path);
                }
            }
            return null;
        }

        static void Log(string msg)
        {
            string line = $"[{DateTime.Now:HH:mm:ss.fff}] {msg}";
            Console.WriteLine(line);
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "vmcalibhelper.log");
                File.AppendAllText(logPath, line + "\n");
            }
            catch { }
        }
    }
}
