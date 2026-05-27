using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using VM.Core;
using VM.PlatformSDKCS;
using IMVSCalibBoardCalibModuCs;
using IMVSContourMatchModuCs;
using IMVSFastFeatureMatchModuCs;
using ImageSourceModuleCs;

// MVDAlgorithmSDK — see HandleMatchQuick comment for rationale.
using VisionDesigner;
using VisionDesigner.FastFeaturePatMatch;

namespace VmCalibHelper
{
    class Program
    {
        private const int Port = 9101;
        private static bool _solutionLoaded;
        private static string _lastError = "";
        private static string _loadedSolPath = "";

        private static readonly string VmAppDir =
            @"C:\Program Files\VisionMaster4.4.0\Applications";
        private static readonly string[] VmLibDirs = new[]
        {
            @"C:\Program Files\VisionMaster4.4.0\Applications\myLibs",
            @"C:\Program Files\VisionMaster4.4.0\Applications",
            @"C:\Program Files\VisionMaster4.4.0\Applications\PublicFile\x64",
            @"C:\Program Files\VisionMaster4.4.0\Development\V4.x\ComControls\Assembly",
            // MVDAlgorithmSDK managed wrappers — referenced by HandleMatchQuick.
            // Listed AFTER the VM lib dirs so that for assemblies present in
            // both trees (e.g. MVDCore.Net.dll) we prefer the VM-bundled copy,
            // keeping bit-for-bit compatibility with VM.PlatformSDKCS, which
            // was already initialised at startup. (They are the same 4.2.1.4
            // build, but using the VM copy means there is no chance of two
            // distinct Assembly objects for the same identity floating in the
            // AppDomain.)
            @"C:\Program Files (x86)\MVDAlgorithmSDK\ReferencedAssemblies\Common",
            @"C:\Program Files (x86)\MVDAlgorithmSDK\ReferencedAssemblies\Algorithms",
        };

        // Where the MVDAlgorithmSDK native DLLs live. The MSIL wrapper
        // assemblies P/Invoke into MVDFastFeaturePatMatchCpp.dll, MVDImageCpp.dll
        // and MVDShapeCpp.dll. The VM's myLibs dir does NOT contain them, so
        // we have to prepend this to PATH so LoadLibrary finds them. (PATH is
        // process-scoped here, won't leak.)
        private static readonly string MvdAlgoRuntimeDir =
            @"C:\Program Files (x86)\MVDAlgorithmSDK\Runtime\x64";

        // SetDllDirectory only lets us set ONE extra search dir and we already
        // need myLibs for VM. We tweak PATH instead, which the loader also
        // honours and which composes with whatever the user already has.
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetDllDirectoryW(string lpPathName);

        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;

            // Make MVDFastFeaturePatMatchCpp.dll & friends discoverable BEFORE
            // we touch any VisionDesigner.* type (else the JIT-time P/Invoke
            // resolution will fail with "DLL not found"). PATH is the most
            // forgiving option here — SetDllDirectory replaces (rather than
            // adds), and we already rely on myLibs being on the loader's
            // search path via working-directory below.
            if (Directory.Exists(MvdAlgoRuntimeDir))
            {
                string curPath = Environment.GetEnvironmentVariable("PATH") ?? "";
                Environment.SetEnvironmentVariable("PATH", MvdAlgoRuntimeDir + ";" + curPath);
                Log($"MVDAlgorithmSDK runtime dir added to PATH: {MvdAlgoRuntimeDir}");
            }
            else
            {
                Log($"WARNING: MVDAlgorithmSDK runtime not found at {MvdAlgoRuntimeDir} — " +
                    "MATCH_QUICK will fail with DllNotFound until the SDK is installed.");
            }

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
                    TryLoadSolution(solPath);
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

        // Encapsulates the actual VmSolution.Load() call plus a verbose
        // error walk. The original implementation hid the real cause behind
        // a generic "Error in the application." log line — VmException
        // carries an HRESULT-like errorCode that often pinpoints the issue
        // (license server, hardware key, stale shared memory, etc.), and
        // chained InnerExceptions sometimes wrap the *real* error several
        // layers deep. Walk the whole chain and log everything.
        static bool TryLoadSolution(string solPath)
        {
            try
            {
                Log($"Loading VM solution: {solPath}");
                VmSolution.Load(solPath, "");
                _solutionLoaded = true;
                _loadedSolPath = solPath;
                _lastError = "";
                Log("VM solution loaded successfully.");
                return true;
            }
            catch (TypeInitializationException ex)
            {
                _lastError = $"VM.Core init failed: {ex.InnerException?.Message ?? ex.Message}";
                Log($"Failed to load solution: {_lastError}");
                DumpExceptionChain(ex);
                Log("Hint: close VisionMaster + any other VM SDK process, then retry. " +
                    "If the issue persists, check VM license dongle / VmDriver service.");
                return false;
            }
            catch (VmException ex)
            {
                _lastError = $"VM error 0x{ex.errorCode:X}: {ex.Message}";
                Log($"Failed to load solution: {_lastError}");
                DumpExceptionChain(ex);
                return false;
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
                Log($"Failed to load solution ({ex.GetType().Name}): {_lastError}");
                DumpExceptionChain(ex);
                return false;
            }
        }

        // Walk Inner/Inner/Inner... and emit Type+Message+first 3 stack
        // frames for each. The VM SDK loves wrapping the meaningful error
        // (e.g. "license check failed", "shared memory in use") inside two
        // or three layers of generic VmException, and the previous logger
        // only printed the outermost layer.
        static void DumpExceptionChain(Exception ex)
        {
            int depth = 0;
            for (var e = ex; e != null && depth < 10; e = e.InnerException, depth++)
            {
                string prefix = new string(' ', depth * 2);
                Log($"{prefix}[{depth}] {e.GetType().FullName}: {e.Message}");
                if (e is VmException vex)
                    Log($"{prefix}    errorCode=0x{vex.errorCode:X}");
                if (!string.IsNullOrEmpty(e.StackTrace))
                {
                    // Just the top few frames — the rest is .NET plumbing.
                    var frames = e.StackTrace.Split('\n');
                    int take = Math.Min(3, frames.Length);
                    for (int i = 0; i < take; i++)
                        Log($"{prefix}    {frames[i].TrimEnd()}");
                }
            }
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

                case "RELOAD":
                    // Manual reload trigger. Useful when a VM SDK environmental
                    // issue (license dongle reconnect, VisionMaster being
                    // closed, etc.) is fixed and we don't want to restart the
                    // whole helper process. parts[1] (if present) is an
                    // alternate sol path.
                    {
                        string sol = parts.Length > 1 && !string.IsNullOrWhiteSpace(parts[1])
                            ? parts[1].Trim()
                            : (string.IsNullOrEmpty(_loadedSolPath) ? FindDefaultSolution() : _loadedSolPath);
                        if (string.IsNullOrEmpty(sol))
                            return "ERROR:No sol path available — pass RELOAD:<path>";
                        if (_solutionLoaded)
                        {
                            try { VmSolution.Instance.CloseSolution(); }
                            catch (Exception ex) { Log($"Close before reload warning: {ex.Message}"); }
                            _solutionLoaded = false;
                            _loadedSolPath = "";
                        }
                        return TryLoadSolution(sol)
                            ? $"OK|RELOADED|{sol}"
                            : $"ERROR|RELOAD_FAILED|{_lastError}";
                    }

                case "DETECT_GRID":
                    return HandleDetectGrid(parts.Length > 1 ? parts[1] : "");

                case "MATCH_CONTOUR":
                    return HandleMatchContour(parts.Length > 1 ? parts[1] : "");

                case "MATCH_QUICK":
                    return HandleMatchQuick(parts.Length > 1 ? parts[1] : "");

                case "MATCH_PREVIEW":
                    // Same template/match pipeline as MATCH_QUICK but the
                    // response ALSO carries the trained feature outline, so
                    // the VisionBridge "Preview…" popup can render the green
                    // edge-point overlay alongside the cyan match boxes.
                    return HandleMatchPreview(parts.Length > 1 ? parts[1] : "");

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

        static readonly string DefaultPartLocateSol =
            @"C:\Users\Dennis WS\Gantry Robot\VMSolutions\part_locate.sol";

        // NOTE: DefaultPartMatchingSol was removed when MATCH_QUICK was
        // rewritten to use the MVDAlgorithmSDK (CFastFeaturePattern.Train).
        // The new path no longer needs a baked-template .sol — the user-drawn
        // ROI is the template, re-trained per call. RELOAD:<path> still works
        // for any caller that wants to swap solutions, but VisionBridge's
        // VmQuickMatchClient should drop its Reload() call (or treat it as
        // a no-op) since there is nothing to reload.

        static bool EnsureSolutionLoaded(string solPath)
        {
            if (_solutionLoaded && _loadedSolPath == solPath)
                return true;

            // Unload current solution before loading a different one
            if (_solutionLoaded)
            {
                try
                {
                    Log($"Unloading current solution: {_loadedSolPath}");
                    VmSolution.Instance.CloseSolution();
                    _solutionLoaded = false;
                    _loadedSolPath = "";
                }
                catch (Exception ex) { Log($"Unload warning: {ex.Message}"); }
            }

            if (!File.Exists(solPath))
            {
                _lastError = $"Solution file not found: {solPath}";
                Log($"ERROR: {_lastError}");
                return false;
            }

            try
            {
                Log($"Loading VM solution: {solPath}");
                VmSolution.Load(solPath, "");
                _solutionLoaded = true;
                _loadedSolPath = solPath;
                Log("VM solution loaded successfully.");
                return true;
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
                Log($"Failed to load solution: {_lastError}");
                return false;
            }
        }

        static string HandleMatchContour(string paramStr)
        {
            // Format: MATCH_CONTOUR:imagePath[,solPath]
            string[] p = paramStr.Split(',');
            if (p.Length < 1 || string.IsNullOrWhiteSpace(p[0]))
                return "ERROR:FORMAT — MATCH_CONTOUR:imagePath[,solPath]";

            string imagePath = p[0].Trim();
            string solPath = (p.Length >= 2 && !string.IsNullOrWhiteSpace(p[1]))
                ? p[1].Trim()
                : DefaultPartLocateSol;

            if (!File.Exists(imagePath))
                return $"ERROR:Image file not found: {imagePath}";

            if (!EnsureSolutionLoaded(solPath))
                return $"ERROR:Cannot load solution: {_lastError}";

            try
            {
                // Get ImageSource module and set the image
                var imgSrc = FindModule<ImageSourceModuleTool>("ImageSource1", "图像源1", "Image Source1");
                if (imgSrc == null)
                    return "ERROR:ImageSource module not found in part_locate.sol";

                SetImageSourcePath(imgSrc, imagePath);

                // Run the procedure
                var proc = FindProcedure("流程1", "Procedure1", "Process1", "Flow1");
                if (proc != null)
                {
                    proc.Run();
                }
                else
                {
                    imgSrc.Run();
                    var contourTool = FindModule<IMVSContourMatchModuTool>(
                        "ContourMatch1", "轮廓匹配1", "Contour Match1");
                    if (contourTool == null)
                        return "ERROR:ContourMatch module not found — run via procedure instead";
                    contourTool.Run();
                }

                // Get ContourMatch module results
                var matchTool = FindModule<IMVSContourMatchModuTool>(
                    "ContourMatch1", "轮廓匹配1", "Contour Match1");
                if (matchTool == null)
                    return "ERROR:ContourMatch module not found in part_locate.sol";

                var result = matchTool.ModuResult;
                int matchNum = result.MatchNum;
                Log($"ContourMatch: {matchNum} match(es) found");

                if (matchNum == 0)
                    return "OK|0";

                var matchPoints = result.MatchPoint;
                var matchScores = result.MatchScore;
                var matchRects  = result.MatchRect;

                var sb = new StringBuilder($"OK|{matchNum}");
                for (int i = 0; i < matchNum; i++)
                {
                    float px = (i < matchPoints.Count) ? matchPoints[i].X : 0;
                    float py = (i < matchPoints.Count) ? matchPoints[i].Y : 0;
                    float score = (i < matchScores.Count) ? matchScores[i] : 0;
                    float angle = (i < matchRects.Count) ? matchRects[i].Angle : 0;
                    sb.Append($"|{px:F2},{py:F2},{angle:F3},{score:F4}");
                    Log($"  Match {i}: px=({px:F1},{py:F1}) angle={angle:F2}° score={score:F4}");
                }
                return sb.ToString();
            }
            catch (VmException ex)
            {
                Log($"VM error: 0x{ex.errorCode:X} {ex.Message}");
                return $"ERROR:VM exception 0x{ex.errorCode:X}: {ex.Message}";
            }
            catch (Exception ex)
            {
                Log($"ContourMatch error: {ex.Message}");
                if (ex.InnerException != null)
                    Log($"  Inner: {ex.InnerException.Message}");
                return $"ERROR:{ex.Message}";
            }
        }

        // ──────────────────────────────────────────────────────────────────
        //  MATCH_QUICK — "快速匹配" (Fast Feature Pat Match), runtime-trained.
        //
        //  Wire-protocol (UNCHANGED from previous revision so VisionBridge
        //  needs no client-side update):
        //    MATCH_QUICK:imagePath,roiCx,roiCy,roiW,roiH,roiAngle[,solPath]
        //
        //  - imagePath        : absolute path of the image to search. The
        //                       WHOLE image is searched — there is no search-
        //                       region constraint on the matcher.
        //  - roiCx..roiAngle  : rectangle in image-pixel coordinates that
        //                       defines the TEMPLATE region (re-trained on
        //                       every call). Center-x, center-y, width,
        //                       height, angle(deg). If absent or zero-area
        //                       we return ERROR.
        //  - solPath          : IGNORED. Left in the wire format so old
        //                       VisionBridge builds still work, but the new
        //                       implementation no longer needs a .sol.
        //
        //  Response format (identical to before):
        //    OK|N|x,y,angle,score,boxW,boxH|x,y,angle,score,boxW,boxH|...
        //
        //  Why this rewrite (and why it's still "VM" / 快速匹配):
        //
        //    The previous implementation drove IMVSFastFeatureMatchModuTool
        //    (the platform-level SDK wrapper) which loads a baked-in template
        //    from Part Matching.sol. The platform SDK has NO call that
        //    accepts (image, ROI) → trained template at runtime — only
        //    ImportModelData(string[] paths) which takes pre-trained model
        //    files produced by the IDE.
        //
        //    HikRobot ships a SECOND SDK alongside Vision Master, the
        //    "MVDAlgorithmSDK" (env var MVDALGO_DEV_ENV, installed under
        //    C:\Program Files (x86)\MVDAlgorithmSDK). Same dongle, same
        //    licence, same algorithm DLLs (MVDFastFeaturePatMatchCpp.dll is
        //    bit-for-bit shared between the two SDKs, version 4.2.1.4), but
        //    a different surface: CFastFeaturePattern.Train() takes an image
        //    + region list and produces a trained pattern in-memory; the
        //    matcher tool takes that pattern and runs it on any image, with
        //    or without an external search-region. This is the SAME
        //    "快速匹配" engine the VM IDE drives when you click "Train" in
        //    its template editor — just exposed at a lower level. It is NOT
        //    OpenCV.
        //
        //  Implementation notes:
        //
        //  • We always full-canvas search (tool.ROI = null). The user's ROI
        //    is exclusively the template region.
        //  • Pattern + tool live for one request only. Train cost on a
        //    ~7785×7786 image with a 200×270 ROI is on the order of 50 ms
        //    (the engine extracts features only inside RegionList); the
        //    cost is dominated by the match Run(), not the train.
        //  • CFastFeaturePattern.MaskImage / RegionList retain native
        //    buffers, so we wrap both pattern & tool in `using` to make sure
        //    Dispose() runs even on exception. CMvdImage is also IDisposable.
        //  • Match limit defaults are MaxMatchNum=200 and MinScore=0.4. The
        //    matcher's own default would be MaxMatchNum=1 (!), which silently
        //    drops every match beyond the strongest — bad for arrays of
        //    flexes. We could expose these on the wire later if needed.
        // Wrap SetRunParam so an unknown / renamed key on the algorithm SDK
        // build degrades gracefully — we log it and move on instead of
        // tearing down the whole MATCH_QUICK call. The platform-SDK schema
        // and the algorithm-SDK runtime SHOULD share key names but we've
        // seen them diverge across point releases (e.g. "MaxNum" vs
        // "MaxMatchNum" in the 4.2.x line).
        static void TrySetParam(CFastFeaturePatMatchTool tool, string key, string value)
        {
            try
            {
                tool.SetRunParam(key, value);
            }
            catch (Exception ex)
            {
                Log($"[Quick] SetRunParam({key}={value}) failed (non-fatal): {ex.Message}");
            }
        }

        // Holds the 4 optional CFastFeaturePattern training params that the
        // VisionBridge "Preview…" popup tunes. ScaleMode/ThresholdMode are
        // bool-as-int (0=Auto, 1=Manual) to match the on-wire format; when a
        // mode is Auto the corresponding value is ignored by the SDK.
        struct TrainParams
        {
            public bool  ManualScale;     // true → use FeatureScale, false → auto
            public float FeatureScale;    // PyramidScaleRough  (typical 0.5–10)
            public bool  ManualThreshold; // true → use ContrastThreshold, false → auto
            public int   ContrastThresh;  // EdgeThreshold       (typical 5–200)
            public bool  Any => ManualScale || ManualThreshold;  // anything to push?
        }

        // Parse positions 7-10 of a comma-split MATCH_QUICK / MATCH_PREVIEW
        // payload as the optional train params. Tolerant: any missing or
        // unparseable field falls back to "all Auto" (the helper's previous
        // behaviour before the popup existed). The legacy solPath slot at
        // position 7 (a non-numeric string) is treated the same as "absent".
        static TrainParams ParseTrainParams(string[] p, int firstIdx)
        {
            var tp = new TrainParams();
            if (p == null || p.Length < firstIdx + 4) return tp;
            var inv = System.Globalization.CultureInfo.InvariantCulture;
            var ns  = System.Globalization.NumberStyles.Float;
            if (!int.TryParse(p[firstIdx].Trim(),     out int   sFlag)) return tp;   // legacy string → bail
            if (!float.TryParse(p[firstIdx + 1].Trim(), ns, inv, out float sRough)) return tp;
            if (!int.TryParse(p[firstIdx + 2].Trim(),   out int   tFlag)) return tp;
            if (!int.TryParse(p[firstIdx + 3].Trim(),   out int   tVal))  return tp;
            tp.ManualScale     = sFlag != 0;
            tp.FeatureScale    = sRough;
            tp.ManualThreshold = tFlag != 0;
            tp.ContrastThresh  = tVal;
            return tp;
        }

        // Apply the popup-tunable params to the trainer BEFORE Train(). Keys
        // ("PyramidScaleFlag", "PyramidScaleRough", "EdgeThresholdFlag",
        // "EdgeThreshold") are lifted straight from the FastFeaturePatMatch
        // C# sample (FastFeaturePatMatchTrainForm.cs) — only push the
        // Manual values when the corresponding mode is Manual, otherwise let
        // the engine compute them.
        static void ApplyTrainParams(CFastFeaturePattern pattern, TrainParams tp)
        {
            try
            {
                pattern.SetRunParam("PyramidScaleFlag",    tp.ManualScale     ? "1" : "0");
                pattern.SetRunParam("EdgeThresholdFlag",   tp.ManualThreshold ? "1" : "0");
                if (tp.ManualScale)
                    pattern.SetRunParam("PyramidScaleRough", tp.FeatureScale.ToString("0.00",
                        System.Globalization.CultureInfo.InvariantCulture));
                if (tp.ManualThreshold)
                    pattern.SetRunParam("EdgeThreshold",     tp.ContrastThresh.ToString(
                        System.Globalization.CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                Log($"[Quick] ApplyTrainParams non-fatal: {ex.Message}");
            }
        }

        // Public entry-point for MATCH_QUICK (returns hits only — wire-
        // identical to the pre-2026-05-24 response so the existing
        // VisionBridge parser is unchanged).
        static string HandleMatchQuick(string paramStr)
            => HandleMatchCore(paramStr, includeOutline: false, callerTag: "Quick");

        // Public entry-point for MATCH_PREVIEW — same train+match pipeline,
        // response prefixed with TRAIN / OUTLINE sections so the popup can
        // render the trained feature points.
        static string HandleMatchPreview(string paramStr)
            => HandleMatchCore(paramStr, includeOutline: true,  callerTag: "Preview");

        // Shared implementation. The only differences between QUICK and
        // PREVIEW are (a) the response shape and (b) whether we read the
        // trained OutlineList back from the pattern. Wire spec for both:
        //
        //   MATCH_*:imagePath,cx,cy,w,h,angle[,scaleFlag,scaleRough,thrFlag,thrVal]
        //
        // The 4 trailing fields are optional; missing/non-numeric → all
        // Auto, matching pre-popup behaviour. A legacy solPath string at
        // position 7 (from an older client) lands in ParseTrainParams as
        // "non-numeric → bail" and is silently ignored.
        //
        // Response shape:
        //   QUICK   → OK|N|cx,cy,a,s,bw,bh|cx,cy,a,s,bw,bh|...
        //   PREVIEW → OK|TRAIN|tplW,tplH,rough,low,high|OUTLINE|N|x1,y1,x2,y2,...|MATCH|M|cx,cy,a,s,bw,bh|...
        //
        // OUTLINE points are in TEMPLATE-LOCAL pixel coords (0..W × 0..H of
        // the cropped ROI), which is exactly the coordinate system the
        // popup's image is shown in. Multiple `CPatMatchOutline` entries
        // (if the engine ever returns >1) are concatenated.
        static string HandleMatchCore(string paramStr, bool includeOutline, string callerTag)
        {
            string[] p = paramStr.Split(',');
            if (p.Length < 1 || string.IsNullOrWhiteSpace(p[0]))
                return $"ERROR:FORMAT — MATCH_{callerTag.ToUpperInvariant()}:imagePath,cx,cy,w,h,angle" +
                       "[,scaleFlag,scaleRough,thrFlag,thrVal]";

            string imagePath = p[0].Trim();

            float roiCx = 0, roiCy = 0, roiW = 0, roiH = 0, roiAngle = 0;
            var inv = System.Globalization.CultureInfo.InvariantCulture;
            var ns  = System.Globalization.NumberStyles.Float;
            bool hasRoi = p.Length >= 6 &&
                float.TryParse(p[1].Trim(), ns, inv, out roiCx) &&
                float.TryParse(p[2].Trim(), ns, inv, out roiCy) &&
                float.TryParse(p[3].Trim(), ns, inv, out roiW)  &&
                float.TryParse(p[4].Trim(), ns, inv, out roiH)  &&
                float.TryParse(p[5].Trim(), ns, inv, out roiAngle);

            if (!hasRoi)
                return $"ERROR:MATCH requires a template ROI (cx,cy,w,h,angle). " +
                       "The user-drawn ROI is the template, re-trained per call.";
            if (roiW < 1 || roiH < 1)
                return $"ERROR:Template ROI too small ({roiW:F1}×{roiH:F1}). " +
                       "Draw a rectangle around at least one full feature before clicking Run.";
            if (!File.Exists(imagePath))
                return $"ERROR:Image file not found: {imagePath}";

            // Optional train params start at index 6.
            TrainParams tp = ParseTrainParams(p, 6);

            // Optional "skipMatch" flag at index 10 — used by the Preview
            // popup during live slider drag so the outline can refresh
            // fast without paying the full match-across-whole-canvas cost
            // on every tick. Defaults to 0 (run match) for backward compat.
            bool skipMatch = false;
            if (p.Length >= 11 && int.TryParse(p[10].Trim(), out int sm))
                skipMatch = sm != 0;

            CMvdImage img = null;
            CFastFeaturePattern pattern = null;
            CFastFeaturePatMatchTool tool = null;
            try
            {
                Log($"[{callerTag}] image={Path.GetFileName(imagePath)} " +
                    $"template_roi=cx={roiCx:F1},cy={roiCy:F1},w={roiW:F1},h={roiH:F1},ang={roiAngle:F2} " +
                    $"params=" + (tp.Any
                        ? $"manualScale={tp.ManualScale}({tp.FeatureScale:F2}) manualThr={tp.ManualThreshold}({tp.ContrastThresh})"
                        : "all-auto"));

                img = new CMvdImage();
                img.InitImage(imagePath);
                Log($"[{callerTag}] Image loaded: {img.Width}x{img.Height} fmt={img.PixelFormat}");

                if (img.PixelFormat != MVD_PIXEL_FORMAT.MVD_PIXEL_MONO_08)
                {
                    Log($"[{callerTag}] Converting from {img.PixelFormat} to MVD_PIXEL_MONO_08");
                    img.ConvertImagePixelFormat(MVD_PIXEL_FORMAT.MVD_PIXEL_MONO_08);
                }

                // ── TRAIN ─────────────────────────────────────────────────
                pattern = new CFastFeaturePattern();
                pattern.InputImage = img;

                var templateRect = new CMvdRectangleF(roiCx, roiCy, roiW, roiH);
                templateRect.Angle = roiAngle;
                pattern.RegionList.Add(new CPatMatchRegion
                {
                    Shape = templateRect,
                    Sign  = true,
                });

                ApplyTrainParams(pattern, tp);

                var trainSw = System.Diagnostics.Stopwatch.StartNew();
                pattern.Train();
                trainSw.Stop();
                var data = pattern.Result?.Data;
                int tplW = data?.Size.nWidth  ?? 0;
                int tplH = data?.Size.nHeight ?? 0;
                float rough = data?.Rough ?? 0f;
                int lowThr  = data?.LowThreshold  ?? 0;
                int highThr = data?.HighThreshold ?? 0;
                Log($"[{callerTag}] Train OK in {trainSw.ElapsedMilliseconds} ms — " +
                    $"template_size={tplW}x{tplH} rough={rough:F2} thr={lowThr}-{highThr}");

                // ── OUTLINE (preview only) ────────────────────────────────
                // Snapshot the trained edge-point list BEFORE matching, since
                // pattern.Run() doesn't touch the pattern's own outline but
                // we'd like the snapshot deterministic regardless.
                string outlineSegment = "";
                if (includeOutline)
                {
                    var outlines = pattern.Result?.OutlineList;
                    var ptsCsv = new StringBuilder();
                    int totalPts = 0;
                    if (outlines != null)
                    {
                        foreach (var ol in outlines)
                        {
                            var pts = ol?.EdgePointList;
                            if (pts == null) continue;
                            for (int j = 0; j < pts.Count; j++)
                            {
                                // CPatMatchEdgePoint exposes a Position
                                // (VisionDesigner.MVD_POINT_F — value type
                                // with float fields fX/fY, NOT properties
                                // X/Y) plus per-point Score and Weight. For
                                // the popup overlay we only need (x, y);
                                // score/weight could drive per-point alpha
                                // in a v2.
                                var pos = pts[j].Position;
                                if (totalPts > 0) ptsCsv.Append(',');
                                ptsCsv.AppendFormat(inv, "{0:F1},{1:F1}", pos.fX, pos.fY);
                                totalPts++;
                            }
                        }
                    }
                    outlineSegment = $"|OUTLINE|{totalPts}|{ptsCsv}";
                    Log($"[{callerTag}] Outline points: {totalPts}");
                }

                // ── MATCH (optional — skipped for live slider drag) ───────
                // skipMatch=true is used by the Preview popup's fast
                // outline-refresh path. The matched-hits list is empty in
                // that branch so downstream parsers see "M=0" and treat it
                // as a no-match-yet preview.
                // Concrete type name is CFastFeatureMatchInfo (no "Pat" in
                // the middle) — verified against MVDFastFeaturePatMatch.Net.xml.
                System.Collections.Generic.IList<CFastFeatureMatchInfo> matches = null;
                int matchNum = 0;
                if (!skipMatch)
                {
                    tool = new CFastFeaturePatMatchTool();
                    tool.InputImage = img;
                    tool.Pattern    = pattern;
                    tool.ROI        = null;  // full-canvas search

                    // Performance tuning — these are the levers that brought a 13+s
                    // full-canvas search on a 7800×7800 stitched image down to
                    // a couple of seconds:
                    //
                    //   AngleStart/AngleEnd (±15°) — the engine's default of ±180°
                    //     forces it to test ~360 orientations per candidate; almost
                    //     all real PCB / tray scenes are within a few degrees of
                    //     the template. ±15° is wide enough to cover hand-placed
                    //     drift but kills the cost of rotation invariance.
                    //   MaxMatchNum (64) — VM keeps a top-N list of partials during
                    //     pyramid refinement; the old 200 meant we paid for 200
                    //     refinements even on scenes with 30-40 real instances.
                    //   MinScore (0.5) — anything below 0.5 is noise for flex-pad
                    //     matching; raising the cutoff prunes weak candidates
                    //     before the slow per-candidate refinement step.
                    //   AccuracyType (0 = fast) — VM offers an "accurate" mode
                    //     that adds sub-pixel refinement; the default centroid
                    //     accuracy is plenty for dispense targeting.
                    TrySetParam(tool, "MaxMatchNum",  "64");
                    TrySetParam(tool, "MinScore",     "0.5");
                    TrySetParam(tool, "AngleStart",   "-15");
                    TrySetParam(tool, "AngleEnd",     "15");
                    TrySetParam(tool, "AccuracyType", "0");
                    // Some 4.2.x builds use these key names instead; harmless
                    // duplicates because TrySetParam silently swallows unknowns.
                    TrySetParam(tool, "AngleRangeStart", "-15");
                    TrySetParam(tool, "AngleRangeEnd",   "15");

                    var runSw = System.Diagnostics.Stopwatch.StartNew();
                    tool.Run();
                    runSw.Stop();

                    matches = tool.Result?.MatchInfoList;
                    matchNum = matches?.Count ?? 0;
                    Log($"[{callerTag}] Match Run in {runSw.ElapsedMilliseconds} ms — found {matchNum} hit(s).");
                }
                else
                {
                    Log($"[{callerTag}] skipMatch=1 — outline-only refresh, no canvas search.");
                }

                // ── Build response. QUICK keeps the legacy unprefixed layout
                // (OK|N|...) so existing parsers don't break; PREVIEW uses
                // tagged sections (OK|TRAIN|...|OUTLINE|...|MATCH|N|...). ─
                StringBuilder sb;
                if (includeOutline)
                {
                    sb = new StringBuilder("OK|TRAIN|");
                    sb.AppendFormat(inv, "{0},{1},{2:F2},{3},{4}", tplW, tplH, rough, lowThr, highThr);
                    sb.Append(outlineSegment);
                    sb.Append($"|MATCH|{matchNum}");
                }
                else
                {
                    sb = new StringBuilder($"OK|{matchNum}");
                }

                for (int i = 0; i < matchNum; i++)
                {
                    var m  = matches[i];
                    var bx = m.MatchBox;
                    sb.AppendFormat(inv, "|{0:F2},{1:F2},{2:F3},{3:F4},{4:F2},{5:F2}",
                        bx.CenterX, bx.CenterY, bx.Angle, m.Score, bx.Width, bx.Height);
                }
                return sb.ToString();
            }
            catch (MvdException ex)
            {
                Log($"[{callerTag}] MVD error 0x{ex.ErrorCode:X}: {ex.Message}");
                return $"ERROR:MVD 0x{ex.ErrorCode:X}: {ex.Message}";
            }
            catch (DllNotFoundException ex)
            {
                Log($"[{callerTag}] Native DLL missing: {ex.Message}");
                return "ERROR:MVDAlgorithmSDK native DLL not found (expected under " +
                       MvdAlgoRuntimeDir + "). " + ex.Message;
            }
            catch (Exception ex)
            {
                Log($"[{callerTag}] {ex.GetType().Name}: {ex.Message}");
                DumpExceptionChain(ex);
                return $"ERROR:{ex.Message}";
            }
            finally
            {
                try { tool?.Dispose(); }    catch { }
                try { pattern?.Dispose(); } catch { }
                try { img?.Dispose(); }     catch { }
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
