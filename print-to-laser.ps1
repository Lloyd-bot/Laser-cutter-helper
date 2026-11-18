# ---------------------------------------------------------------------------
#  Laser Prep - print bridge
#  Sends a whole job from the tool straight to the laser, as real cut lines.
# ---------------------------------------------------------------------------
#  Why this exists
#
#  The Universal driver decides cut-versus-engrave by looking at the PEN each
#  line was drawn with. A GDI pen is either GEOMETRIC (a real width) or
#  COSMETIC (no width, always one device dot). Only cosmetic gets cut.
#
#  Chrome rasterises the page before printing, so no pens reach the driver at
#  all and everything engraves. Illustrator works because it emits cosmetic
#  pens. This script emits them too, talking to gdi32 directly - the same
#  approach Inkscape had to take for its Win32 Vector Print extension.
#
#  So: the tool saves a .lpj file (plain text, just the lines in millimetres),
#  this script draws it into a print job. Illustrator is no longer in the way.
#
#  Usage, from a PowerShell window in this folder:
#
#     powershell -ExecutionPolicy Bypass -File .\print-to-laser.ps1
#
#  With no -Job it takes the newest .lpj from this folder, then Downloads.
#  With no -Printer it looks for a Universal/VLS/ULS queue and, failing that,
#  lists what it can see and stops without printing anything.
#
#  Useful switches:
#     -DryRun                 read the job and report, print nothing
#     -TestSquare             ignore the job file, send one 100 mm square
#     -Yes                    skip the confirmation prompt
#     -Job <file>             a particular .lpj
#     -Printer "<name>"       a particular queue
#     -Out <file>             for file-based printers, e.g. Print to PDF
#
#  To check the output without touching the laser:
#     .\print-to-laser.ps1 -Printer "Microsoft Print to PDF" -Out "$env:TEMP\t.pdf" -Yes
# ---------------------------------------------------------------------------
param(
  [string]$Job,
  [string]$Printer,
  [string]$Out,
  [switch]$TestSquare,
  [double]$SizeMm = 100,
  [double]$LeftMm = 10,
  [double]$TopMm  = 10,
  [switch]$DryRun,
  [switch]$Yes
)
$ErrorActionPreference = 'Stop'

Add-Type @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

public class LaserGdi {

  // ---- one drawable thing: a run of lines, or a filled region -------------
  public class Shape {
    public bool Fill;          // true = filled area (engraves), false = stroked line (cuts)
    public int  Color;         // COLORREF, 0x00BBGGRR
    public bool EvenOdd;
    public List<double[]> Polys = new List<double[]>();   // x0,y0,x1,y1,... in mm
  }

  public static List<Shape> Shapes = new List<Shape>();
  public static double PageW, PageH, Cal = 1;
  public static string JobName = "", Code = "", Material = "";
  public static int Points, Strokes, Fills;
  // Where the job really reaches, measured from the points rather than taken
  // on trust from the PAGE line. A size correction scales position as well as
  // size, so a job can genuinely sit outside the page it declares.
  public static double MinX, MinY, MaxX, MaxY;

  // ---- read a .lpj --------------------------------------------------------
  // Deliberately dull line-based text: one SHAPE line, then one P line per
  // loop, points as x,y in millimetres from the top-left of the page. If this
  // ever has to be debugged at the laser PC it can be read in Notepad.
  public static void Parse(string path){
    Shapes = new List<Shape>(); Points = 0; Strokes = 0; Fills = 0;
    Cal = 1; PageW = 0; PageH = 0; JobName = ""; Code = ""; Material = "";
    MinX = double.MaxValue; MinY = double.MaxValue;
    MaxX = double.MinValue; MaxY = double.MinValue;
    CultureInfo inv = CultureInfo.InvariantCulture;
    Shape cur = null;
    int lineNo = 0;
    bool header = false;
    foreach(string raw in File.ReadLines(path)){
      lineNo++;
      string line = raw.Trim();
      if(line.Length == 0) continue;
      int sp = line.IndexOf(' ');
      string key = sp < 0 ? line : line.Substring(0, sp);
      string rest = sp < 0 ? "" : line.Substring(sp + 1);
      switch(key){
        case "LPJ":
          header = true;
          if(rest.Trim() != "1") throw new Exception("This job file is version " + rest.Trim()
            + ", and this script only knows version 1. Update one of them.");
          break;
        case "NAME":     JobName  = rest; break;
        case "CODE":     Code     = rest; break;
        case "MATERIAL": Material = rest; break;
        case "CAL":      Cal = double.Parse(rest, inv); break;
        case "PAGE": {
          string[] a = rest.Split(' ');
          PageW = double.Parse(a[0], inv); PageH = double.Parse(a[1], inv);
          break;
        }
        case "SHAPE": {
          string[] a = rest.Split(' ');
          cur = new Shape();
          cur.Fill    = (a[0] == "fill");
          cur.Color   = ColorRef(a[1]);
          cur.EvenOdd = (a.Length > 2 && a[2] == "evenodd");
          Shapes.Add(cur);
          if(cur.Fill) Fills++; else Strokes++;
          break;
        }
        case "P": {
          if(cur == null) throw new Exception("Line " + lineNo + ": points before any SHAPE line.");
          string[] t = rest.Split(' ');
          double[] pts = new double[t.Length * 2];
          int n = 0;
          for(int i = 0; i < t.Length; i++){
            int c = t[i].IndexOf(',');
            if(c < 0) continue;
            pts[n++] = double.Parse(t[i].Substring(0, c), inv);
            pts[n++] = double.Parse(t[i].Substring(c + 1), inv);
          }
          if(n < 4) break;                       // a single point draws nothing
          if(n != pts.Length){ Array.Resize(ref pts, n); }
          for(int i = 0; i < n; i += 2){
            if(pts[i]   < MinX) MinX = pts[i];
            if(pts[i]   > MaxX) MaxX = pts[i];
            if(pts[i+1] < MinY) MinY = pts[i+1];
            if(pts[i+1] > MaxY) MaxY = pts[i+1];
          }
          cur.Polys.Add(pts);
          Points += n / 2;
          break;
        }
        case "END": break;
        default: break;                          // unknown keys ignored, so the
      }                                          // format can grow without breaking
    }
    if(!header) throw new Exception("That file does not start with LPJ - it is not a Laser Prep job file.");
  }

  static int ColorRef(string rrggbb){
    string s = rrggbb.TrimStart('#');
    int r = Convert.ToInt32(s.Substring(0,2), 16);
    int g = Convert.ToInt32(s.Substring(2,2), 16);
    int b = Convert.ToInt32(s.Substring(4,2), 16);
    return (b << 16) | (g << 8) | r;             // COLORREF is 0x00BBGGRR
  }

  // One 100 mm square, standing in for a job. The yes/no test of the whole idea.
  public static void TestSquare(double size, double left, double top){
    Shapes = new List<Shape>(); Points = 5; Strokes = 1; Fills = 0;
    PageW = left + size + left; PageH = top + size + top; Cal = 1;
    JobName = size + " mm size check"; Code = ""; Material = "";
    Shape s = new Shape();
    s.Fill = false; s.Color = 0x000000FF;        // red
    s.Polys.Add(new double[]{ left,top, left+size,top, left+size,top+size, left,top+size, left,top });
    Shapes.Add(s);
    MinX = left; MinY = top; MaxX = left + size; MaxY = top + size;
  }

  // ---- gdi32 --------------------------------------------------------------
  [StructLayout(LayoutKind.Sequential)] public struct POINT { public int x; public int y; }
  [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
  public struct DOCINFO {
    public int cbSize;
    [MarshalAs(UnmanagedType.LPWStr)] public string lpszDocName;
    [MarshalAs(UnmanagedType.LPWStr)] public string lpszOutput;
    [MarshalAs(UnmanagedType.LPWStr)] public string lpszDatatype;
    public int fwType;
  }
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, EntryPoint="CreateDCW")]
  public static extern IntPtr CreateDC(string drv, string dev, string port, IntPtr dm);
  [DllImport("gdi32.dll")] public static extern bool DeleteDC(IntPtr hdc);
  [DllImport("gdi32.dll", CharSet=CharSet.Unicode, EntryPoint="StartDocW")]
  public static extern int StartDoc(IntPtr hdc, ref DOCINFO di);
  [DllImport("gdi32.dll")] public static extern int EndDoc(IntPtr hdc);
  [DllImport("gdi32.dll")] public static extern int AbortDoc(IntPtr hdc);
  [DllImport("gdi32.dll")] public static extern int StartPage(IntPtr hdc);
  [DllImport("gdi32.dll")] public static extern int EndPage(IntPtr hdc);
  [DllImport("gdi32.dll")] public static extern IntPtr CreatePen(int style, int width, int color);
  [DllImport("gdi32.dll")] public static extern IntPtr CreateSolidBrush(int color);
  [DllImport("gdi32.dll")] public static extern IntPtr GetStockObject(int i);
  [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr hdc, IntPtr o);
  [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr o);
  [DllImport("gdi32.dll")] public static extern bool Polyline(IntPtr hdc, POINT[] p, int n);
  [DllImport("gdi32.dll")] public static extern bool PolyPolygon(IntPtr hdc, POINT[] p, int[] c, int n);
  [DllImport("gdi32.dll")] public static extern int SetPolyFillMode(IntPtr hdc, int mode);
  [DllImport("gdi32.dll")] public static extern int GetDeviceCaps(IntPtr hdc, int i);

  const int LOGPIXELSX=88, LOGPIXELSY=90, HORZSIZE=4, VERTSIZE=6,
            HORZRES=8, VERTRES=10, PHYSOFFX=112, PHYSOFFY=113;
  const int NULL_PEN = 8, ALTERNATE = 1, WINDING = 2;
  const int CHUNK = 8000;     // some drivers choke on one enormous Polyline

  // What the driver says about itself. Read before printing so a wrong page
  // size in the driver shows up as numbers rather than as a ruined sheet.
  public static string Probe(string printer, out double sheetW, out double sheetH){
    IntPtr hdc = CreateDC("WINSPOOL", printer, null, IntPtr.Zero);
    if(hdc == IntPtr.Zero) throw new Exception("Could not open printer '" + printer + "'. Check the name is exact.");
    try {
      int dx = GetDeviceCaps(hdc, LOGPIXELSX), dy = GetDeviceCaps(hdc, LOGPIXELSY);
      sheetW = GetDeviceCaps(hdc, HORZSIZE); sheetH = GetDeviceCaps(hdc, VERTSIZE);
      double ox = GetDeviceCaps(hdc, PHYSOFFX) * 25.4 / Math.Max(1, dx);
      double oy = GetDeviceCaps(hdc, PHYSOFFY) * 25.4 / Math.Max(1, dy);
      StringBuilder sb = new StringBuilder();
      sb.Append("  resolution  : " + dx + " x " + dy + " dots per inch\n");
      sb.Append("  page in driver: " + sheetW + " x " + sheetH + " mm\n");
      sb.Append("  margin it will not print into: "
        + ox.ToString("0.0") + " x " + oy.ToString("0.0") + " mm"
        + (ox > 0.05 || oy > 0.05 ? "  (compensated for)" : "") + "\n");
      if(dx <= 0 || dy <= 0) throw new Exception("The driver reported no resolution.");
      return sb.ToString();
    } finally { DeleteDC(hdc); }
  }

  public static string Print(string printer, string outFile){
    IntPtr hdc = CreateDC("WINSPOOL", printer, null, IntPtr.Zero);
    if(hdc == IntPtr.Zero) throw new Exception("Could not open printer '" + printer + "'.");
    int dpiX = GetDeviceCaps(hdc, LOGPIXELSX), dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
    // GDI draws relative to the printable area, not the paper. Where a driver
    // keeps an unprintable margin, subtracting it puts the job back where the
    // tool said it was - at the top-left corner of the bed.
    int offX = GetDeviceCaps(hdc, PHYSOFFX), offY = GetDeviceCaps(hdc, PHYSOFFY);

    DOCINFO di = new DOCINFO();
    di.cbSize = Marshal.SizeOf(typeof(DOCINFO));
    di.lpszDocName = "Laser Prep" + (JobName.Length > 0 ? " - " + JobName : "")
                   + (Code.Length > 0 ? " [" + Code + "]" : "");
    if(!string.IsNullOrEmpty(outFile)) di.lpszOutput = outFile;

    if(StartDoc(hdc, ref di) <= 0){ DeleteDC(hdc); throw new Exception("StartDoc failed - the driver refused the job."); }
    bool started = false;
    int drawn = 0, failed = 0;
    try {
      StartPage(hdc); started = true;
      IntPtr nullPen = GetStockObject(NULL_PEN);
      foreach(Shape s in Shapes){
        if(s.Polys.Count == 0) continue;
        if(s.Fill){
          // Filled black areas are what the driver rasterises, i.e. engraves.
          // All the loops of one shape go in a single PolyPolygon so that holes
          // stay holes instead of filling in solid.
          IntPtr br = CreateSolidBrush(s.Color);
          IntPtr ob = SelectObject(hdc, br);
          IntPtr op = SelectObject(hdc, nullPen);
          SetPolyFillMode(hdc, s.EvenOdd ? ALTERNATE : WINDING);
          int total = 0;
          foreach(double[] p in s.Polys) total += p.Length / 2;
          POINT[] pts = new POINT[total];
          int[] counts = new int[s.Polys.Count];
          int k = 0, ci = 0;
          foreach(double[] p in s.Polys){
            counts[ci++] = p.Length / 2;
            for(int i = 0; i < p.Length; i += 2){
              pts[k].x = Dot(p[i], dpiX, offX); pts[k].y = Dot(p[i+1], dpiY, offY); k++;
            }
          }
          if(PolyPolygon(hdc, pts, counts, counts.Length)) drawn++; else failed++;
          SelectObject(hdc, op);
          SelectObject(hdc, ob);
          DeleteObject(br);
        } else {
          // PS_SOLID with width 0. The ZERO is the whole point: it makes the
          // pen cosmetic, and cosmetic is the only thing the driver cuts.
          IntPtr pen = CreatePen(0, 0, s.Color);
          if(pen == IntPtr.Zero) throw new Exception("CreatePen failed.");
          IntPtr old = SelectObject(hdc, pen);
          foreach(double[] p in s.Polys){
            int n = p.Length / 2;
            POINT[] pts = new POINT[n];
            for(int i = 0; i < n; i++){
              pts[i].x = Dot(p[2*i], dpiX, offX); pts[i].y = Dot(p[2*i+1], dpiY, offY);
            }
            int at = 0;
            while(at < n - 1){
              int take = Math.Min(CHUNK, n - at);
              POINT[] part = pts;
              if(at != 0 || take != n){ part = new POINT[take]; Array.Copy(pts, at, part, 0, take); }
              if(Polyline(hdc, part, take)) drawn++; else failed++;
              at += take - 1;                    // overlap one point, no gap in the line
            }
          }
          SelectObject(hdc, old);
          DeleteObject(pen);
        }
      }
      EndPage(hdc); started = false;
      EndDoc(hdc);
    } catch {
      if(started) EndPage(hdc);
      AbortDoc(hdc); DeleteDC(hdc); throw;
    }
    DeleteDC(hdc);
    return drawn + (drawn == 1 ? " drawing call" : " drawing calls") + " accepted"
         + (failed > 0 ? ", " + failed + " REFUSED by the driver" : ", none refused");
  }

  static int Dot(double mm, int dpi, int off){
    return (int)Math.Round(mm / 25.4 * dpi) - off;
  }
}
'@

# --- find the job -----------------------------------------------------------
if(-not $TestSquare){
  if(-not $Job){
    # Run as a file, that is the script's own folder. Handed to PowerShell as
    # text - pasted into a window, or run by the .cmd launcher, which is how a
    # PC that refuses to run .ps1 files gets to use this at all - there is no
    # file to ask about, and the folder the window is sitting in is the answer.
    $here = if($MyInvocation.MyCommand.Path){ Split-Path -Parent $MyInvocation.MyCommand.Path }
            else { (Get-Location).Path }
    $dirs = @($here, (Join-Path $env:USERPROFILE 'Downloads')) | Where-Object { Test-Path $_ }
    $found = $dirs | ForEach-Object { Get-ChildItem -Path $_ -Filter *.lpj -File -ErrorAction SilentlyContinue } |
             Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if(-not $found){
      Write-Host "No job file found. Nothing has been printed." -ForegroundColor Yellow
      Write-Host "`nIn Laser Prep, click 'Send to the laser' - it saves a .lpj file."
      Write-Host "Looked in:"
      $dirs | ForEach-Object { Write-Host "   $_" }
      Write-Host "`nOr run with -TestSquare to send a 100 mm square instead.`n"
      exit 1
    }
    $Job = $found.FullName
  }
  if(-not (Test-Path $Job)){
    Write-Host "No such job file: $Job" -ForegroundColor Yellow
    exit 1
  }
  $Job = (Resolve-Path $Job).Path
  # A half-written or wrong file should say so in one plain line, not in a
  # wall of .NET exception text at someone mid-lesson.
  try { [LaserGdi]::Parse($Job) }
  catch {
    $m = $_.Exception
    while($m.InnerException){ $m = $m.InnerException }
    Write-Host "`nThat job file could not be read. Nothing has been printed." -ForegroundColor Yellow
    Write-Host "   $Job"
    Write-Host "   $($m.Message)`n"
    exit 1
  }
} else {
  [LaserGdi]::TestSquare($SizeMm, $LeftMm, $TopMm)
}

# --- pick a printer ---------------------------------------------------------
if(-not $Printer){
  $all = @(Get-Printer | Select-Object -ExpandProperty Name)
  $guess = @($all | Where-Object { $_ -match 'VLS|Universal|ULS|Laser' })
  if($guess.Count -eq 1){
    $Printer = $guess[0]
  } else {
    Write-Host "Could not identify the laser automatically. Nothing has been printed." -ForegroundColor Yellow
    Write-Host "`nPrinters on this PC:`n"
    $all | ForEach-Object { Write-Host "   $_" }
    Write-Host "`nRe-run naming one, for example:"
    Write-Host "   .\print-to-laser.ps1 -Printer `"VLS4.60`"`n"
    exit 1
  }
}

# --- report, then ask -------------------------------------------------------
$sheetW = 0.0; $sheetH = 0.0
$info = [LaserGdi]::Probe($Printer, [ref]$sheetW, [ref]$sheetH)

Write-Host ""
Write-Host "Job     : $([LaserGdi]::JobName)$(if([LaserGdi]::Code){ '  [' + [LaserGdi]::Code + ']' })"
if(-not $TestSquare){ Write-Host "File    : $Job" }
if([LaserGdi]::Material){ Write-Host "Material: $([LaserGdi]::Material)" }
Write-Host ("Page    : {0} x {1} mm" -f [LaserGdi]::PageW, [LaserGdi]::PageH)
$plural = { param($n,$one,$many) "$n " + $(if($n -eq 1){ $one } else { $many }) }
Write-Host ("Content : {0}, {1}, {2} points" -f `
  (& $plural ([LaserGdi]::Strokes) 'cut/score line' 'cut/score lines'),
  (& $plural ([LaserGdi]::Fills) 'engraved area' 'engraved areas'),
  [LaserGdi]::Points)
if([LaserGdi]::Cal -ne 1){
  Write-Host ("Size correction of {0:0.00}% is already baked in - do not re-scale in the laser software." -f `
    (([LaserGdi]::Cal - 1) * 100)) -ForegroundColor Cyan
}
Write-Host ""
Write-Host "Printer : $Printer"
Write-Host $info -NoNewline

# The lines themselves decide whether this fits, not the page the file claims.
$x0 = [LaserGdi]::MinX; $y0 = [LaserGdi]::MinY
$x1 = [LaserGdi]::MaxX; $y1 = [LaserGdi]::MaxY
if([LaserGdi]::Points -gt 0){
  Write-Host ("Lies at : {0:0.0},{1:0.0} to {2:0.0},{3:0.0} mm from the top-left corner" -f $x0,$y0,$x1,$y1)
  if($x0 -lt -0.05 -or $y0 -lt -0.05){
    Write-Host "WARNING: part of this job sits off the top or left of the bed and will be lost." -ForegroundColor Yellow
  }
  if($sheetW -gt 0 -and ($x1 -gt $sheetW + 0.5 -or $y1 -gt $sheetH + 0.5)){
    Write-Host ("WARNING: this job reaches past the page set in the driver ({0} x {1} mm)." -f $sheetW, $sheetH) -ForegroundColor Yellow
    Write-Host "         Anything past the edge will be cut off. Set the page size in the" -ForegroundColor Yellow
    Write-Host "         printer's Properties to the bed, or move the design in Laser Prep." -ForegroundColor Yellow
  }
}
if([LaserGdi]::Strokes -eq 0 -and [LaserGdi]::Fills -gt 0){
  Write-Host "NOTE: this job is all engraving - there is nothing in it to cut." -ForegroundColor Yellow
}

if($DryRun){
  Write-Host "`n-DryRun: read and checked, nothing sent.`n" -ForegroundColor Green
  exit 0
}
if(-not $Yes){
  Write-Host ""
  $a = Read-Host "Send this to $Printer? [y/N]"
  if($a -notmatch '^(y|yes)$'){ Write-Host "Cancelled. Nothing sent.`n"; exit 0 }
}

# --- send -------------------------------------------------------------------
if($Out){
  New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
  if(Test-Path $Out){ Remove-Item $Out -Force }
}
$result = [LaserGdi]::Print($Printer, $Out)

Write-Host ""
Write-Host "Sent. $result" -ForegroundColor Green
if($Out){
  Start-Sleep -Milliseconds 1500
  if(Test-Path $Out){ Write-Host "Wrote $Out ($((Get-Item $Out).Length) bytes)" }
  else { Write-Host "No output file yet - the spooler may still be writing." -ForegroundColor Yellow }
} else {
  Write-Host "It should now be waiting in the Universal software as a job."
  Write-Host "Check there that the red lines show as vectors before you run it."
  if($TestSquare){ Write-Host ("Measure the square after cutting: it should be exactly {0} mm." -f $SizeMm) }
}
Write-Host ""
