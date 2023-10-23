using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

// --- setup and verification ---

if(args.Length < 2)
{
    Console.WriteLine("Args not specified");
    Environment.Exit(0);
}
// get directory of the files and subdirectories to read
string directory = args[0];

// verify the directory
if (!Directory.Exists(directory))
{
    Console.WriteLine("Specified directory not found");
    Environment.Exit(0);
}

// get and parse depth
int depth = 1;
try
{
    depth = int.Parse(args[1]);
}
catch(Exception e)
{
    Console.WriteLine("Depth specified incorrectly, reson: " + e.Message);
    Environment.Exit(0);
}

// specify target path of the document
string targetFile = @"..\..\..\lab2.xlsx";

// make space for the document
if(File.Exists(targetFile))
{
    File.Delete(targetFile);
}

string firstWorksheetName = "Struktura katalogu";
string secondWorksheetName = "Statystyki";


// --- utility functions ---

ExcelPieChart SetupChart(ExcelPieChart chart, int x, int y)
{
    chart.SetPosition(x, 0, y, 0);
    chart.SetSize(400, 300);
    chart.ShowDataLabelsOverMaximum = true;
    chart.ShowHiddenData = true;
    chart.DataLabel.ShowCategory = true;
    chart.DataLabel.ShowPercent = true;
    return chart;
}

string[] suffixes =
{ "Bytes", "KB", "MB", "GB" };

string FormatSize(long bytes)
{
    int counter = 0;
    decimal number = (decimal)bytes;
    while (Math.Round(number / 1024) >= 1)
    {
        number = number / 1024;
        counter++;
        if (counter == 3)
            break;
    }
    return string.Format("{0:n1}{1}", number, suffixes[counter]);
}

void fillFileInfo(FileInfo f, int cellIndex,
    int OutlineLevel, ExcelWorksheet ws, List<fileStruct> files = null)
{
    ws.Cells[cellIndex, 1].Value = f.FullName;
    ws.Cells[cellIndex, 2].Value = f.Extension;
    ws.Cells[cellIndex, 3].Value = FormatSize(f.Length);
    ws.Cells[cellIndex, 4].Value = f.Attributes;
    ws.Row(cellIndex).OutlineLevel = OutlineLevel;
    if(files!=null)
        files.Add(new fileStruct(f.FullName, f.Length));
}

void RecursiveSubfoldersAndFiles(string path, int depth,
    int cellIndex, int OutlineLevel, ExcelWorksheet ws, List<fileStruct> files = null)
{
    try
    {
        // iterate through directories
        foreach (string directory in Directory.EnumerateDirectories(path))
        {
            // save current directory
            ws.Cells[cellIndex, 1].Value = directory;
            ws.Row(cellIndex).OutlineLevel = OutlineLevel;
            cellIndex++;
            if (depth > 0)
            {
                // handle files in the directory
                try
                {
                    foreach (string file in Directory.EnumerateFiles(directory))
                    {
                        fillFileInfo(new FileInfo(file), cellIndex, OutlineLevel + 1, ws, files);
                        cellIndex++;
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    Console.WriteLine("Unauthorized access exception: " + ex.Message);
                }
                catch (DirectoryNotFoundException ex)
                {
                    Console.WriteLine("Files not found exception: " + ex.Message);
                }
                catch (IOException ex)
                {
                    Console.WriteLine("I/O exception: " + ex.Message);
                }

                // handle subdirectories (recursive)
                RecursiveSubfoldersAndFiles(directory, depth - 1, cellIndex, OutlineLevel + 1, ws, files);
            }
        }
    }
    catch (UnauthorizedAccessException ex)
    {
        Console.WriteLine("Unauthorized access exception: " + ex.Message);
    }
    catch (DirectoryNotFoundException ex)
    {
        Console.WriteLine("Directory not found exception: " + ex.Message);
    }
    catch (IOException ex)
    {
        Console.WriteLine("I/O exception: " + ex.Message);
    }
}

// --- "main" ---

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var ep = new ExcelPackage(new FileInfo(targetFile));

ExcelWorksheet ws = ep.Workbook.Worksheets.Add(firstWorksheetName);
ExcelWorksheet secondWs = ep.Workbook.Worksheets.Add(secondWorksheetName);

int cellIndex = 1, OutlineLevel = 1;
List<fileStruct> files = new List<fileStruct>();

// save root directory
ws.Cells[cellIndex, 1].Value = directory;
ws.Row(cellIndex).OutlineLevel = OutlineLevel;
cellIndex++;

// handle files in the directory
foreach (string file in Directory.EnumerateFiles(directory))
{
    fillFileInfo(new FileInfo(file), cellIndex, OutlineLevel, ws, files);
    cellIndex++;
}

// handle subdirectories recursivly
RecursiveSubfoldersAndFiles(directory, depth, cellIndex, OutlineLevel, ws, files);

// LINQ
// sort files by size and get 10 largest in the descending order
var largestFiles = files.OrderByDescending(i => i.size).Take(10);

// handle selected files and summarize extensions for the charts
Dictionary<string, int> numberOfExtensionsDict = new Dictionary<string, int>();
Dictionary<string, long> sizeOfExtensionsDict = new Dictionary<string, long>();

cellIndex = 1; OutlineLevel = 1;
foreach( fileStruct file in largestFiles )
{
    // handle files
    FileInfo f = new FileInfo(file.path);
    fillFileInfo(f, cellIndex, OutlineLevel, secondWs);
    // handle extensions
    if (numberOfExtensionsDict.ContainsKey(f.Extension))
    {
        numberOfExtensionsDict[f.Extension]++;
        sizeOfExtensionsDict[f.Extension] += f.Length;
    }
    else
    {
        sizeOfExtensionsDict.Add(f.Extension, f.Length);
        numberOfExtensionsDict.Add(f.Extension, 1);
    }
    cellIndex++;
}

// save extensions in excel cells
cellIndex = 1;
foreach(var extension in numberOfExtensionsDict.Keys)
{
    secondWs.Cells[cellIndex, 5].Value = extension;
    secondWs.Cells[cellIndex, 6].Value = numberOfExtensionsDict[extension];
    secondWs.Cells[cellIndex, 7].Value = sizeOfExtensionsDict[extension];
    cellIndex++;
}
cellIndex--;

// create charts
ExcelPieChart chart = secondWs.Drawings.AddPieChart("ExtensionsNumber",
ePieChartType.Pie3D);
chart.Title.Text = "Percentage of the extension";
chart = SetupChart(chart, 1, 8);

var ser1 = (chart.Series.Add(secondWs.Cells[1,6,cellIndex,6],
secondWs.Cells[1,5,cellIndex,5])) as ExcelPieChartSerie;
ser1.Header = "Amount";

ExcelPieChart chartSize = secondWs.Drawings.AddPieChart("ExtensionsSize",
ePieChartType.Pie3D);
chartSize.Title.Text = "Percentage of total size";
chartSize = SetupChart(chartSize, 20, 8);

var ser2 = (chartSize.Series.Add(secondWs.Cells[1, 7, cellIndex, 7],
secondWs.Cells[1, 5, cellIndex, 5])) as ExcelPieChartSerie;
ser2.Header = "Amount";


// adjust the column sizes
ws.Column(1).AutoFit();
secondWs.Column(1).AutoFit();

// save the file
ep.Save(@"");

Console.WriteLine("Finished");

public struct fileStruct
{
    public fileStruct(string file, long size)
    {
        this.path = file;
        this.size = size;
    }
    public string path { get; }
    public long size { get; }
}