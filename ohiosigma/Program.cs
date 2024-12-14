using System.IO.Compression;
using ClosedXML.Excel;

string directory = @"Laporan";
string output = @"Output";

try
{
    if (! Directory.Exists(directory))
    {
        throw new Exception("Belum ada folder 'Laporan'");
    }

    if ( Directory.GetFiles(directory).Length == 0 )
    {
        throw new Exception("Folder 'Laporan' kosong");
    }    

    if (! Directory.Exists(output) )
    {
        throw new Exception("Belum ada folder 'Output' untuk output konversi format");
    }

    if ( Directory.GetFiles(output).Length != 0 )
    {
        throw new Exception("Folder 'Output' belum kosong. Kosongkan terlebih dahulu.");
    }
}
catch (Exception e)
{   
    Helpers.ShowErrorMessage(e);
    return;
}

String[] knownFile = [
    "adp01.xlsx", "ads01.xlsx", "pdp01.xlsx", "pds01.xlsx", "pkp01.xlsx", "pks01.xlsx", "rap01.xlsx", "ras01.xlsx", "scb01.xlsx"
];

String[] files = Directory.GetFiles(directory);

/** Cek apakah semua file valid di dalam folder */
foreach (String file in files)
{
    String check = file.Split('\\')[1];
    try
    {
        if (! knownFile.Contains<String>(check))
        {
            throw new Exception("Nama file tidak dikenal untuk "+check);
        }
    }
    catch (Exception e)
    {
        Helpers.ShowErrorMessage(e);
        return;
    }
}

/** Jika semua file valid di dalamm folder convert ke format txt dengan separator | 
    Tidak perlu mengecek kebenaran data karena dicek oleh Sistem
*/

List<FileContent> fileContents = [];

foreach (var file in files)
{
    try{
        XLWorkbook workbook = new XLWorkbook(file);
        IXLWorksheet ws = workbook.Worksheet(1);

        int firstRowNumber = (ws.FirstRowUsed()?.RowNumber()) ??
                             throw new Exception("File kosong untuk "+file.Split('\\')+". But why?");
        int rowNumber = (ws.LastRowUsed()?.RowNumber()) ??
                             throw new Exception("File kosong untuk "+file.Split('\\')+". But why?");

        int i = 1;
        if (ws.Row(firstRowNumber).Cell(1).Value.Equals("ID Pelapor") || ws.Row(firstRowNumber).Cell(1).Value.Equals("idPelapor"))
        {
            i++;

            if (ws.Row(i).Cell(1).IsEmpty())
            {
                throw new Exception("Data kosong untuk "+file+". What are you doing bruh");
            }
        }

        string text = "";
        string idPelapor = "";
        string status = "";
        string date = "";
        for (; i <= rowNumber; i++)
        {
            foreach (IXLCell cell in ws.Row(i).CellsUsed())
            {
                if (idPelapor.Equals("") && i > 1)
                {
                    idPelapor = ws.Row(i).Cell(1).Value.ToString();
                }

                if (status.Equals("") && i > 1)
                {
                    status = ws.Row(i).Cell(2).Value.ToString();

                    if ( !status.Equals("A") && !status.Equals("Q") && !status.Equals("I"))
                    {
                        throw new Exception("Periode Laporan salah pada Baris ke - "+i+" pada "+file);
                    }
                }

                if (date.Equals("") && i > 1)
                {
                    string parseMe = ws.Row(i).Cell(3).Value.ToString();
                    string[] resDate = parseMe.Split('-');

                    if (resDate.Length < 3 || resDate.Length > 3)
                    {
                        throw new Exception("Tanggal salah pada Baris ke - "+i+" pada "+file);
                    }

                    date = resDate[0]+resDate[1]+resDate[2];
                }

                text += cell.Value;

                if (cell.Address.ColumnNumber != ws.Row(i).LastCellUsed().Address.ColumnNumber)
                {
                    text += "|";
                }

            }
            text += Environment.NewLine;
        }

        // Write the string array to a new file named "WriteLines.txt".
        string filename = idPelapor+status+date+file.Split('\\')[1].Split('.')[0]+".txt";

        fileContents.Add(new FileContent(contents: text, original:file.Split('\\')[1], filename: filename));

        Console.WriteLine("Finished converting content of "+file.Split('\\')[1]);
    }
    catch(Exception e){
        Helpers.ShowErrorMessage(e);
        return;
    }
}

Console.WriteLine("Now making new file format");

string folderText = Path.Combine(output, "txt");
if (!Directory.Exists(folderText))
{
    Directory.CreateDirectory(folderText);
}

foreach (FileContent item in fileContents)
{
    string text = item.contents;
    string filename = item.filename;

    File.WriteAllText(Path.Combine(folderText, filename), text);

    Console.WriteLine("Finished converting "+filename);
}

Console.WriteLine("Package the files (Zipping)");
string folderZip = Path.Combine(output, "zip");

string zipName = Path.Combine(folderZip, "Laporan.zip");
if (!Directory.Exists(folderZip))
{
    Directory.CreateDirectory(folderZip);
}

File.Delete(zipName);

ZipFile.CreateFromDirectory(folderText, zipName);

Console.WriteLine("Finished creating package with name 'Laporan.zip'");
Console.WriteLine();
Console.WriteLine("Selesai");
Console.WriteLine("Press any key to exit");
Console.ReadKey();