using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

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
    Console.WriteLine(e.Message);
    return;
}

String[] knownFile = ["adp01.xlsx", "pks01.xlsx"];

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
        Console.WriteLine(e.Message);
        return;
    }
}

/** Jika semua file valid di dalamm folder convert ke format txt dengan separator | 
    Tidak perlu mengecek kebenaran data karena dicek oleh Sistem
*/
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
                throw new Exception("Data kosong. What are you doing bruh");
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

                    if ( !status.Equals("A") && !status.Equals("Q"))
                    {
                        throw new Exception("Periode Laporan salah pada Baris ke - "+i);
                    }
                }

                if (date.Equals("") && i > 1)
                {
                    string parseMe = ws.Row(i).Cell(3).Value.ToString();
                    string[] resDate = parseMe.Split('-');

                    if (resDate.Length < 3 || resDate.Length > 3)
                    {
                        throw new Exception("Tanggal salah pada Baris ke - "+i);
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
        File.WriteAllText(Path.Combine(output, filename), text);

        Console.WriteLine("Finished converting "+file.Split('\\')[1]+" to "+filename);
    }
    catch(Exception e){
        Console.WriteLine(e.Message);
        return;
    }
}