using System;
using System.IO;
using Windows.System.Profile;

public abstract class DirectoryManipulator
{
    public enum FileExtension
    {
        Excel,
        Log,
        Zip
    }

    public virtual string CreateDirectory(string path)
    {
        if (Directory.Exists(path))
        {
            return "";
        }
        Directory.CreateDirectory(path);
        return path;
    }

    public virtual string DeleteFilesBase(string path, string extension)
    {
        DirectoryInfo directory = new DirectoryInfo(path);
        foreach (FileInfo file in directory.GetFiles($"{extension}"))
        {
            file.Delete();
        }
        foreach (DirectoryInfo subDirectory in directory.GetDirectories())
        {
            DeleteFilesBase(subDirectory.FullName, extension);
        }
        return "";
    }


    public virtual string ZipDirectory(string sourcePath, string destinationPath)
    {

        return "";
    }

    public virtual string SendFileToUrl(string filePath, string url)
    {

        return "";       
    }


}

public class  MyDirectoryManipulator : DirectoryManipulator
{
   

    public override string CreateDirectory(string path)
    {
        var value = base.CreateDirectory(path)  == "" ? "" : $"Creating directory at {path}";
        return value;

    }

    public string DeleteFiles(string path, FileExtension fileExtension)
    {
        string extension = string.Empty;

        switch (fileExtension)
        {
            case FileExtension.Excel:
                extension = "*.xl*";
                break;
            case FileExtension.Log:
                extension = "*.log";
                break;
            case FileExtension.Zip:
                extension = "*.zip";
                break;
        }
        base.DeleteFilesBase(path, extension);
        return ($"Deleting files with extension {extension} in {path}");
    }

    public override string ZipDirectory(string sourcePath, string destinationPath)
    {
        base.ZipDirectory(sourcePath, destinationPath);
        return ($"Zipping directory at {sourcePath} to {destinationPath}");
    }

    public override string SendFileToUrl(string filePath, string url)
    {
        base.SendFileToUrl(filePath, url);
        return ($"Sending file at {filePath} to {url}");

    }


}

public static class DateManipul
{
    public static string GetPrevMonth()
    {
        return DateTime.Now.AddMonths(-1).ToString("MM");
    }

    public static string GetPrevYear()
    {
        return DateTime.Now.AddMonths(-1).ToString("yyyy");
    }

    public static string GetFirstDate()
    {
        return "01";
    }

    public static string GetLastDayOfPrevMonth()
    {
        var lastDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddDays(-1);
        return lastDay.ToString("dd");
    }

    public static string GetDateFrom()
    {
        return $@"{GetFirstDate}/{GetPrevMonth}/{GetPrevYear} 00:00";
    }

    public static string GetDateTo()
    {
        return $@"{GetLastDayOfPrevMonth}/{GetPrevMonth}/{GetPrevYear} 00:00";
    }

    public static string GetDSPeriod()
    {
        return $@"{GetPrevYear}{GetPrevMonth}";
    }

   
}

public class Datemanipulation
{
    private readonly string _getPrevMonth;
    private readonly string _getPrevYear;
    private readonly string _getFirstDateofPrevMonth;
    private readonly string _getLastDateofPrevMonth;
    private readonly string _getDateFrom;
    private readonly string _getDateTo;
    private readonly string _getDSPeriod;

    public string GetPrevMonth => _getPrevMonth;

    public string GetPrevYear => _getPrevYear;

    public string GetFirstDateofPrevMonth => _getFirstDateofPrevMonth;

    public string GetLastDateofPrevMonth => _getLastDateofPrevMonth;

    public string GetDateFrom => _getDateFrom;

    public string GetDateTo => _getDateTo;

    public string GetDSPeriod => _getDSPeriod;

    public Datemanipulation()
    {
        _getPrevMonth = DateTime.Now.AddMonths(-1).ToString("MM");
        _getPrevYear = DateTime.Now.AddMonths(-1).ToString("yyyy");
        _getFirstDateofPrevMonth = DateTime.Now.AddMonths(-1).ToString("MM");
        _getLastDateofPrevMonth = new  DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddDays(-1).ToString("DD");
        _getDateFrom = $@"{_getFirstDateofPrevMonth}/{_getPrevMonth}/{_getPrevYear} 00:00";
        _getDateTo = $@"{_getLastDateofPrevMonth}/{_getPrevMonth}/{_getPrevYear} 00:00";
        _getDSPeriod = $@"{_getPrevYear}{_getPrevMonth}";
    }

}




