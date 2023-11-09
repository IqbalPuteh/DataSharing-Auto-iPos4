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
        //ProcessStartInfo startInfo = new ProcessStartInfo();
        //startInfo.FileName = "7z.exe";
        //startInfo.Arguments = $"a -tzip \"{destinationPath}\" \"{sourcePath}\" -mx=9";
        //startInfo.WindowStyle = ProcessWindowStyle.Hidden;
        //Process process = new Process();
        //process.StartInfo = startInfo;
        //process.Start();
        //process.WaitForExit();
        return "";
    }

    public virtual string SendFileToUrl(string filePath, string url)
    {
        //ProcessStartInfo startInfo = new ProcessStartInfo();
        //startInfo.FileName = "curl.exe";
        //startInfo.Arguments = $"-F \"file=@{filePath}\" \"{url}\"";
        //startInfo.WindowStyle = ProcessWindowStyle.Hidden;
        //Process process = new Process();
        //process.StartInfo = startInfo;
        //process.Start();
        //process.WaitForExit();
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

public class MyDateManipulator
{
    private static string GetPrevMonth()
    {
        return DateTime.Now.AddMonths(-1).ToString("MM");
    }

    private static string GetPrevYear()
    {
        return DateTime.Now.AddMonths(-1).ToString("yyyy");
    }


    private static string GetFirstDate()
    {
        return "01";
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

    private static string GetLastDayOfPrevMonth()
    {
        var lastDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddDays(-1);
        return lastDay.ToString("dd");
    }
}




