using System;
using System.IO;
using System.Diagnostics;
using System.Runtime.CompilerServices;

public abstract class DirectoryManipulator
{
    private string _status;

    public string Status
    {
        get { return _status; }
        //private set { _status = value; }
    }

    private void SetStatus(string status)
    {
        _status = status;
    }
    public virtual void CreateDirectoryIfNotExist(string path)
    {
        if (!Directory.Exists(path))
        {
            Directory.CreateDirectory(path);
        }
    }

    public virtual void DeleteFiles(string path, string extension)
    {
        DirectoryInfo directory = new DirectoryInfo(path);
        foreach (FileInfo file in directory.GetFiles($"*.{extension}"))
        {
            file.Delete();
        }
        foreach (DirectoryInfo subDirectory in directory.GetDirectories())
        {
            DeleteFiles(subDirectory.FullName, extension);
        }
    }


    public virtual void ZipDirectory(string sourcePath, string destinationPath)
    {
        ProcessStartInfo startInfo = new ProcessStartInfo();
        startInfo.FileName = "7z.exe";
        startInfo.Arguments = $"a -tzip \"{destinationPath}\" \"{sourcePath}\" -mx=9";
        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
        Process process = new Process();
        process.StartInfo = startInfo;
        process.Start();
        process.WaitForExit();
    }

    public virtual void SendFileToUrl(string filePath, string url)
    {
        ProcessStartInfo startInfo = new ProcessStartInfo();
        startInfo.FileName = "curl.exe";
        startInfo.Arguments = $"-F \"file=@{filePath}\" \"{url}\"";
        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
        Process process = new Process();
        process.StartInfo = startInfo;
        process.Start();
        process.WaitForExit();
        
    }

}
public class MyDirectoryManipulator : DirectoryManipulator
{
    
    public override void CreateDirectoryIfNotExist(string path)
    {
        // Implement your own logic here
        Console.WriteLine($"Creating directory at {path}");
        base.CreateDirectoryIfNotExist(path);
    }

    public override void DeleteFiles(string path, string extension)
    {
        // Implement your own logic here
        Console.WriteLine($"Deleting files with extension {extension} in {path}");
        base.DeleteFiles(path, extension);
    }

    public override void ZipDirectory(string sourcePath, string destinationPath)
    {
        // Implement your own logic here
        Console.WriteLine($"Zipping directory at {sourcePath} to {destinationPath}");
        base.ZipDirectory(sourcePath, destinationPath);
    }

    public override void SendFileToUrl(string filePath, string url)
    {
        // Implement your own logic here
        Console.WriteLine($"Sending file at {filePath} to {url}");
        base.SendFileToUrl(filePath, url);
    }
}


