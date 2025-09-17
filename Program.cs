using System;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using Microsoft.Extensions.Configuration;

namespace RemoteMailFetcher
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load settings from appsettings.json
            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            string remoteComputer = config["WinRM:RemoteComputer"];
            string username = config["WinRM:Username"];
            string password = config["WinRM:Password"];
            string remoteFile = config["WinRM:RemoteFile"];
            int interval = int.Parse(config["WinRM:IntervalSeconds"] ?? "10");

            Console.WriteLine($"🔗 Monitoring {remoteFile} on {remoteComputer} every {interval} seconds...");

            var securePassword = new System.Security.SecureString();
            foreach (char c in password)
                securePassword.AppendChar(c);

            var credentials = new PSCredential(username, securePassword);

            var connectionInfo = new WSManConnectionInfo(
                new Uri($"http://{remoteComputer}:5985/wsman"),
                "http://schemas.microsoft.com/powershell/Microsoft.PowerShell",
                credentials);

            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Default;

            int lastLineCount = 0;

            while (true)
            {
                try
                {
                    using (var runspace = RunspaceFactory.CreateRunspace(connectionInfo))
                    {
                        runspace.Open();
                        using (var ps = PowerShell.Create())
                        {
                            ps.Runspace = runspace;
                            ps.AddScript($"Get-Content -Path '{remoteFile}'");
                            var results = ps.Invoke();

                            if (!ps.HadErrors)
                            {
                                int currentLineCount = results.Count;

                                if (currentLineCount > lastLineCount)
                                {
                                    Console.WriteLine($"\n📩 New mail lines detected at {DateTime.Now}:");

                                    // Print only the new lines
                                    for (int i = lastLineCount; i < currentLineCount; i++)
                                    {
                                        Console.WriteLine(results[i].ToString());
                                    }

                                    lastLineCount = currentLineCount;
                                }
                            }
                            else
                            {
                                Console.WriteLine("⚠️ Error reading remote file");
                                foreach (var err in ps.Streams.Error)
                                    Console.WriteLine(err.ToString());
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Exception: {ex.Message}");
                }

                System.Threading.Thread.Sleep(interval * 1000);
            }
        }
    }
}
