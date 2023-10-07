using System;
using System.Collections;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;

namespace USBPOC
{
    /// <summary>
    /// Note: System.Management.Automation is incompatible with .NET 6 from v7.3 onwards.
    /// Note: Runspace passes in default InitialSessionState, which has 'Get-PnpDevice'.
    /// Not all commands(modules) work for BlueHippo's approach: https://www.youtube.com/watch?v=KUh-RlBcfI8
    /// </summary>
    public static class PowerShellHandler
    {
        /// <summary>
        /// Main function to be called (for now).
        /// </summary>
        public static async Task<String> RunCommand(string script)
        {
            StringBuilder sb = new();
            InitialSessionState initialSessionState = InitialSessionState.CreateDefault();
            initialSessionState.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Unrestricted;

            //using this try-catch-finally, instead of `using`, as didn't like the nested try blocks in IL.
            var run = RunspaceFactory.CreateRunspace(initialSessionState);

            try
            {
                run.Open();
                var ps = PowerShell.Create(run);

                var err = run.SessionStateProxy.PSVariable.GetValue("error");

                ps.Commands.AddScript(script);
                //ps.AddParameter("SkipEditionCheck"); //Not needed for 'Get-PnpDevice'

                var results = await ps.InvokeAsync();
                run.CloseAsync();

                #region error handling, if any
                //Written this way to only unbox once.
                if (err is not null)
                {
                    ArrayList errList = err as ArrayList;
                    if (errList.Count > 0)
                    {
                        string eList = string.Empty;

                        foreach (var e in errList)
                        {
                            eList += e.ToString();
                        }
                        throw new Exception(eList);
                    }
                }
                #endregion

                foreach (var item in results)
                {
                    sb.AppendLine(item.BaseObject.ToString());
                }

            }
            catch (Exception ex)
            {
                await Console.Out.WriteLineAsync(ex.Message);
            }
            finally
            {
                (run as IDisposable)?.Dispose();
            }


            return sb.ToString().Trim();
        }
    }
}
