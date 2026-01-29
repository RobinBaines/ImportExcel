//
// @Copyright 2026 Robin Baines
// Licensed under the MIT license. See MITLicense.txt file in the project root for details.
//
//------------------------------------------------
//Name: Module for ImportExcel service. Program.cs
//Function: 
//Created Jan 2018.
//Notes: 
//Modifications:
//------------------------------------------------
using System;
using System.ServiceProcess;
using System.Threading;

namespace ImportExcel
{
    static class Program
    {
        private static EventWaitHandle _waitHandle;
        private static ImportExcel _service;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {
            var runConsole = false;
            foreach (var arg in args)
            {
                // When -console is given as argument, we want to debug
                if (arg.ToLowerInvariant().Equals("-console"))
                {
                    runConsole = true;
                }
            }

            _service = new ImportExcel();

            if (runConsole)
            {
                // Run in console mode for debugging
                _waitHandle = new EventWaitHandle(false, EventResetMode.ManualReset);

                Console.WriteLine("Starting Workflow Service in Console Mode");
                Console.WriteLine("Press Ctrl+C to exit Console Mode");
                Console.CancelKeyPress += OnCancelKeyPress;

                _service.StartMailServiceThread();

                WaitHandle.WaitAll(new WaitHandle[] { _waitHandle });
            }

            // Start the service
            var ServicesToRun = new ServiceBase[] 
            { 
				new ImportExcel() 
			};
            ServiceBase.Run(ServicesToRun);
        }

        static void OnCancelKeyPress(object sender, ConsoleCancelEventArgs e)
        {
            //DatFileProcessor.ShouldRun = false;
            _service.Stop();
            _waitHandle.Set();
        }
    }
}
