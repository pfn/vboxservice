// Copyright 2010 Perry Nguyen <pfnguyen@hanhuy.com>
// All rights reserved

using System;
using System.Linq;
using System.Configuration.Install;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using System.ServiceProcess;
using System.Runtime.InteropServices;

using VirtualBox;

namespace HanHuy.VBoxService {

	delegate void VMAction(IMachine m);

	class MainClass : ServiceBase {
		private  const string EXTRADATA_KEY = "Service";
		internal const string SERVICE_NAME  = "VirtualBox VM service";

		private VirtualBox.VirtualBox vb;
		private IMachine[] Machines;

		public static void Main(string[] args) {
			// try to be smart about commandline vs service invocation
			var c      = new ServiceController(SERVICE_NAME);
			var status = ServiceControllerStatus.Stopped;
			var serviceInstalled = false;

			try {
				status           = c.Status; // invalid op if not installed
				serviceInstalled = true;
			} catch (InvalidOperationException) { }
			c.Close();

			if (serviceInstalled && args.Length == 0 &&
					status != ServiceControllerStatus.StartPending) {
				Usage(serviceInstalled);
				return;
			}
			if (!serviceInstalled && args.Length == 0) {
				Usage(serviceInstalled);
				return;
			}
			if (args.Length == 0) {
				var m = new MainClass();
				try {
					// prevents startup if vbox is not available
				    m.Init();
				} catch {
					m.EventLog.WriteEntry("Unable to load VBox COM object",
							EventLogEntryType.Error);
					return;
				}
				ServiceBase.Run(m);
			} else {
				switch (args[0]) {
				case "-install":
					try {
						// prevent installation if vbox is not available
						new VirtualBox.VirtualBox();
						ManagedInstallerClass.InstallHelper(new string[] {
								"/LogToConsole=false",
								Assembly.GetExecutingAssembly().Location
						});
						Console.WriteLine(SERVICE_NAME + " installed");
					} catch (Exception e) {
						Console.WriteLine("Unable to install service: " + e);
					}
					break;
				case "-uninstall":
					try {
						ManagedInstallerClass.InstallHelper(new string[] {
								"/LogToConsole=false", "/u",
								Assembly.GetExecutingAssembly().Location
						});
						Console.WriteLine(SERVICE_NAME + " uninstalled");
					} catch (Exception e) {
						Console.WriteLine("Unable to uninstall service: " + e);
					}
					break;
				default: Usage(serviceInstalled); break;
				}
			}
		}
		private static void Usage(bool installed) {
			Console.WriteLine("Usage: vboxservice <-install|-uninstall>");
			Console.WriteLine("\r\n\tService installed: " + installed);
		}

		private void Init() {
		    vb          = new VirtualBox.VirtualBox();
			ServiceName = SERVICE_NAME;
			CanShutdown = true;
			CanStop     = true;
			Machines    = vb.Machines;
		}

		private void StartVMs() {
			ProcessVMs(new MachineState[] {
							MachineState.MachineState_PoweredOff,
							MachineState.MachineState_Saved
					}, delegate(IMachine m) {
				var session = new Session();
				EventLog.WriteEntry("Starting VM: " + m.Name);
				try {
					var p = vb.OpenRemoteSession(session, m.Id, "vrdp", "");
					WaitForCompletion(p, true);
				} catch (Exception e) {
					EventLog.WriteEntry(String.Format(
							"Error starting VM {0}\r\n{1}",
							m.Name, e.ToString()),
						EventLogEntryType.Error);
				}
			});
		}

		private void StopVMs(bool requestTime) {
			ProcessVMs(new MachineState[] {
							MachineState.MachineState_Running
					}, delegate(IMachine m) {
				var session = new Session();
				EventLog.WriteEntry("Stopping VM: " + m.Name);
				try {
					vb.OpenExistingSession(session, m.Id);
					WaitForCompletion(
							session.Console.SaveState(), requestTime);
				} catch (Exception e) {
					EventLog.WriteEntry(String.Format(
							"Error saving VM {0}\r\n{1}",
							m.Name, e.ToString()),
						EventLogEntryType.Error);
				}
			});
		}

		private void ProcessVMs(MachineState[] states, VMAction action) {
			var s = new HashSet<MachineState>(states);
			var machines = from m in Machines
					where m.GetExtraData(EXTRADATA_KEY).ToLower() == "yes" &&
						s.Contains(m.State)
					select m;
			foreach (var m in machines) action(m);
		}

		private void WaitForCompletion(IProgress p, bool requestTime) {
			try {
				while (p.Completed == 0 && p.Canceled == 0) {
					if (requestTime && p.TimeRemaining > 3)
						RequestAdditionalTime(1000);
					Thread.Sleep(1000);
				}
	
				if (p.ErrorInfo != null)
					EventLog.WriteEntry("VM operation failed: " + p.ErrorInfo,
							EventLogEntryType.Error);
			} catch (COMException) {
				// means that the VM process has exited (?)
			}
		}

		protected override void OnStart(string[] args) {
			StartVMs();
		}
		protected override void OnStop() {
			StopVMs(true);
		}
		protected override void OnShutdown() {
			StopVMs(false);
		}
	}
	
	[RunInstaller(true)]
	public class MainInstaller : Installer {
		private const string DESCRIPTION =
				"Start selected virtual machines as a service";

		public MainInstaller() {
			var spi = new ServiceProcessInstaller();
			var si  = new ServiceInstaller();

			si.DisplayName = MainClass.SERVICE_NAME;
			si.ServiceName = MainClass.SERVICE_NAME;
			si.Description = DESCRIPTION;
			si.StartType   = ServiceStartMode.Automatic;
			spi.Account    = ServiceAccount.LocalSystem;

			Installers.Add(spi);
			Installers.Add(si);
		}
	}
}
