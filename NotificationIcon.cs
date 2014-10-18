/*
 * NotificationIcon.cs
 * TextFileXpander
 *
 * Created by wanswings on 2014/08/21.
 * Copyright (c) 2014 wanswings. All rights reserved.
 */
using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace TextFileXpander
{
	public sealed class NotificationIcon
	{
		[DllImport("user32.dll") ]
		private static extern IntPtr GetForegroundWindow();
		[DllImport("user32.dll")]
		private static extern bool SetForegroundWindow(IntPtr hWnd);
		[DllImport("user32.dll")]
		private static extern Int32 GetWindowThreadProcessId(IntPtr hWnd, out Int32 lpdwProcessId);

		private static System.Timers.Timer aTimer;
		private static IntPtr lastHwnd;
		private static Int32 lastPid;
		private static string[] ignoreProcs = {"TextFileXpander", "explorer"};
		// HKEY_CURRENT_USER\Software\wanswings\TextFileXpander
		private static string regKeyProc = @"Software\wanswings\TextFileXpander";
		// HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run
		private static string regKeyStartup = @"Software\Microsoft\Windows\CurrentVersion\Run";

		private NotifyIcon notifyIcon;
		private ContextMenuStrip fpMenu;
		private Bitmap imgLaunchAtStartup;
		private int idxLaunchAtStartup;

		#region Initialize
		public NotificationIcon()
		{
			string appName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
			notifyIcon = new NotifyIcon();
			notifyIcon.Text = appName;
			fpMenu = new ContextMenuStrip();
			ComponentResourceManager resources = new ComponentResourceManager(typeof(NotificationIcon));
			imgLaunchAtStartup = (Bitmap)resources.GetObject("Check.Image");
			InitializeMenu();
			notifyIcon.Icon = (Icon)resources.GetObject("$this.Icon");
			notifyIcon.ContextMenuStrip = fpMenu;
			notifyIcon.MouseUp += new MouseEventHandler(notifyIconMouseUp);

			// 0.5sec timer
			aTimer = new System.Timers.Timer(500);
			aTimer.Elapsed += new System.Timers.ElapsedEventHandler(OnTimedEvent);
			aTimer.Enabled = true;
		}

		private void InitializeMenu()
		{
			fpMenu.Items.Clear();
			Regex regexp1 = new Regex("^(-{2}-+)\\s*(.*)", RegexOptions.IgnoreCase);
			Regex regexp2 = new Regex("^marker:(strong:|weak:)?\\s*(.+)", RegexOptions.IgnoreCase);

			int idxMain = 0;
			RegistryKey regkey = Registry.CurrentUser.OpenSubKey(regKeyProc, false);
			if (regkey != null) {
				string dirPath = (string)regkey.GetValue("dirPath");
				if (dirPath != null) {
					Debug.WriteLine("Load path: " + dirPath);
					// Get files
					bool existData = false;
					try {
						string[] dirs = Directory.GetFiles(dirPath, "*.txt", SearchOption.TopDirectoryOnly);
						Array.Sort(dirs);

						foreach (string fullPath in dirs) {
							Debug.WriteLine("File: " + fullPath);
							// Create submenu
							ToolStripMenuItem submenu = new ToolStripMenuItem(Path.GetFileName(fullPath), null, new EventHandler(launchApplication));
							int idxSub = 0;
							StreamReader reader = (new StreamReader(fullPath, System.Text.Encoding.UTF8));
							while (reader.Peek() >= 0) {
								string line = reader.ReadLine();
								if (line.Length > 0) {
									Match match1 = regexp1.Match(line);
									if (match1.Success) {
										ToolStripSeparator submenuSeparator = new ToolStripSeparator();
										submenu.DropDownItems.Add(submenuSeparator);
									}
									else {
										Color fg;
										Match match2 = regexp2.Match(line);
										if (match2.Success) {
											string matchCmd = match2.Groups[1].Value;
											line = match2.Groups[2].Value;
											if (matchCmd.Equals("strong:")) {
												fg = Color.Red;
											}
											else if (matchCmd.Equals("weak:")) {
												fg = Color.LightGray;
											}
											else {
												fg = Color.Blue;
											}
										}
										else {
											fg = Color.Black;
										}
										string itemName;
										if (line.Length > 64) {
											itemName = line.Substring(0, 64) + "...";
										}
										else {
											itemName = line;
										}
										ToolStripMenuItem item = new ToolStripMenuItem(itemName, null, new EventHandler(PushData));
										item.ForeColor = fg;
										item.Tag = line;
										submenu.DropDownItems.Add(item);
									}
									idxSub++;
								}
							}
							reader.Close();

							existData = true;
							submenu.Tag = fullPath;
							fpMenu.Items.Add(submenu);
							idxMain++;
						}
					}
					catch (System.Exception excpt) {
						Console.WriteLine("System error!");
						Console.WriteLine(excpt.Message);
					}
					if (existData) {
						fpMenu.Items.Add(new ToolStripSeparator());
						idxMain++;
					}
				}
				else {
					Debug.WriteLine("Load path: No data!");
				}
				regkey.Close();
			}
			else {
				Debug.WriteLine("Load path: No regkey!");
			}

			fpMenu.Items.Add(new ToolStripMenuItem("Select Directory", null, new EventHandler(SelectDir)));
			idxMain++;
			fpMenu.Items.Add(new ToolStripMenuItem("Refresh", null, new EventHandler(RefreshData)));
			idxMain++;

			Bitmap img = null;
			if (isLaunchAtStartup(false)) {
				img = imgLaunchAtStartup;
			}
			fpMenu.Items.Add(new ToolStripMenuItem("Launch at startup", img, new EventHandler(ToggleLaunchAtStartup)));
			idxLaunchAtStartup = idxMain++;
			fpMenu.Items.Add(new ToolStripMenuItem("Quit", null, new EventHandler(Terminate)));
		}
		#endregion

		#region Main - Program entry point
		/// <summary>Program entry point.</summary>
		/// <param name="args">Command Line Arguments</param>
		[STAThread]
		public static void Main(string[] args)
		{
			lastHwnd = GetForegroundWindow();
			lastPid = GetWindowProcessID(lastHwnd);

			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);

			bool isFirstInstance;
			// Please use a unique name for the mutex to prevent conflicts with other programs
			using (Mutex mtx = new Mutex(true, "TextFileXpander", out isFirstInstance)) {
				if (isFirstInstance) {
					NotificationIcon notificationIcon = new NotificationIcon();
					notificationIcon.notifyIcon.Visible = true;
					Application.Run();
					notificationIcon.notifyIcon.Dispose();
				}
				else {
					// The application is already running
					MessageBox.Show("The application is already running");
				}
			} // releases the Mutex
		}
		#endregion

		#region Event Handlers
		private static void OnTimedEvent(object source, System.Timers.ElapsedEventArgs e)
		{
			IntPtr newHwnd = GetForegroundWindow();
			Int32 newPid = GetWindowProcessID(newHwnd);

			bool ignore = false;
			foreach (string str in ignoreProcs) {
				if (Process.GetProcessById(newPid).ProcessName.Equals(str)) {
					ignore = true;
					break;
				}
			}
			if (!ignore && newPid != lastPid) {
				lastHwnd = newHwnd;
				lastPid = newPid;
				Debug.WriteLine("newPid: " + newPid);
			}
		}

		private void notifyIconMouseUp(object sender, MouseEventArgs e) {
			if (e.Button == MouseButtons.Left) {
				MethodInfo mi = typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic);
				mi.Invoke(notifyIcon, null);
			}
		}

		private static Int32 GetWindowProcessID(IntPtr hwnd)
		{
			Int32 pid;
			GetWindowThreadProcessId(hwnd, out pid);
			return pid;
		}

		private void PushData(object sender, EventArgs e)
		{
			string str = (string)((ToolStripMenuItem)sender).Tag;

			Regex regexp = new Regex("^([a-z]+):\\s*(.+)", RegexOptions.IgnoreCase);
			Match match = regexp.Match(str);
			if (match.Success) {
				string matchCmd = match.Groups[1].Value;
				Debug.WriteLine("matchCmd: " + matchCmd);
				string matchStr = match.Groups[2].Value;
				Debug.WriteLine("matchStr: " + matchStr);

				string sendStr = null;

				if (matchCmd.Equals("dict")) {
					// dict
					str = matchStr;
				}
				else if (matchCmd.Equals("flight")) {
					// flight
					sendStr = "http://www.google.com/search?q=flight%20" + Uri.EscapeUriString(matchStr);
				}
				else if (matchCmd.Equals("mailto")) {
					// mailto
					sendStr = "mailto:" + matchStr;
				}
				else if (matchCmd.Equals("map")) {
					// map
					sendStr = "http://maps.google.com/maps?q=" + Uri.EscapeUriString(matchStr);
				}
				else if (matchCmd.Equals("near")) {
					// near
					sendStr = "http://foursquare.com/explore?near=" + Uri.EscapeUriString(matchStr);
				}
				else if (matchCmd.Equals("people")) {
					// people
					str = matchStr;
				}
				else if (matchCmd.Equals("recipe")) {
					// recipe
					sendStr = "http://www.epicurious.com/tools/searchresults?search=" + Uri.EscapeUriString(matchStr);
				}
				else if (matchCmd.Equals("route")) {
					// route
					Regex regexp2 = new Regex("^\\s*from:\\s*(.+)\\s+to:\\s*(.+)", RegexOptions.IgnoreCase);
					Match match2 = regexp2.Match(matchStr);
					if (match2.Success) {
						string matchfrom = match2.Groups[1].Value;
						Debug.WriteLine("matchfrom: " + matchfrom);
						string matchto = match2.Groups[2].Value;
						Debug.WriteLine("matchto: " + matchto);

						sendStr = "http://maps.google.com/maps?saddr=" + Uri.EscapeUriString(matchfrom)
															+ "&daddr=" + Uri.EscapeUriString(matchto);
					}
				}
				else if (matchCmd.Equals("tel")) {
					// tel
					str = matchStr;
				}
				else if (matchCmd.Equals("twitter")) {
					// twitter
					str = matchStr;
				}
				else if (matchCmd.Equals("url")) {
					// url
					sendStr = matchStr;
				}
				else if (matchCmd.Equals("weather")) {
					// weather
					sendStr = "http://www.weather.com/search/enhancedlocalsearch?where=" + Uri.EscapeUriString(matchStr);
				}
				else if (matchCmd.Equals("youtube")) {
					// youtube
					sendStr = "http://www.youtube.com/results?search_query=" + Uri.EscapeUriString(matchStr);
				}

				if (sendStr != null) {
					Debug.WriteLine(matchCmd + ": " + sendStr);
					try {
						Process p = Process.Start(sendStr);
					}
					catch (System.Exception ex) {
						System.Console.WriteLine(ex.Message);
					}
					return;
				}
			}

			// To Clipboard
			Clipboard.SetText(str);
			Debug.WriteLine("To Pasteboard: " + str);
			// To Active App
			SetForegroundWindow(lastHwnd);
			SendKeys.Send("^v");
		}

		private void launchApplication(object sender, EventArgs e)
		{
			string fullPath = (string)((ToolStripMenuItem)sender).Tag;
			// Launch App
			Debug.WriteLine("Launch text editor with: " + fullPath);
			Process p = Process.Start(fullPath);
		}

		private bool isLaunchAtStartup(bool doAddRemove)
		{
			bool registered = false;
			RegistryKey regkey = Registry.CurrentUser.OpenSubKey(regKeyStartup, true);
			string appName = (string)regkey.GetValue(Application.ProductName);
			if (appName == null) {
				if (doAddRemove) {
					// Add
					regkey.SetValue(Application.ProductName, Application.ExecutablePath);
				}
			}
			else {
				registered = true;
				if (doAddRemove) {
					// Remove
					regkey.DeleteValue(Application.ProductName);					
				}
			}
			regkey.Close();

			if (doAddRemove) {
				return !registered;
			}
			else {
				return registered;
			}
		}

		private void ToggleLaunchAtStartup(object sender, EventArgs e)
		{
			Bitmap img = null;
			if (isLaunchAtStartup(true)) {
				img = imgLaunchAtStartup;
			}
			fpMenu.Items[idxLaunchAtStartup].Image = img;
		}

		private void RefreshData(object sender, EventArgs e)
		{
			InitializeMenu();
		}

		private void SelectDir(object sender, EventArgs e)
		{
			FolderBrowserDialog fbd = new FolderBrowserDialog();
			fbd.Description = "Choose your TextFileXpander folder";
			fbd.ShowNewFolderButton = false;
			if (fbd.ShowDialog() == DialogResult.OK) {
				RegistryKey regkey = Registry.CurrentUser.CreateSubKey(regKeyProc);
				regkey.SetValue("dirPath", fbd.SelectedPath);
				regkey.Close();
				Debug.WriteLine("Save path: " + fbd.SelectedPath);

				InitializeMenu();
			}
		}

		private void Terminate(object sender, EventArgs e)
		{
			Application.Exit();
		}
		#endregion
	}
}
