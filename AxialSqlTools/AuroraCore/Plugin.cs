// Copyright (C) 2006-2010 Jim Tilander. See COPYING for and README for more details.
using System;
using System.Collections.Generic;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Windows.Forms;

namespace Aurora
{
	// Wrapper class around registering other classes to handle the actual commands.
	// Interfaces with visual studio and handles the dispatch.
	public class Plugin
	{
		private DTE2 m_application;
		private OutputWindowPane m_outputPane;
		private IVsProfferCommands3 m_profferCommands;
		private OleMenuCommandService m_oleMenuCommandService;
		private string m_panelName;
		private ImageList m_icons;

		private Dictionary<string, CommandBase> m_commands = new Dictionary<string, CommandBase>();
		private Dictionary<string, Feature> m_features = new Dictionary<string, Feature>();
		private string m_connectPath;
		private object mOptions;

		public OutputWindowPane OutputPane { get { return GetOutputPane(); } }
		public string Prefix { get { return m_connectPath; } }
		public DTE2 App	{ get { return m_application; }	}
		public Commands Commands {	get { return m_application.Commands; }}
		public OleMenuCommandService MenuCommandService { get { return m_oleMenuCommandService; } }
		public IVsProfferCommands3 ProfferCommands { get { return m_profferCommands; } }
		public object Options { get { return mOptions; } set { mOptions = value; } }
		public ImageList Icons { get { return m_icons; } }

		public Plugin(DTE2 application, IVsProfferCommands3 profferCommands, ImageList icons, OleMenuCommandService oleMenuCommandService, string panelName, string connectPath, object options)
		{
			// TODO: This can be figured out from traversing the assembly and locating the Connect class...
			m_connectPath = connectPath;

			m_application = application;
			m_panelName = panelName;
			m_profferCommands = profferCommands;
			m_oleMenuCommandService = oleMenuCommandService;
			mOptions = options;
			m_icons = icons;
		}

		public void AddFeature(Feature feature)
		{
			m_features.Add(feature.Name, feature);
		}

		public void RegisterCommand(string commandName, CommandBase command)
		{
			m_commands.Add(commandName, command);
		}

		public bool CanHandleCommand(string commandName)
		{
			// TODO: Gotta be a better way to do this... std::find anyone?
			foreach(string key in m_commands.Keys)
			{
				if(commandName.EndsWith("." + key))
					return true;
			}

			return false;
		}

		public bool IsCommandEnabled(string commandName)
		{
			foreach(string key in m_commands.Keys)
			{
				if(commandName.EndsWith("." + key))
					return m_commands[key].IsEnabled();
			}

			return false;
		}

		public bool OnCommand(string name)
		{
			// TODO: Gotta be a better way to do this... std::find anyone?
			foreach(string key in m_commands.Keys)
			{
				if(name.EndsWith("." + key))
				{
					m_commands[key].OnCommand();
					return true;
				}
			}

			return false;
		}

		public CommandEvents FindCommandEvents(string commandName)
		{
			CommandEvents events = null;
			try
			{
				Command command = App.DTE.Commands.Item(commandName, -1);
				if(command != null)
					events = App.DTE.Events.get_CommandEvents(command.Guid, command.ID);
			}
			catch
			{
			}
			return events;
		}

		private Command GetCommand(string commandName)
		{
			Commands2 commands = (Commands2)m_application.Commands;

			string fullName = m_connectPath + "." + commandName;

			try
			{
				Command command = commands.Item(fullName, 0);
				return command;
			}
			catch(System.ArgumentException)
			{
				return null;
			}
		}

		private bool IsCommandRegistered(string commandName)
		{
			return GetCommand(commandName) != null;
		}

		public static bool HasCommand(CommandBar commandBar, string caption)
		{
			foreach(CommandBarControl control in commandBar.Controls)
			{
				if(control.Caption == caption)
					return true;
			}
			return false;
		}

		public CommandBar AddCommandBar(string name, MsoBarPosition position)
		{
            ThreadHelper.ThrowIfNotOnUIThread();
            CommandBars cmdBars = (Microsoft.VisualStudio.CommandBars.CommandBars)m_application.CommandBars;
			CommandBar existingCmdBar = null;
			try {
				existingCmdBar = cmdBars[name];
			} catch (Exception) {
			}

			if (existingCmdBar != null)
			{
				return existingCmdBar;
			}
			else
			{
				object CmdBarObj;
				int result = ProfferCommands.AddCommandBar(name, (uint)vsCommandBarType.vsCommandBarTypeToolbar, null, 0, out CmdBarObj);
				if (result != Microsoft.VisualStudio.VSConstants.S_OK)
					System.Diagnostics.Debug.WriteLine("Unable to add command bar - already exists, perhaps?");
				return CmdBarObj as CommandBar;
			}

		}

        public CommandBar AddCommandBarMenu(string name, MsoBarPosition position, CommandBar parentBar)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CommandBars cmdBars = (Microsoft.VisualStudio.CommandBars.CommandBars)m_application.CommandBars;
            CommandBar existingCmdBar = null;
            try
            {
                existingCmdBar = cmdBars[name];
            }
            catch (Exception)
            {
            }

            if (existingCmdBar != null)
            {
                return existingCmdBar;
            }
            else
            {
                object CmdBarObj;
                int result = ProfferCommands.AddCommandBar(name, (uint)vsCommandBarType.vsCommandBarTypeMenu, parentBar, (uint)parentBar.Controls.Count, out CmdBarObj);
                if (result != Microsoft.VisualStudio.VSConstants.S_OK)
                    System.Diagnostics.Debug.WriteLine("Unable to add command bar - already exists, perhaps?");

                existingCmdBar = CmdBarObj as CommandBar;


                //existingCmdBar.Picture = ImageConverter.LoadPictureFromImage(Plugin.Icons.Images[IconIndex]);
                //existingCmdBar.Style = MsoButtonStyle.msoButtonIconAndCaption;

                return existingCmdBar;
            }

        }

        private OutputWindowPane GetOutputPane()
		{
			if (m_outputPane != null)
				return m_outputPane;
			try
			{
				m_outputPane = AquireOutputPane(m_application, m_panelName);
			}
			catch (Exception)
			{
			}
			return m_outputPane;
		}

		private static OutputWindowPane AquireOutputPane(DTE2 app, string name)
		{
            ThreadHelper.ThrowIfNotOnUIThread();
            if ("" == name)
				return null;

			OutputWindowPane result = FindOutputPane(app, name);
			if(null != result)
				return result;

			OutputWindow outputWindow = (OutputWindow)app.Windows.Item(EnvDTE.Constants.vsWindowKindOutput).Object;
			OutputWindowPanes panes = outputWindow.OutputWindowPanes;
			return panes.Add(name);
		}

		public static OutputWindowPane FindOutputPane(DTE2 app, string name)
		{
            ThreadHelper.ThrowIfNotOnUIThread();
            if ("" == name)
				return null;

			OutputWindow outputWindow = (OutputWindow)app.Windows.Item(EnvDTE.Constants.vsWindowKindOutput).Object;
			OutputWindowPanes panes = outputWindow.OutputWindowPanes;

			foreach(OutputWindowPane pane in panes)
			{
				if(name != pane.Name)
					continue;

				return pane;
			}

			return null;
		}

		public static void AddKeyboardBinding(Command command, string binding)
		{
			command.Bindings = binding;
		}
	}
}
