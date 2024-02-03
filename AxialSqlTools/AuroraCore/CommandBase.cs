// Copyright (C) 2006-2010 Jim Tilander. See COPYING for and README for more details.
using System;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using stdole;
using System.Windows.Forms;
using System.Drawing;

namespace Aurora
{

	public abstract class CommandBase
	{
		private readonly Plugin mPlugin;
		private readonly string mName;
		private readonly string mCanonicalName;
		private readonly string mTooltip;

		public Plugin Plugin	{ get { return mPlugin; } }
		public string Name		{ get { return mName; } }
		public string CanonicalName { get { return mCanonicalName; } }
		public string AbsName { get { return mPlugin.Prefix + "." + mCanonicalName; } }
		public string Tooltip	{ get { return mTooltip; } }
		public virtual int IconIndex { get { return 0; } }
		public virtual int FolderIconIndex { get { return 0; } }
		public virtual int ScriptIconIndex { get { return 0; } }

        public CommandBase(string name, string canonicalName, Plugin plugin, string tooltip)
		{
			mName = name;
			mCanonicalName = canonicalName;
			mPlugin = plugin;
			mTooltip = tooltip;
		}

		public virtual bool RegisterGUI(OleMenuCommand vsCommand, CommandBar vsCommandbar, bool toolBarOnly)
		{
			// The default command is registered in the toolbar.
			if (IconIndex >= 0 && toolBarOnly)
			{
                _RegisterGUIBar(vsCommand, vsCommandbar);

            }

			return true;
		}

		public virtual void BindToKeyboard(Command vsCommand)
		{
		}

		protected void AssignIcon(CommandBarButton target, int IconIndex)
		{
			target.Picture = ImageConverter.LoadPictureFromImage(Plugin.Icons.Images[IconIndex]);
            target.Style = MsoButtonStyle.msoButtonIconAndCaption;
        }

		public abstract bool OnCommand();	// returns if the command was dispatched or not.
		public abstract bool IsEnabled();	// is the command active?

		protected void _RegisterGUIBar(OleMenuCommand vsCommand, CommandBar vsCommandbar)
		{
            ThreadHelper.ThrowIfNotOnUIThread();
            CommandBarButton bn = null;
			try
			{
				CommandBarControl existingButton = vsCommandbar.Controls[CanonicalName];
				bn = existingButton as CommandBarButton;
			} catch (Exception) {
			}

			if (bn == null)
			{
				object newButton;
				int res = Plugin.ProfferCommands.AddCommandBarControl(CanonicalName, vsCommandbar, (uint)vsCommandbar.Controls.Count, (uint)vsCommandBarType.vsCommandBarTypeToolbar, out newButton);
				if (res != 0)
				{
					// something went wrong
				}
				bn = newButton as CommandBarButton;
			}
			if (bn != null)
			{
				AssignIcon(bn, ScriptIconIndex);
			}
		}

		protected void _RegisterGuiContext(OleMenuCommand vsCommand, string name)
		{
			CommandBar b = ((CommandBars)Plugin.App.CommandBars)[name];
			if(null != b)
			{
                _RegisterGUIBar(vsCommand, b);
			}
		}
	}

	/// <summary>
	/// Gives access to the otherwise protected GetIPictureDispFromPicture()
	/// </summary>
	class ImageConverter : AxHost
	{
		private ImageConverter() : base("63109182-966B-4e3c-A8B2-8BC4A88D221C")
		{
		}

		public static StdPicture LoadPictureFromImage(Image i)
		{
			return (StdPicture)GetIPictureDispFromPicture(i);
		}
	}
}
