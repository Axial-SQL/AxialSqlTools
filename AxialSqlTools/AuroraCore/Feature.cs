using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using EnvDTE80;

namespace Aurora
{
	public abstract class Feature
	{
		private string mName;
		private string mTooltip;

		public string Name { get { return mName; } }
		
		public Feature(string name, string tooltip)
		{
			mName = name;
			mTooltip = tooltip;
		}

		public virtual bool Execute()
		{
			return true;
		}
	};

	public abstract class PreCommandFeature : Feature
	{
		private List<CommandEvents> mEvents = new List<CommandEvents>();
		protected Plugin mPlugin;

		public PreCommandFeature(Plugin plugin, string name, string tooltip)
			: base(name, tooltip)
		{
			mPlugin = plugin;
		}

		protected void RegisterHandler(string commandName, _dispCommandEvents_BeforeExecuteEventHandler handler)
		{
			CommandEvents events = mPlugin.FindCommandEvents(commandName);
			if(null == events)
				return;
			events.BeforeExecute += handler;
			mEvents.Add(events);
		}
	};
}
