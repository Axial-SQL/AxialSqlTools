// Copyright (C) 2006-2010 Jim Tilander. See COPYING for and README for more details.
using System;

namespace Aurora
{
	public sealed class Singleton<T> where T : class, new()
	{
		private Singleton() {}
		public static T Instance = new T();
	}
}
