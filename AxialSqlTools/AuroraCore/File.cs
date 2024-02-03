using System;
using System.IO;
using System.Xml.Serialization;
using System.Collections.Generic;

namespace Aurora
{
	public static class File
	{
		public static Type LoadXML<Type>(string pathname) where Type : class
		{
			using(StreamReader reader = new StreamReader(pathname))
			{
				XmlSerializer serializer = new XmlSerializer(typeof(Type));
				return (Type)serializer.Deserialize(reader);
			}
		}
		
		public static void SaveXML<Type>(string pathname, Type obj) where Type : class 
		{
			using(StreamWriter writer = new StreamWriter(pathname))
			{
				XmlSerializer serializer = new XmlSerializer(typeof(Type));
				serializer.Serialize(writer, obj);
			}
		}
	}
}