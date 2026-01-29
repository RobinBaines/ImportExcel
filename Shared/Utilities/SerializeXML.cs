//
// @Copyright 2026 Robin Baines
// Licensed under the MIT license. See MITLicense.txt file in the project root for details.
//
//------------------------------------------------
//Name: Module for SerializeXml.cs
//Function: 
//Created Jan 2018.
//Notes: 
//Modifications:
//------------------------------------------------
using System;
using System.IO;
using System.Text;
namespace Shared.Utilities
{
    public class SerializeXml
    {
        /// <summary>
        /// XML serialize an object to a file
        /// </summary>
        /// <param name="str">filename to serialize to</param>
        /// <param name="o">object to serialize</param>
        public static void ToFile(string str,object o)
        {
            if (!Directory.Exists(Path.GetDirectoryName(str)))
                Directory.CreateDirectory(Path.GetDirectoryName(str));
            TextWriter wr = new StreamWriter(File.Open(str, FileMode.Create));
            ToTextWriter(wr, o);
            wr.Close();
        }

        /// <summary>
        /// Serialize an object to a text writer
        /// </summary>
        /// <param name="wr">text writer to serialize to</param>
        /// <param name="o">object to serialize</param>
        public static void ToTextWriter(TextWriter wr, object o)
        {
            //
            // XML export the Split Info
            //
            System.Xml.Serialization.XmlSerializer xmser = new System.Xml.Serialization.XmlSerializer(o.GetType());
            xmser.Serialize(wr, o);
        }

        /// <summary>
        /// De-Serialize an object from a file
        /// </summary>
        /// <param name="str">filename </param>
        /// <param name="t">type of the object ( object.GetType() )</param>
        /// <returns>de-serialized object</returns>
        public static object ReadObject(string str,Type t)
        {
            object o;
            TextReader rd=new StreamReader(File.Open(str,FileMode.Open));
            o = ReadObject(rd,t);
            rd.Close();
            return(o);
        }

        /// <summary>
        /// De-Serialize an object from a text reader
        /// </summary>
        /// <param name="rd">Text reader to read from</param>
        /// <param name="t">type of the object ( object.GetType() )</param>
        /// <returns>de-serialized object</returns>
        public static object ReadObject(TextReader rd, Type t)
        {
            object o;
            System.Xml.Serialization.XmlSerializer xmser = new System.Xml.Serialization.XmlSerializer(t);
            o = xmser.Deserialize(rd);
            return (o);
        }

        public static object FromString(string str, Type t)
        {
            byte[] data = Encoding.UTF8.GetBytes(str);
            MemoryStream mstr = new MemoryStream(data);
            StreamReader rd = new StreamReader(mstr);
            Object o = ReadObject(rd, t);
            rd.Close();
            //mstr.Close();
            return (o);
        }

        public static string ToString(Object o)
        {
            MemoryStream mstr = new MemoryStream();
            StreamWriter wr = new StreamWriter(mstr);
            ToTextWriter(wr, o);
            wr.Close();
            byte[] data = mstr.GetBuffer();
            string res = Encoding.UTF8.GetString(data);
            return (res);
        }
    }

}
