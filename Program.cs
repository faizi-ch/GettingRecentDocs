using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace GettingRecentDocs
{
    class Program
    {
        static void Main(string[] args)
        {
            string registry_key = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs";
            /*using (Microsoft.Win32.RegistryKey key = Registry.CurrentUser.OpenSubKey(registry_key))
            {
                foreach (string subkey_name in key.GetSubKeyNames())
                {
                    using (RegistryKey subkey = key.OpenSubKey(subkey_name))
                    {
                        string[] ss=subkey.Name.Split('\\');
                        //Console.WriteLine(ss[7]);
                        if (!ss[7].Contains("com"))
                        {
                            Console.WriteLine(BytesToStringConverted(GetRecentlyOpenedFile(ss[7])));
                        }
                        
                        //using (RegistryKey sbkey_name = key.OpenSubKey(subkey.Name))
                        //{
                        //using (string exname = subkey.Name)
                        //{
                        //Console.WriteLine(sbkey_name.Name);
                        //}
                        //}

                    }
                }
            }*/
            //Console.WriteLine(GetRecentlyOpenedFile(".mp3"));

            RegistryKey rk;
            RegistryKey FolderKey = Registry.CurrentUser.OpenSubKey(registry_key, true); 
            Dictionary<String, String> FileToIndexMapping = new Dictionary<String, String>();
            Dictionary<String, String> IndexToFileMapping = new Dictionary<String, String>();
            List<string>  MruListEx = new List<string>();
            List<string>  sFileNames = new List<string>();
            
            //get mru binary list MRUListEx
            byte[] Bytes = (byte[])FolderKey.GetValue("MRUListEx");

            //build MruListEx list
            for (int index = 0; index < Bytes.Length /
                                sizeof(UInt32); index++)
            {
                UInt32 val = RegTools.GetMruEntry(index, ref Bytes);
                string sVal = val.ToString();
                MruListEx.Add(sVal);
            }

            //get list of keyValues under this registry key
            //and create dictionary mapping between filename and indexname
            List<string> slist = RegTools.ExtractKeyValues(FolderKey);
            foreach (string s in slist)
            {
                object obj = FolderKey.GetValue(s);
                string filename =
                       RegTools.ExtractUnicodeStringFromBinary(obj);
                if (!FileToIndexMapping.ContainsKey(filename))
                {
                    FileToIndexMapping.Add(filename, s);
                
                }
                
                IndexToFileMapping.Add(s, filename);
                sFileNames.Add(filename);
            }

            List<string> sIndexNames = RegTools.ExtractKeyValues(FolderKey);

            foreach (var f in sFileNames)
            {
                Console.WriteLine(f);
            }
            Console.WriteLine(sFileNames.Count);
            Console.ReadKey();
        }

        public static byte[] GetRecentlyOpenedFile(string extention)
        {
            RegistryKey regKey = Registry.CurrentUser;
            byte[] recentFile = null;
            regKey = regKey.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RecentDocs");

            if (string.IsNullOrEmpty(extention))
                extention = ".docs";

            RegistryKey myKey = regKey.OpenSubKey(extention);

            if (myKey == null && regKey.GetSubKeyNames().Length > 0)
                myKey = regKey.OpenSubKey(
                               regKey.GetSubKeyNames()[regKey.GetSubKeyNames().Length - 2]);
            if (myKey != null)
            {
                string[] names = myKey.GetValueNames();
                if (names != null && names.Length > 0)
                {
                    recentFile = (byte[])myKey.GetValue(names[names.Length - 2]);

                }

            }

            return recentFile;
        }
        static string BytesToStringConverted(byte[] bytes)
        {
            using (var stream = new MemoryStream(bytes))
            {
                using (var streamReader = new StreamReader(stream))
                {
                    return streamReader.ReadToEnd();
                }
            }
        }
    }

    class RegTools
    {
        /// <summary>
        /// Given a binary registry key, extract a Unicode string from beginning 
        /// </summary>
        /// <param name="keyObj">the registry key to be examined.</param>
        public static String ExtractUnicodeStringFromBinary(object keyObj)
        {
            String Value = keyObj.ToString();    //get object value 
            string Type = keyObj.GetType().ToString();  //get object type
            Type = Type.Substring(7, Type.Length - 7); //strip off "System."

            if (Type == "Byte[]")
            {
                Value = "";
                byte[] Bytes = (byte[])keyObj;
                //this seems crude but cannot find a way to 'cast' a Unicode string to byte[]
                //even in case where we know the beginning format is Unicode
                //so do it the hard way

                char[] chars = Encoding.Unicode.GetChars(Bytes);
                foreach (char bt in chars)
                {
                    if (bt != 0)
                    {
                        Value = Value + bt; //construct string one at a time
                    }
                    else
                    {
                        break;  //apparently found 0,0 (end of string)
                    }
                }
            }
            return Value;
        }

        /// <summary>
        ///Compare two byte arrays for equality
        /// </summary>
        /// <param name="b1">array 1.</param>
        /// <param name="b2">array 2.</param>
        /// <returns>true if equal</returns>
        public static bool IsBinaryMatch(byte[] b1, byte[] b2)
        {
            bool isMatch = true;   //assume success
            int s1 = b1.Length;
            int s2 = b2.Length;
            if (s1 == s2)
            {
                isMatch = true;   //assume success
                //can't use foreach since byte[] does not implement IEnumerator
                for (int i = 0; i < s1; i++)
                {
                    if (b1[i] != b2[i])
                    {
                        isMatch = false;    //mismatch
                        break;
                    }
                }
            }
            else
            {
                isMatch = false;    //mismatch
            }
            return isMatch;
        }

        /// <summary>
        /// Given a binary registry key, return a List of kay values of interest
        /// these are values with a numerical representation like 0,1,99,128, etc
        /// this will ignore the MRUListEx and ViewStream key values
        /// </summary>
        /// <param name="rkey">the registry key to be examined.</param>
        /// <returns>"List of subkeys names of interest</returns>
        public static List<string> ExtractKeyValues(RegistryKey rkey)
        {
            List<string> sRecentList = new List<string>();
            if (rkey != null)
            {
                string[] SubKeys = rkey.GetValueNames();
                foreach (String s in SubKeys)
                {			//subkey value name
                    if (Regex.IsMatch(s, "[0-9]"))
                    {
                        sRecentList.Add(s);
                    }
                }
            }
            return sRecentList;
        }

        /// <summary>
        /// rebuild binary MRUListEx to be written to registry
        /// </summary>
        /// <param name="rKey">the registry key in which value MRUListEx is to be updated</param>
        /// <param name="MruListEx">the List version of MRUListEx</param>
        public static void RebuildMruList(RegistryKey rKey, List<string> MruListEx)
        {
            byte[] mruUpdated = new byte[MruListEx.Count * 4]; //last entry 0xffff
            for (int index = 0; index < MruListEx.Count; index++)
            {
                string sVal = MruListEx[index];
                UInt32 val = UInt32.Parse(sVal);
                SetMruEntry(index, val, ref mruUpdated);
            }
            rKey.SetValue("MRUListEx", mruUpdated);   //write new MRUListEx to registry

        }

        /// <summary>
        /// Given an index and value, write a little endian encoded stream to supplied mru array
        /// </summary>
        /// <param name="index">starting index into mru[].</param>
        /// <param name="value">value to be written at index into mru[].</param>
        /// <param name="mru">the byte array to be modified.</param>
        public static void SetMruEntry(int index, UInt32 value, ref byte[] mru)
        {
            mru[index * 4 + 0] = (byte)(value & 0xff);
            mru[index * 4 + 1] = (byte)((value >> 8) & 0xff);
            mru[index * 4 + 2] = (byte)((value >> 16) & 0xff);
            mru[index * 4 + 3] = (byte)((value >> 24) & 0xff);
        }

        /// <summary>
        /// Given an index, read a little endian encoded stream from supplied mru array
        /// </summary>
        /// <param name="index">starting index into mru[].</param>
        /// <param name="mru">the mru[] to be updated</param>
        /// <returns>"UInt32 decoded value</returns>
        public static UInt32 GetMruEntry(int index, ref byte[] mru)
        {
            UInt32 value;
            value = mru[index * 4 + 0];
            value += (UInt32)(mru[index * 4 + 1]) << 8;
            value += (UInt32)(mru[index * 4 + 2]) << 16;
            value += (UInt32)(mru[index * 4 + 3]) << 24;
            return value;
        }

    }
}
