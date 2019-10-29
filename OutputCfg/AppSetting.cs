﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public class AppSetting {
    public Dictionary<string, string> Settings = new Dictionary<string, string>();

    public static AppSetting Ins {
        get {
            if (ins == null) {
                ins = new AppSetting();
            }

            return ins;
        }
    }

    static AppSetting ins;

    [NonSerialized]
    public static string settingPath = AppDomain.CurrentDomain.BaseDirectory + "/" + "AppSetting.txt";


    private AppSetting() {
        InitSetting();
    }


    public string GetValue(string name) {
        if (Ins.Settings.ContainsKey(name)) {
            return Ins.Settings[name];
        }
        else {

            return string.Empty;
        }
    }

    void InitSetting() {
        try {
            using (StreamReader sr = new StreamReader(settingPath)) {
                string content = sr.ReadToEnd();
                string[] settings = content.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var item in settings) {
                    if (item.StartsWith("//")) { continue; }

                    string[] item2 = item.Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);
                    if (item2.Length == 2) {
                        Settings.Add(item2[0], item2[1]);
                    }
                }
            }
        }
        catch (Exception ex) {
            throw ex;
        }
    }
}
