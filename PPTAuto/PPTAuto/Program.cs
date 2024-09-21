// See https://aka.ms/new-console-template for more information
using Microsoft.Office;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using VBIDE = Microsoft.Vbe.Interop;
using System;
using System.IO;
using System.Diagnostics;

internal class Program
{
    private static void Main(string[] args)
    {
        Application application = new Application();
        application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityByUI;

        string code = File.ReadAllText(@".\MacroVba\macro.txt");
        string vaName = File.ReadAllText(@".\MacroVba\setting.txt");
        if (!Directory.Exists(@".\PptDatas"))
        {
            Directory.CreateDirectory(@".\PptDatas");
        }

        var currentTime = DateTimeOffset.Now.ToUnixTimeMilliseconds().ToString();
        var filename = AppDomain.CurrentDomain.BaseDirectory + $"/PptDatas/PPT{currentTime}.pptx";

        var ppt = application.Presentations.Add();
        Slide slide = ppt.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
        VBIDE.VBProject vbProject = ppt.VBProject;
        // 创建一个新的 VBA 模块
        VBIDE.VBComponent vbmodule = vbProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);
        // 设置模块的名称
        vbmodule.Name = "MyModule";
        // 运行宏
        vbmodule.CodeModule.AddFromString(code);
        application.Run(vaName);

        ppt.SaveAs(filename);
        Environment.Exit(0);
        Process.GetCurrentProcess().CloseMainWindow();
        
    }
   
}