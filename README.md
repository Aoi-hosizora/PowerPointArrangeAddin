# PowerPoint Arrangement Assistant Add-in

[![Release](https://img.shields.io/github/v/release/Aoi-hosizora/PowerPointArrangeAddin)](https://github.com/Aoi-hosizora/PowerPointArrangeAddin/releases)
[![License](https://img.shields.io/badge/license-mit-blue.svg)](./LICENSE)

+ A PowerPoint add-in (VSTO) for assisting arrangement operations, which is inspired by [iSlide Addin](https://www.islide.cc/).
+ Development environment: .NET Framework 4.8 (C# 7.3 / C# 9.0).
+ Supported languages: English, Simplified Chinese, Traditional Chinese, Japanese.
+ Prerequisite: Microsoft Office >= 2010 (x64), .NET Framework 4.8 Runtime ([click here to install](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48)).

### Install and uninstall

+ Install: download `setup.exe` from [Release](https://github.com/Aoi-hosizora/PowerPointArrangeAddin/releases) or [AppCenter](https://install.appcenter.ms/users/aoihosizora/apps/powerpointarrangeaddin/distribution_groups/public), double click the installer, and install to a specific location, that's done!
+ Uninstall: go to "Start Menu" or "Control Panel > Programs and Features", and choose to uninstall "PowerPoint Arrangement Assistant Add-in".

### Build manually

> Note: Before building, you may need to generate your pfx file for signing.

+ Only build the add-in without anything

```bash
call build.bat

# You can go to "./PowerPointArrangeAddin/bin/x64/Release/" to find the built dll and vsto files.

# Note that this will also register the add-in, you can run `clean.bat` to unregister.
```

+ Build the add-in and the installation

```bash
call build_solution.bat

# You can go to "./PowerPointArrangeAddinSetup/PowerPointArrangeAddinInstallerLauncher/bin/x64/Release/box/" to find the built installer.

# Note this will not register the add-in, you have to use the built setup.exe to install.
```

### Screenshots

<details open>
    <summary>Japanese, PowerPoint 2010</summary>
    <table>
        <tbody>
            <tr>
                <td align="center" colspan="4"><img src="./assets/screenshot1.jpg" alt="screenshot1" /></td>
            </tr>
            <tr>
                <td align="center" colspan="4">Groups in "Arrangement" tab</td>
            </tr>
            <tr>
                <td align="center"><img src="./assets/screenshot2.jpg" alt="screenshot2" /></td>
                <td align="center"><img src="./assets/screenshot3.jpg" alt="screenshot3" /></td>
                <td align="center"><img src="./assets/screenshot4.jpg" alt="screenshot4" /></td>
                <td align="center"><img src="./assets/screenshot5.jpg" alt="screenshot5" /></td>
            </tr>
            <tr>
                <td align="center">"Arrangement" group</td>
                <td align="center">"Textbox" group</td>
                <td align="center">"Replace picture" group</td>
                <td align="center">"Size and position" group</td>
            </tr>
            <tr>
                <td align="center"><img src="./assets/screenshot6.jpg" alt="screenshot6" /></td>
                <td align="center"><img src="./assets/screenshot7.jpg" height="70%" width="70%" alt="screenshot7" /></td>
                <td align="center" colspan="2"><img src="./assets/screenshot8.jpg" alt="screenshot8" /></td>
            </tr>
            <tr>
                <td align="center">"Arrangement" menu</td>
                <td align="center">Add-in setting dialog</td>
                <td align="center" colspan="2">Add-in installation dialog</td>
            </tr>
        </tbody>
    </table>
</details>

<details open>
    <summary>Simplified Chinese, PowerPoint 2019</summary>
    <table>
        <tbody>
            <tr>
                <td align="center" colspan="4"><img src="./assets/screenshot9.jpg" alt="screenshot9" /></td>
            </tr>
            <tr>
                <td align="center" colspan="4">Groups in "Arrangement" tab</td>
            </tr>
            <tr>
                <td align="center"><img src="./assets/screenshot10.jpg" alt="screenshot10" /></td>
                <td align="center"><img src="./assets/screenshot11.jpg" alt="screenshot11" /></td>
                <td align="center"><img src="./assets/screenshot12.jpg" alt="screenshot12" /></td>
                <td align="center"><img src="./assets/screenshot13.jpg" alt="screenshot13" /></td>
            </tr>
            <tr>
                <td align="center">"Arrangement" group</td>
                <td align="center">"Textbox" group</td>
                <td align="center">"Replace picture" group</td>
                <td align="center">"Size and position" group</td>
            </tr>
            <tr>
                <td align="center"><img src="./assets/screenshot14.jpg" alt="screenshot14" /></td>
                <td align="center"><img src="./assets/screenshot15.jpg" height="70%" width="70%" alt="screenshot15" /></td>
                <td align="center" colspan="2"><img src="./assets/screenshot16.jpg" alt="screenshot16" /></td>
            </tr>
            <tr>
                <td align="center">"Arrangement" menu</td>
                <td align="center">Add-in setting dialog</td>
                <td align="center" colspan="2">Add-in installation dialog</td>
            </tr>
        </tbody>
    </table>
</details>

### Tips

+ You are required to add the following PATH if you want to build the project by command line:
    + `C:\Windows\Microsoft.NET\Framework64\v4.x.xxxxx`
    + `...\Microsoft Visual Studio\xxxx\Enterprise\Common7\Tools`
    + `...\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools`
+ If you have some problems on installation or building, try one of the following solutions:
    + **ATTENTION**: Before you perform operation, you must know that some operations are quite **DANGEROUS**, so please **MAKE SURE** that items you are about to modify are only related to "PowerPointArrangeAddin".
    1. Clean the solution and rebuild it, if it don't work, just restart your PC.
    2. Uninstall the add-in from "Control Panel" if "PowerPointArrangeAddin" exists.
    3. Remove following registry entries (keys or values) that are related to "PowerPointArrangeAddin".
        + `HKCU\SOFTWARE\Microsoft\Office\PowerPoint\Addins`
        + `HKCU\SOFTWARE\Microsoft\VSTA`
        + `HKCU\SOFTWARE\Microsoft\VSTO`
        + `HKCR\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0\Components`
        + `HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`
    4. Remove subfolders and files in `C:\Users\<Username>\AppData\Local\Apps\2.0\*` that are related to "PowerPointArrangeAddin".
    5. Move the "Release" or "Publish" folders to a different location and try to install again. (Not Recommended)
    6. Regenerate a temporary key pfx file and try to build again build. (Not Recommended)

### References

+ [【顺其自然~】VSTO开发指南](https://blog.csdn.net/fuhanghang/article/details/101533271)
+ [VSTO之旅系列(三)：自定义Excel UI](https://blog.51cto.com/learninghard/1144298)
+ [Walkthrough: Create your first VSTO Add-in for PowerPoint](https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-powerpoint)
+ [How to: Customize a built-in tab (with list of control IDs)](https://github.com/MicrosoftDocs/visualstudio-docs/blob/main/docs/vsto/how-to-customize-a-built-in-tab.md)
+ [C# VSTO Add-in Excel: What is the name of this Excel Super-Tab Control and how to make it?](https://stackoverflow.com/questions/61189402/c-sharp-vsto-add-in-excel-what-is-the-name-of-this-excel-super-tab-control-and)
+ [【VBA A2Z】Playlist: VBA to .NET / VSTO](https://www.youtube.com/playlist?list=PLo0aMPtFIFDqaRyd0KZ0DLXFD3rfhI4SU)
+ [C# 8/9の言語機能を.NET Frameworkで使う](https://qiita.com/kenichiuda/items/fada6068ea265fd6a389)
+ [Custom UI XML Markup Specification (with Customui Schema)](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/31f152d6-2a5d-4b50-a867-9dbc6d01aa43)
+ [Customizing a Ribbon Through Size Definitions and Scaling Policies](https://learn.microsoft.com/en-us/windows/win32/windowsribbon/windowsribbon-templates)
+ [Registry entries for VSTO Add-ins](https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins)
+ [Customizing the 2007 Office Fluent Ribbon for Developers](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa338202(v=office.12))
+ [Replace a picture on a slide in PowerPoint using VSTO](https://stackoverflow.com/questions/76696349/replace-a-picture-on-a-slide-in-powerpoint-using-vsto)
+ [Exception from HRESULT 0x800A01A8 in PowerPoint solutions](https://www.add-in-express.com/creating-addins-blog/exception-hresult-0x800a01a8/)
+ [<customization> element (Application manifests for Office solutions)](https://learn.microsoft.com/en-us/visualstudio/vsto/customization-element-office-development-in-visual-studio?view=vs-2019)
+ [给VSTO 解决方案指定产品名、发布者以及其他属性信息](https://www.cnblogs.com/monster1799/p/1310866.html)
+ [ClickOnceでの再インストールでエラー発生](https://blog.regrex.jp/2016/09/02/post-972/)
+ [Deploying a VSTO Solution Using Windows Installer](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer?view=vs-2022)
+ [【Visual Studio2017/2019/2022】コマンドラインからビルドすると「8000000A」エラーが発生する](https://juraku-software.net/visual-studio2017-command-build-8000000a-error/)
+ [WiX Toolset v3 Tutorial](https://www.firegiant.com/docs/wix/v3/tutorial/)
+ [WiX でセットアッププロジェクト](https://qiita.com/hiro_t/items/2b51ec2d495eb31a07b0)
+ [【stoneniqiu】随笔分类 - Wix](https://www.cnblogs.com/stoneniqiu/category/522235.html)
+ [Windows Installer手引書 Part.13 カスタムアクションを実行させるタイミング](https://qiita.com/tohshima/items/8d1d7e702d58dc1429d2)
+ [Windows Installer手引書 Part.14 インストール、アンインストールの区別](https://qiita.com/tohshima/items/72d1e7602a48055c55f5)
+ ["Create Shortcut" Checkbox](https://stackoverflow.com/questions/4658220/create-shortcut-checkbox)
+ [(Wix) MSI uninstall is very slow. Log shows slowness is when shortcuts are being removed](https://stackoverflow.com/questions/63581670/wix-msi-uninstall-is-very-slow-log-shows-slowness-is-when-shortcuts-are-being)
+ [External annotations (JetBrains ReSharper)](https://www.jetbrains.com/help/resharper/Code_Analysis__External_Annotations.html)
+ [Visual Studio App Center, Distribute, Release a Build](https://learn.microsoft.com/en-us/appcenter/distribution/uploading)
+ [App Center API Documentation](https://learn.microsoft.com/en-us/appcenter/api-docs/#app-center-openapi-specification-swagger)
+ [C# で TaskDialog を使う](http://grabacr.net/archives/105)
+ [C#: comctl32.dll version 6 in debugger](https://stackoverflow.com/questions/1415270/c-comctl32-dll-version-6-in-debugger)
+ [imageMso検索・imageMso一覧](https://ymrt.jp/imagemso/index.html)
