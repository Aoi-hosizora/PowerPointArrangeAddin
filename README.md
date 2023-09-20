# ppt-arrange-addin

+ A PowerPoint add-in (VSTO) for assisting arrangement operations, which is inspired by [iSlide Addin](https://www.islide.cc/).
+ Development envrionment: Visual Studio 2019, .NET Framework 4.8 (C# 7.3 / C# 9.0), PowerPoint 2010.
+ Requirements: Microsoft Office >= 2010.
+ Supported languaged: English, Simplified Chinese, Traditional Chinese, Japanese.

### Build

```bat
cd ppt-arrange-addin
call vsdevcmd.bat
msbuild ppt-arrange-addin.csproj /p:Configuration=Release /p:Platform=x64
```

### Screenshots

| ![screenshot1](./assets/screenshot1.jpg) | ![screenshot2](./assets/screenshot2.jpg) | ![screenshot3](./assets/screenshot3.jpg) |
|:--:|:--:|:--:|
| "Arrangement" group | "Textbox" group | "Replace picture" group |
| ![screenshot4](./assets/screenshot4.jpg) | ![screenshot5](./assets/screenshot5.jpg) | ![screenshot6](./assets/screenshot6.jpg) |
| "Size and position" group | "Arrangement" menu | Add-in setting dialog |

### References

+ [VSTO开发指南](https://blog.csdn.net/fuhanghang/article/details/101533271)
+ [VSTO之旅系列(三)：自定义Excel UI](https://blog.51cto.com/learninghard/1144298)
+ [Walkthrough: Create your first VSTO Add-in for PowerPoint](https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-powerpoint)
+ [Application object (PowerPoint)](https://learn.microsoft.com/en/office/vba/api/powerpoint.application)
+ [How to: Customize a built-in tab](https://github.com/MicrosoftDocs/visualstudio-docs/blob/main/docs/vsto/how-to-customize-a-built-in-tab.md)
+ [C# VSTO Add-in Excel: What is the name of this Excel Super-Tab Control and how to make it?](https://stackoverflow.com/questions/61189402/c-sharp-vsto-add-in-excel-what-is-the-name-of-this-excel-super-tab-control-and)
+ [Playlist: VBA to .NET / VSTO](https://www.youtube.com/playlist?list=PLo0aMPtFIFDqaRyd0KZ0DLXFD3rfhI4SU)
+ [Walkthrough: Update the controls on a ribbon at run time](https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-updating-the-controls-on-a-ribbon-at-run-time)
+ [C# 8/9の言語機能を.NET Frameworkで使う](https://qiita.com/kenichiuda/items/fada6068ea265fd6a389)
+ [Custom UI XML Markup Specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/31f152d6-2a5d-4b50-a867-9dbc6d01aa43)
+ [Replace a picture on a slide in PowerPoint using VSTO](https://stackoverflow.com/questions/76696349/replace-a-picture-on-a-slide-in-powerpoint-using-vsto)
+ [Open Specifications - Customui Schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/5f3e35d6-70d6-47ee-9e11-f5499559f93a)
+ [Registry entries for VSTO Add-ins](https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins)
