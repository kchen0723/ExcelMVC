ExcelMVC.AddinDna - NuGet package
=================================
ExcelMvc.AddinDna is an Excel-DNA Addin with functions required for running ExcelMvc applications.

It can be packed in its own Addin or linked into an existing Excel-Dna Addin. 
See Sample.Application.Dna.PostBuild.cmd to see how an Excel-DNA Addin is packed.

Documentation and sample application can be found at the project site (https://sourceforge.net/projects/excelmvc/).

During installation of the ExcelMVC NuGet package the following changes were applied to your project:
1. Dependent package Excel-DNA was installed. Please refer to readme.txt of that package.
2. Added a reference to <package>\lib\<targetFramework>\ExcelMvc.dll.
3. Added a reference to <package>\lib\<targetFramework>\ExcelMvc.AddinDna.dll
2. Added a file ExcelMvc.Addin.xll to your project.

Uninstalling
------------
* If the ExcelMVC.AddinDna NuGet package is uninstalled, the references to ExcelMvc.dll and ExcelMvc.AddinDna.dll as well as the file ExcelMvc.Addin.xll are removed.
* Note that uninstallation of the dependent Excel-DNA package might leave some of the installed modifications in your project. Please refer to readme.txt of Excel-DNA package for detailed information.
