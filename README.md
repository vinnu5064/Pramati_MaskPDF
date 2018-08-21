# Pramati_MaskPDF
This file contains instructions on how to set up your environment to work with the source code:

Necessary Softwares:

1. Microsoft Visual Studio Community Edition(.Net framework 4.6)
It is used to develop/debug the source code. 

2. ImageMagick v7.0.7 (https://www.imagemagick.org/script/download.php)
It converts PDF files to jpg images and jpg images to PDF.

3. Tesseract v4.00.00 (https://www.imagemagick.org/script/download.php)
It converts jpg images to ocr file.

4. NuGetPackageExplorer v4.4.7 (https://github.com/NuGetPackageExplorer/NuGetPackageExplorer/releases)
In order to create a NuGet package you need to install NuGetPackageExplorer. 

External DLL:
1. HtmlAgilityPack.dll
To use HTML related classes.


Running the tests:

The source code is ready to build as an Uipath Activity. Ensure that below references have been added to the project

  i)   system.Activities
  
 ii)   system.Componentmodel.Composistion
 
iii)   system.HtmlAgilityPack


Deployment:

Once you are done with cloning the source code, build the solution. The Output panel is displayed, informing you that the file is built and displays its path. In our case the Pramati_MaskPDF.dll file is created. 
Follow the steps listed here https://activities.uipath.com/docs/creating-a-custom-activity#section-creating-the-nuget-package to create .nupkg file.
Copy the .nupkg file inside the Packages folder of your UiPath Studio install location (%USERPROFILE%\.nuget\Packages).The NuGet Package containing your custom activity is now ready to be loaded in UiPath Studio.(Refer-https://activities.uipath.com/docs/creating-a-custom-activity#section-loading-the-nuget-package-in-uipath-studio)


License

This project is licensed under the MIT License - see the LICENSE.md file for details
