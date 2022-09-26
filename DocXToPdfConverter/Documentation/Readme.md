# Report-From-DocX-HTML-To-PDF-Converter - Create custom reports based on Word docx or HTML documents and convert to PDF with .NET CORE

## What's the functionality?

Report from DOCX / HTML to PDF Converter can parse the source document and introduce the dynamic content into predefined __placeholders__ It works on Windows (tested) and should work on Linux and MacOS. Then it can perform the following conversions:

* DOCX to DOCX 
* DOCX to PDF
* DOCX to HTML
* HTML to HTML
* HTML to DOCX
* HTML to PDF

## What do you need to get started?
Don't get scared away that I use LibreOffice, it is easier than you may think!
1. LibreOffice - just get the PORTABLE EDITION as you don't screw up your webserver with an installation. The portable version just runs without any installation. We need LibreOffice for converting from DOCX or from HTML to PDF and DOCX, etc

2. Nuget: 
* Microsoft.NetCore.App
* Document.Format.OpenXml
* System.Drawing.Common

3. The OpenXml PowerTools (thanks to Eric White for this great work), which I already included into the project, the whole code!

## Get started here!

1. Add dependencies to a) Document.Format.OpenXml (by Microsoft). For the OpenXml Powertools by Eric White (which are already included in the project), add System.Drawing.Common (for .NET Core). The regular System.Drawing won't work!

2. Download LibreOffice (https://www.libreoffice.org/download/portable-versions/). I recommend the portable edition as it does not install anything in your server. It is like unzipping files onto your harddrive. Note the path to "soffice.exe" (I don't know what the file is called in Linux / MacOS, probably just soffice. It is an executable to run a headless, mute version of LibreOffice for conversion processes). On my Windows machine, it is under: C:\PortableApps\LibreOfficePortable\App\libreoffice\program\soffice.exe. 

3. Have your templates (Word docx or HTML) ready. In the repository of the project ExampleApplication, i added both, docx and HTML: Test-html-page.html and Test-Template.docx. When you run the sample application, the output will land in: 
\DocXToPdfConvter\ExampleApplication\bin\Debug\netcoreapp2.1

```csharp
 string executableLocation = Path.GetDirectoryName(
                Assembly.GetExecutingAssembly().Location);

            //Here are the 2 test files as input. They contain placeholders
            string docxLocation = Path.Combine(executableLocation, "Test-Template.docx");
            string htmlLocation = Path.Combine(executableLocation, "Test-HTML-page.html");
```            

4. In this repository, there are 2 projects: the library itself and the ExampleApplication. Have a look at program.cs on how to implement the library. Here are the steps:

a) Add the path to your soffice.exe from LibreOffice, e. g.:

```csharp
string locationOfLibreOfficeSoffice =
                @"C:\PortableApps\LibreOfficePortable\App\libreoffice\program\soffice.exe";
```


b) Define your __placeholders__, which you want to use either in your Word document or in a HTML document. There are 3 types of placeholders: one for plain text, one for table rows and one for images. A placeholder consists of a start tag, a string for the placeholder and an end tag. For instance, in your Word document or HTML document, you can place "##ThisIsAPlaceholder## - then the start and end tags are "##" and the "ThisIsAPlaceholder" is the string. To define placeholders, you have to create an object of the Placeholders class. You can customize the placeholder start and end tags and the "NewLineTag" (only for docx documents - for HTML, you just use the standard &lt;br/&gt;). If you don't define them, the following standard placeholders will be used. A start and an end tag do not have to be the same, they can also differ. __Important__: Different placeholder types (text, table row and images) MUST have different start/end tags!

So, here we first define the NewLineTag (only for Word documents) and the start/end tags. If you want to use these exactly, you don't have to define them, they are autocreated by the constructor.

```csharp
var placeholders = new Placeholders();
placeholders.NewLineTag = "<br/>";
placeholders.TextPlaceholderStartTag = "##";
placeholders.TextPlaceholderEndTag = "##";
placeholders.TablePlaceholderStartTag = "==";
placeholders.TablePlaceholderEndTag = "==";
placeholders.ImagePlaceholderStartTag = "++";
placeholders.ImagePlaceholderEndTag = "++";

```

Now, let's make placeholders for texts:

```csharp
placeholders.TextPlaceholders = new Dictionary<string, string>
            {
                {"Name", "Mr. Miller" },
                {"Street", "89 Brook St" },
                {"City", "Brookline MA 02115<br/>USA" },
                {"InvoiceNo", "5" },
                {"Total", "U$ 4,500" },
                {"Date", "28 Jul 2019" }
            };
```

A cool feature is creating placeholders for table rows. The rows of the template will be multiplied according to the number of placeholders in an array. In this example, we have 2 different table rows (actually, 2 different tables), which are automatically duplicated and filled:

```csharp
  placeholders.TablePlaceholders = new List<Dictionary<string, string[]>>
  {

          new Dictionary<string, string[]>()
          {
              {"Name", new string[]{ "Homer Simpson", "Mr. Burns", "Mr. Smithers" }},
              {"Department", new string[]{ "Power Plant", "Administration", "Administration" }},
              {"Responsibility", new string[]{ "Oversight", "CEO", "Assistant" }},
              {"Telephone number", new string[]{ "888-234-2353", "888-295-8383", "888-848-2803" }}
          },
          new Dictionary<string, string[]>()
          {
              {"Qty", new string[]{ "2", "5", "7" }},
              {"Product", new string[]{ "Software development", "Customization", "Travel expenses" }},
              {"Price", new string[]{ "U$ 2,000", "U$ 1,000", "U$ 1,500" }},
          }

  };
```


Of course, you can also add images of many formats (jpg/jpeg, png is supported and a couple of others, too) into placeholders. Here is an example:


```csharp
var productImage =
                StreamHandler.GetFileAsMemoryStream(Path.Combine(executableLocation, "ProductImage.jpg"));

var qrImage =
    StreamHandler.GetFileAsMemoryStream(Path.Combine(executableLocation, "QRCode.PNG"));

placeholders.ImagePlaceholders = new Dictionary<string, MemoryStream>
{
    {"QRCode", qrImage },
    {"ProductImage", productImage }
};
```

c) As we now have everything setup, we can now start the conversion process(es). While converting, the placeholders are filled with the values. Of course docx to docx and html to html aren't really conversions, but I also added them to the functionality of the library, because that may be useful for some people.

```csharp
var test = new ReportGenerator(locationOfLibreOfficeSoffice);

//Convert from HTML to HTML
test.Convert(htmlLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Test-HTML-page-out.html"), placeholders);

//Convert from HTML to PDF
test.Convert(htmlLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Test-HTML-page-out.pdf"), placeholders);

//Convert from HTML to DOCX
test.Convert(htmlLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Test-HTML-page-out.docx"), placeholders);

//Convert from DOCX to DOCX
test.Convert(docxLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Test-Template-out.docx"), placeholders);

//Convert from DOCX to HTML
test.Convert(docxLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Test-Template-out.html"), placeholders);

//Convert from DOCX to PDF
test.Convert(docxLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Test-Template-out.pdf"), placeholders);
         
```
