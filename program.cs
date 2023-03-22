Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
Workbook workbook = excel.Workbooks.Add();
Worksheet worksheet = workbook.ActiveSheet;
string text = webBrowser1.Document.Body.InnerText;
worksheet.Cells[1, 1] = metin;
workbook.SaveAs("C:\\text.xlsx");
workbook.Close();
excel.Quit();

Microsoft.Office.Interop.PowerPoint.Application powerpoint = new Microsoft.Office.Interop.PowerPoint.Application();
Presentation presentation = powerpoint.Presentations.Add();
Slide slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
HtmlElementCollection elements = webBrowser1.Document.GetElementsByTagName("img");
Image image = null;
 foreach (HtmlElement element in elements)
{
  if (element.TagName.ToLower() == "img")
  {
    string imageUrl = element.GetAttribute("src");
    image = Image.FromFile(imageUrl);
     break;
   }
 }
slide.Background.Fill.UserPicture(image.ToString());
presentation.SaveAs("C:\\presentation.pptx");
presentation.Close();
powerpoint.Quit();
Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
Document document = word.Documents.Add();
string text = webBrowser1.Document.Body.InnerText;
Microsoft.Office.Interop.Word.Range range = document.Range();
range.Text = text;
object filename = @"C:\document.docx";
document.SaveAs2(ref filename);
document.Close();
word.Quit();
