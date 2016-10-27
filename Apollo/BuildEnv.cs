using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace Apollo {
  public class BuildEnv {

    private const string InstructorSlidesFolderName = @"\InstructorSlides";
    private const string CourseDescriptionsFolderName = @"\CourseDescriptions";
    private const string CourseSellsheetsFolderName = @"\CourseSellsheets";

    public static string BuildFolder { get; set; }
    public static string OutputFolder { get; set; }
    public static string InstructorSlidesFolder { get; set; }
    public static string CourseDescriptionsFolder { get; set; }
    public static string CourseSellsheetsFolder { get; set; }

    public static string OutputFileName { get; set; }
    public static string VersionNumber { get; set; }
    public static string OutputFileNameWithVersion { get; set; }

    public static string StyleTemplateFilePath { get; set; }
    public static string SellsheetStyleTemplateFilePath { get; set; }
    public static string CourseManualDocxFilePath { get; set; }
    public static string CourseManualPdfFilePath { get; set; }
    public static string CourseDescriptionXmlFilePath { get; set; }
    public static string CourseDescriptionJsonFilePath { get; set; }

    public static bool BuildManaual { get; set; }
    public static bool UIRefreshingEnabled { get; set; }

    public static void Initialize(CptCourseInfo courseInfo, bool BuildManaual, bool RefreshUI) {

      BuildEnv.BuildManaual = BuildManaual;
      BuildEnv.UIRefreshingEnabled = RefreshUI;

      // reset number for next course in build set
      CptModule.Reset();
      CptAppendix.Reset();

      OutputFileName = courseInfo.OutputFileName;
      VersionNumber = courseInfo.Version;
      OutputFileNameWithVersion = OutputFileName + "_v" + courseInfo.Version;


      // ensure build folder exists
      BuildFolder = courseInfo.BuildFolder;
      if (string.IsNullOrEmpty(BuildFolder)) {
        BuildFolder = Directory.GetCurrentDirectory() + @"\build";
      }      
      Directory.CreateDirectory(BuildFolder);

      // ensure output folder exists
      OutputFolder = courseInfo.OutputFolder;
      if (string.IsNullOrEmpty(OutputFolder)) {
        OutputFolder = Directory.GetCurrentDirectory() + @"\output";
      }
      Directory.CreateDirectory(OutputFolder);
      Directory.CreateDirectory(OutputFolder + @"\" + courseInfo.CourseCode + "_v" + courseInfo.Version);
      
      CourseManualDocxFilePath = OutputFolder + @"\" + courseInfo.CourseCode + "_v" + courseInfo.Version + @"\" + OutputFileNameWithVersion + ".docx";
      CourseManualPdfFilePath = OutputFolder + @"\" + courseInfo.CourseCode + "_v" + courseInfo.Version + @"\" + OutputFileNameWithVersion + ".pdf";

      CourseDescriptionsFolder = OutputFolder + CourseDescriptionsFolderName;
      CourseDescriptionXmlFilePath = CourseDescriptionsFolder + @"\" + OutputFileNameWithVersion + ".xml";
      CourseDescriptionJsonFilePath = CourseDescriptionsFolder + @"\" + OutputFileNameWithVersion + ".json";
      Directory.CreateDirectory(CourseDescriptionsFolder);

      CourseSellsheetsFolder = courseInfo.OutputFolder + CourseSellsheetsFolderName;
      Directory.CreateDirectory(CourseSellsheetsFolder);

      // ensure InstructorSlidesFolder
      Directory.CreateDirectory(OutputFolder + InstructorSlidesFolderName);
      Directory.CreateDirectory(OutputFolder + InstructorSlidesFolderName + @"\" + courseInfo.CourseCode + "_v" + courseInfo.Version);

      InstructorSlidesFolder = OutputFolder +
                               @"\" + courseInfo.CourseCode + "_v" + courseInfo.Version + @"\" +
                               OutputFileName + "_InstructorSlides_v" + courseInfo.Version + @"\"; ;

      Directory.CreateDirectory(InstructorSlidesFolder);

      if (BuildEnv.BuildManaual) { 
        StyleTemplateFilePath = BuildFolder + @"\CptCourseManual.dotx";
        FileStream stream = new FileStream(StyleTemplateFilePath, FileMode.Create);
        byte[] templateFile = Properties.Resources.CptCourseManual_dotx;
        stream.Write(templateFile, 0, templateFile.Length);
        stream.Close();
      }

      SellsheetStyleTemplateFilePath = BuildFolder + @"\CptSellsheet.dotx";
      FileStream streamSellsheet = new FileStream(SellsheetStyleTemplateFilePath, FileMode.Create);
      byte[] templateFileSellsheet = Properties.Resources.CptSellsheet_dotx;
      streamSellsheet.Write(templateFileSellsheet, 0, templateFileSellsheet.Length);
      streamSellsheet.Close();


      FileStream streamCoverHeader = new FileStream(BuildFolder + @"\ApolloCoverHeader.png", FileMode.Create);
      byte[] fileCoverHeader = Properties.Resources.ApolloCoverHeader_png;
      streamCoverHeader.Write(fileCoverHeader, 0, fileCoverHeader.Length);
      streamCoverHeader.Close();

      FileStream streamCoverFooter = new FileStream(BuildFolder + @"\ApolloCoverFooter.png", FileMode.Create);
      byte[] fileCoverFooter = Properties.Resources.ApolloCoverFooter_png;
      streamCoverFooter.Write(fileCoverFooter, 0, fileCoverFooter.Length);
      streamCoverFooter.Close();

      FileStream streamTitleHeader = new FileStream(BuildFolder + @"\TitleHeaderImage.gif", FileMode.Create);
      byte[] fileTitleHeader = Properties.Resources.TitleHeaderImage_gif;
      streamTitleHeader.Write(fileTitleHeader, 0, fileTitleHeader.Length);
      streamTitleHeader.Close();

      FileStream streamSectionHeader = new FileStream(BuildFolder + @"\SectionHeaderImage.gif", FileMode.Create);
      byte[] fileSectionHeader = Properties.Resources.SectionHeaderImage_gif;
      streamSectionHeader.Write(fileSectionHeader, 0, fileSectionHeader.Length);
      streamSectionHeader.Close();

      FileStream streamApolloCover = new FileStream(BuildFolder + @"\ApolloCover.jpg", FileMode.Create);
      byte[] fileApolloCover = Properties.Resources.ApolloCover_jpg;
      streamApolloCover.Write(fileApolloCover, 0, fileApolloCover.Length);
      streamApolloCover.Close();

      FileStream streamSellsheetBanner= new FileStream(BuildFolder + @"\CptCourseDescriptionHeader.png", FileMode.Create);
      byte[] fileSellsheetBanner = Properties.Resources.CptCourseDescriptionHeader_png;
      streamSellsheetBanner.Write(fileSellsheetBanner, 0, fileSellsheetBanner.Length);
      streamSellsheetBanner.Close();

      StreamWriter streamCourseDescriptionXsl = new StreamWriter(CourseDescriptionsFolder + @"\CptCourseDescription.xsl", false, Encoding.UTF8);
      streamCourseDescriptionXsl.Write(Properties.Resources.CptCourseDescription_xsl);
      streamCourseDescriptionXsl.Close();

      StreamWriter streamCourseDescriptionCss = new StreamWriter(CourseDescriptionsFolder + @"\CptCourseDescription.css", false, Encoding.UTF8);
      streamCourseDescriptionCss.Write(Properties.Resources.CptCourseDescription_css);
      streamCourseDescriptionCss.Close();

    }


    private static PowerPoint.Application ppApp;
    public static PowerPoint.Application PowerPointApplication {
      get {
        if (ppApp == null) {
          ppApp = new PowerPoint.Application();
          try {
            ppApp.Visible = MsoTriState.msoFalse;
          }
          catch { }
          ppApp.WindowState = PowerPoint.PpWindowState.ppWindowMinimized;
          
        }
        return ppApp;
      }
    }

    public static void ActivatePowerPoint() {
      if (UIRefreshingEnabled && BuildEnv.BuildManaual) {
        PowerPointApplication.Activate();
      }
    }

    private static Word.Application wordApp;
    public static Word.Application WordApplication {
      get {
        if (wordApp == null) {
          wordApp = new Word.Application();
          wordApp.Visible = UIRefreshingEnabled;
        }
        return wordApp;
      }
    }

    public static void ActivateWord() {
      if (UIRefreshingEnabled && BuildEnv.BuildManaual) {
        try {
          WordApplication.Activate();
        }
        catch { }
      }
    }

    public static void RefreshView(Word.Document document) {
      if (UIRefreshingEnabled && BuildEnv.BuildManaual) {
        try {
          document.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
          document.ActiveWindow.ActivePane.View.Zoom.PageColumns = 2;
          document.ActiveWindow.ActivePane.View.Zoom.PageRows = 1;
        }
        catch { }
      }
    }

    public static void RefreshView(PowerPoint.Presentation presentation) {
      if (UIRefreshingEnabled && BuildEnv.BuildManaual) {
        try {
          presentation.Application.ActiveWindow.ViewType = PowerPoint.PpViewType.ppViewSlideSorter;
          PowerPoint.Pane pane = presentation.Application.ActiveWindow.Panes[1];
          pane.Activate();
          presentation.Application.ActiveWindow.View.Zoom = 100;
        }
        catch { }
      }
    }

    public static void ScollDocument(Word.Range range) {
      if (UIRefreshingEnabled) {
        try { WordApplication.ActiveWindow.ScrollIntoView(range); }
        catch { }
      }
    }

    public static void QuitWord() {
      try {
        ((Word._Application)wordApp).Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
      }
      catch { }
    }

    public static void QuitPowerPoint() {
      try { PowerPointApplication.Quit(); }
      catch { }
    }

  }
}
