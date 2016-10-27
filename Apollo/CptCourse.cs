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
using Newtonsoft.Json;

namespace Apollo {

  public class CptCourse {

    public static CptCourse Current { get; private set; }

    public CptCourseInfo courseInfo;
    private CptCourseDescription courseDescription;
    public CptCourseDescription CourseDescription {
      get {
        return courseDescription;
      }
    }
    
    public CptCourse(CptCourseInfo courseInfo) {
      this.courseInfo = courseInfo;
      this.courseDescription = new CptCourseDescription();
      this.courseDescription.CourseTitle = this.courseInfo.CourseTitle;
      this.courseDescription.CourseCode = this.courseInfo.CourseCode;
      this.courseDescription.CourseSubtitle = this.courseInfo.CourseSubtitle;
      this.courseDescription.Version = this.courseInfo.Version;
      CptCourse.Current = this;
    }

    private List<CptCourseFile> courseFiles = new List<CptCourseFile>();
  
    private void LoadCourseFiles() {
      string[] omittedModules = courseInfo.OmittedModules.Split(';');
      string dir = courseInfo.SourceDirectory;
      foreach (string file in Directory.GetFiles(dir)) {
        if (!file.Contains("~")) {
          if ((file.EndsWith(@".pptx", StringComparison.CurrentCultureIgnoreCase)) ||
             (file.EndsWith(@".docx", StringComparison.CurrentCultureIgnoreCase))) {
            int FileNameStartPosition = file.LastIndexOf(@"\") + 1;
            string FileNameNumber = file.Substring(FileNameStartPosition, 2);
            if (!omittedModules.Contains(FileNameNumber)) {
              CptCourseFile courseFile = CptCourseFile.Create(file);
              if (courseFile != null) {
                courseFiles.Add(courseFile);
              }
            }
          }
        }
      }

    }

    public void CreateOutput() {

      Word.Document manual = null;
      DocWriter docWriter = null;

      if (BuildEnv.BuildManaual) {
        manual = BuildEnv.WordApplication.Documents.Add();
        manual.ShowGrammaticalErrors = false;
        manual.ShowSpellingErrors = false;
        BuildEnv.RefreshView(manual);
        docWriter = new DocWriter(manual, this.courseInfo);
      }

      courseDescription =
        new CptCourseDescription {
          Version = courseInfo.Version,
          CourseTitle = courseInfo.CourseTitle,
          CourseSubtitle = courseInfo.CourseSubtitle,
          CourseCode = courseInfo.CourseCode
        };

      LoadCourseFiles();
      
      foreach (var courseFile in courseFiles) {
        courseFile.OpenFile();
        courseFile.AddDescription(courseDescription);
        if (BuildEnv.BuildManaual) {
          courseFile.WriteContent(docWriter);
          courseFile.SaveIntructorMaterials();
        }
        courseFile.CloseFile();
      }

      Console.WriteLine(" - saving course decription as XML file...");
      SaveCourseDescription();

      Console.WriteLine(" - saving course sellsheet as Word doc...");
      CreateSellSheet();

      if (BuildEnv.BuildManaual) {
        docWriter.UpdateToc();
        docWriter.RemoveAllComments();
        Console.WriteLine(" - saving student manual as Word doc...");  
        manual.SaveAs2(BuildEnv.CourseManualDocxFilePath);
        if (courseInfo.BuildPDF) {
          Console.WriteLine(" - saving student manual as PDF file...");
          manual.SaveAs(BuildEnv.CourseManualPdfFilePath, Word.WdSaveFormat.wdFormatPDF);
        }

        ((Word._Document)manual).Close(Word.WdSaveOptions.wdDoNotSaveChanges);
        Console.WriteLine(" - processing course " + 
                          CptCourse.Current.CourseDescription.CourseCode + 
                          " complete");
        Console.WriteLine();
      }

    }

    private void SaveCourseDescription() {
      XmlSerializer seriaizer = new XmlSerializer(typeof(CptCourseDescription));
      FileStream stream = new FileStream(BuildEnv.CourseDescriptionXmlFilePath, FileMode.Create);
      XmlWriterSettings settings = new XmlWriterSettings();
      settings.Indent = true;
      settings.NewLineHandling = NewLineHandling.Entitize;
      settings.OmitXmlDeclaration = false;
      settings.NamespaceHandling = NamespaceHandling.OmitDuplicates;
      settings.Encoding = Encoding.UTF8;
      XmlWriter writer = XmlWriter.Create(stream, settings);
      writer.WriteProcessingInstruction("xml-stylesheet", "type='text/xsl' href='CptCourseDescription.xsl'");
      seriaizer.Serialize(writer, courseDescription);
      stream.Dispose();

      string json = JsonConvert.SerializeObject(courseDescription);
      FileStream streamJson = new FileStream(BuildEnv.CourseDescriptionJsonFilePath, FileMode.Create);
      StreamWriter writerJson = new StreamWriter(streamJson);
      writerJson.Write(json);
      writerJson.Flush();
      writerJson.Dispose();
      streamJson.Dispose();

    }

    private void ZipUpInstructorSlides() {
    }

    private void CreateSellSheet() {

      string pathSellSheet = BuildEnv.CourseSellsheetsFolder + @"\" + BuildEnv.OutputFileNameWithVersion + ".docx";
      string pathStyleSheet = BuildEnv.BuildFolder + @"\CptSellsheet.dotx";
      
      CptSellsheetWriter.CreateSellSheet(CourseDescription,pathSellSheet, pathStyleSheet);
      
    }
   
  }


}
