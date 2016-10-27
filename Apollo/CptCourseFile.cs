using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;


namespace Apollo {

  abstract public class CptCourseFile {
    public string FilePath { get; set; }
    public string Title { get; set; }
    virtual public void OpenFile() { }
    virtual public void CloseFile() { }
    virtual public void AddDescription(CptCourseDescription description) { }
    virtual public void SaveIntructorMaterials() { }

    abstract public void WriteContent(DocWriter docWriter);
    static public CptCourseFile Create(string FilePath) {

      if (FilePath.EndsWith("pptx", StringComparison.CurrentCultureIgnoreCase)) {
        return new CptModule(FilePath);
      }

      if ((FilePath.EndsWith("docx", StringComparison.CurrentCultureIgnoreCase) &&
        (FilePath.Contains("AboutCourse")))) {
        string prefix = "AboutCourse_";
        int index = FilePath.IndexOf(prefix) + prefix.Length;
        string CourseCode = FilePath.Substring(index).Replace(".docx", "");
        if (CourseCode.ToLower().Equals(CptCourse.Current.courseInfo.CourseCode.ToLower())) {
          return new CptAboutCoursePage(FilePath);
        }
        else {
          return null;
        }

      }


      if ((FilePath.EndsWith("docx", StringComparison.CurrentCultureIgnoreCase) &&
           (FilePath.Contains("00")))) {
        return new CptIntroPage(FilePath);

      }

      if ((FilePath.EndsWith("docx", StringComparison.CurrentCultureIgnoreCase) &&
           (FilePath.Contains("Appendix")))) {
        return new CptAppendix(FilePath);
      }

      if ((FilePath.EndsWith("docx", StringComparison.CurrentCultureIgnoreCase) &&
           (FilePath.ToLower().Contains("backcoverpage")))) {
        return new CptBackPage(FilePath);
      }

      if (FilePath.EndsWith("docx", StringComparison.CurrentCultureIgnoreCase)) {
        return new CptLab(FilePath);
      }

      return null;
    }
  }

  public class CptModule : CptCourseFile {

    private PowerPoint.Presentation presentation;

    private int ModuleNumber;
    private static int ModuleCount = 0;
    public static string GetCurrentModuleNumber() {
      return ModuleCount.ToString("00");
    }
    public static void Reset() { ModuleCount = 0; }

    public CptModule(string FilePath) {
      ModuleCount += 1;
      ModuleNumber = ModuleCount;
      this.FilePath = FilePath;
    }

    public bool IsTopicSlide(PowerPoint.Slide slide) {
      // return true if slide has no title
      if (slide.Shapes.HasTitle == MsoTriState.msoFalse) {
        return true;
      }
      // return false for (1) intro slide, (2) agenda slides, (3) summary slide and (4) demo slides
      string SlideTitle = slide.Shapes.Title.TextFrame.TextRange.Text;
      if ((slide.SlideIndex == 1) ||
          (SlideTitle.ToLower().Equals("agenda")) ||
          (SlideTitle.ToLower().Equals("summary")) ||
          (SlideTitle.ToLower().Equals("demo")) ||
          (slide.CustomLayout.Name.ToLower() == "demo layout")) { // added this line to look for slide layout
        return false;
      }
      else {
        return true;
      }

    }

    public override void OpenFile() {
      BuildEnv.ActivatePowerPoint();

      MsoTriState ShowWindow;

      if (BuildEnv.UIRefreshingEnabled) {
        ShowWindow = MsoTriState.msoTrue;
      }
      else {
        ShowWindow = MsoTriState.msoFalse;
      }

      presentation = BuildEnv.PowerPointApplication.Presentations.Open(FilePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

      Title = presentation.BuiltInDocumentProperties.Item["Title"].Value;
      BuildEnv.RefreshView(presentation);
      Console.WriteLine("Processing Module " + ModuleNumber.ToString("00") + ": " + Title);
      try {
        BuildEnv.PowerPointApplication.Visible = MsoTriState.msoFalse;
      }
      catch { }
      BuildEnv.PowerPointApplication.WindowState = PowerPoint.PpWindowState.ppWindowMinimized;

    }

    override public void AddDescription(CptCourseDescription description) {
      CptCourseModule module = new CptCourseModule();
      module.Number = this.ModuleNumber.ToString("00");
      module.Title = this.Title;



      PowerPoint.Slide introSlide = presentation.Slides[1];
      // insert course description
      if (introSlide.HasNotesPage == MsoTriState.msoTrue) {
        if (introSlide.NotesPage.Shapes[2].TextFrame2.HasText == MsoTriState.msoTrue) {
          TextRange2 range = introSlide.NotesPage.Shapes[2].TextFrame2.TextRange;
          module.Description = range.Text;
        }
      }

      List<string> agendaItems = GetAgendaItems(presentation);
      foreach (string item in agendaItems) {
        module.AgendaTopics.Add(item);
      }

      description.Modules.Add(module);
    }

    public override void WriteContent(DocWriter docWriter) {
      docWriter.AddModuleSection("Module " + ModuleNumber.ToString("00") + ": " + Title);
      InsertModuleIntro(presentation, docWriter);
      string BuildFolder = BuildEnv.BuildFolder + @"/" + CptCourse.Current.courseInfo.CourseCode + @"/Module" + this.ModuleNumber.ToString();
      Directory.CreateDirectory(BuildFolder);
      Console.WriteLine(" - saving slides as PNG files...");
      presentation.SaveAs(BuildFolder, PowerPoint.PpSaveAsFileType.ppSaveAsPNG);

     // BuildEnv.ActivateWord();
      // begin enumeration through slides
      Console.WriteLine(" - adding slides to Word doc...");
      foreach (PowerPoint.Slide slide in presentation.Slides) {
        if ((IsTopicSlide(slide)) && (slide.Shapes.HasTitle == MsoTriState.msoTrue)) {
          string SlideTitle = slide.Shapes.Title.TextFrame.TextRange.Text;
          string pngPath = BuildFolder + @"/Slide" + slide.SlideIndex.ToString() + ".png";
          docWriter.AddSlide(pngPath);
          CopyAndPasteNotes(slide, docWriter);
        }
        // end enumeration through slides      
      }

    }

    public override void SaveIntructorMaterials() {

      int SlideNotesInspector = 4;

      MsoDocInspectorStatus status;
      string result;

      presentation.DocumentInspectors[SlideNotesInspector].Fix(out status, out result);
      presentation.RemoveDocumentInformation(PowerPoint.PpRemoveDocInfoType.ppRDIComments);
      presentation.RemoveDocumentInformation(PowerPoint.PpRemoveDocInfoType.ppRDIDocumentProperties);
      presentation.RemoveDocumentInformation(PowerPoint.PpRemoveDocInfoType.ppRDIRemovePersonalInformation);

      presentation.BuiltInDocumentProperties.Item["Title"].Value = this.Title;
      presentation.BuiltInDocumentProperties.Item["Author"].Value = "Critical Path Training";
      presentation.BuiltInDocumentProperties.Item["Company"].Value = "Critical Path Training, LLC";

      if (presentation.HasNotesMaster)
        presentation.NotesMaster.HeadersFooters.Clear();
      if (presentation.HasHandoutMaster)
        presentation.HandoutMaster.HeadersFooters.Clear();
      if (presentation.HasTitleMaster == MsoTriState.msoTrue)
        presentation.TitleMaster.HeadersFooters.Clear();

      string FileName = this.FilePath.Substring(this.FilePath.LastIndexOf(" ") + 1);
   
      presentation.SaveAs(BuildEnv.InstructorSlidesFolder + @"\" + ModuleNumber.ToString("00") + "-" + FileName);

    }

    public override void CloseFile() {
      presentation.Close();
    }

    List<string> GetAgendaItems(PowerPoint.Presentation presentation) {
      List<string> agenda = new List<string>();
      foreach (PowerPoint.Slide slide in presentation.Slides) {
        if (slide.Shapes.HasTitle == MsoTriState.msoTrue) {
          if (slide.Shapes.Title.TextFrame.TextRange.Text.ToLower().Contains("agenda")) {
            TextRange2 agendaRange = slide.Shapes[2].TextFrame2.TextRange;
            string agendaString = agendaRange.Text;
            char[] splitters = { '\r', '\v', '\t' };
            string[] agendaArray = agendaString.Split(splitters);
            foreach (string agendaItem in agendaArray) {
              agenda.Add(agendaItem);
            }
            return agenda;
          }
        }
      }
      return agenda;
    }

    List<string> GetTopicsCovered(PowerPoint.Presentation presentation) {
      List<string> topics = new List<string>();
      foreach (PowerPoint.Slide slide in presentation.Slides) {
        if (slide.Shapes.HasTitle == MsoTriState.msoTrue) {
          if ((IsTopicSlide(slide)) && (slide.Shapes.Title != null)) {
            string SlideTitle = slide.Shapes.Title.TextFrame.TextRange.Text;
            topics.Add(SlideTitle);
          }
        }
      }
      return topics;
    }

    List<string> GetDemos(PowerPoint.Presentation presentation) {
      List<string> demos = new List<string>();
      foreach (PowerPoint.Slide slide in presentation.Slides) {
        if (slide.Shapes.HasTitle == MsoTriState.msoTrue) {
          if (slide.CustomLayout.Name.ToLower() == "demo layout") {
            string demo = slide.Shapes.Title.TextFrame2.TextRange.Text;
            demo = demo.Replace("\r", " ").Replace("\t", " ").Replace("\v", " ");
            demos.Add(demo);
          }
        }
      }
      return demos;
    }

    public void InsertModuleIntro(PowerPoint.Presentation presentation, DocWriter docWriter) {
      PowerPoint.Slide introSlide = presentation.Slides[1];

      docWriter.AddParagraph("Module Description", "Module Intro Header");

      // insert course description
      CopyAndPasteNotesModuleDescription(introSlide, docWriter);

      List<string> agendaItems = GetAgendaItems(presentation);
      if (agendaItems.Count > 0) {
        docWriter.AddParagraph("Module Agenda", "Module Intro Header");
        docWriter.AddList(agendaItems, "Module Agenda Item");
      }

      // insert topics covered
      docWriter.AddParagraph("Topics Covered", "Module Intro Header");
      List<string> topics = GetTopicsCovered(presentation);
      docWriter.AddTopicsList(topics);


      List<string> demos = GetDemos(presentation);

      if (demos.Count > 0) {
        // insert topics covered
        docWriter.AddParagraph("Instructor Demos", "Module Intro Header");
        docWriter.AddList(demos, "Lab Exercise Item");

      }


    }

    public void CopyAndPasteNotesModuleDescription(PowerPoint.Slide slide, DocWriter docWriter) {
      if (slide.HasNotesPage == MsoTriState.msoTrue) {
        if (slide.NotesPage.Shapes[2].TextFrame2.HasText == MsoTriState.msoTrue) {
          TextRange2 range = slide.NotesPage.Shapes[2].TextFrame2.TextRange;
          docWriter.AddModuleDescription(range);
        }
        else { docWriter.AddModuleDescription(); }
      }
      else { docWriter.AddModuleDescription(); }

    }

    public void CopyAndPasteNotes(PowerPoint.Slide slide, DocWriter docWriter) {
      if (slide.HasNotesPage == MsoTriState.msoTrue) {
        if (slide.NotesPage.Shapes[2].TextFrame2.HasText == MsoTriState.msoTrue) {
          TextRange2 range = slide.NotesPage.Shapes[2].TextFrame2.TextRange;
          docWriter.AddNotes(range);
        }
        else { docWriter.AddNotes(); }
      }
      else { docWriter.AddNotes(); }
    }

  }

  public class CptAboutCoursePage : CptCourseFile {

    Word.Document document;

    public override void OpenFile() {
      BuildEnv.ActivateWord();
      document = BuildEnv.WordApplication.Documents.Open(FilePath, Visible: MsoTriState.msoFalse);
      BuildEnv.RefreshView(document);
    }

    public override void CloseFile() {
      ((Word._Document)document).Close();
    }

    override public void AddDescription(CptCourseDescription description) {

      foreach (Word.ContentControl control in document.ContentControls) {

        if (control.Tag.Equals("Audience")) {
          description.Audience = control.Range.Text.Replace("\r", "").Replace("\t", "");
        }

        if (control.Tag.Equals("Format")) {
          description.Format = control.Range.Text.Replace("\r", "").Replace("\t", "");
        }

        if (control.Tag.Equals("Length")) {
          description.Length = control.Range.Text.Replace("\r", "").Replace("\t", "");
        }

        if (control.Tag.Equals("CourseDescription")) {
          List<string> DescriptionParagraphs = new List<string>();
          foreach (Word.Paragraph p in control.Range.Paragraphs) {
            DescriptionParagraphs.Add(p.Range.Text.Replace("\r", "").Replace("\t", ""));
          }
          description.Description = DescriptionParagraphs;
        }

        if (control.Tag.Equals("StudentPrerequisites")) {
          List<string> PrerequisitesParagraphs = new List<string>();
          foreach (Word.Paragraph p in control.Range.Paragraphs) {
            PrerequisitesParagraphs.Add(p.Range.Text.Replace("\r", "").Replace("\t", ""));
          }
          description.Prerequisites = PrerequisitesParagraphs;
        }
        // ensure About page has correct title and course code
        if (control.Tag.Equals("CourseTitle")) {
          control.Range.Text = description.CourseTitle;
        }

        if (control.Tag.Equals("CourseCoude")) {
          control.Range.Text = description.CourseCode;
        }
        // save about page with updated course title and course code
        try {
          document.Save();
        }
        catch { }

      }
    }

    public CptAboutCoursePage(string FilePath) {
      this.FilePath = FilePath;
    }
    public override void WriteContent(DocWriter docWriter) {
      docWriter.AddIntroPage(FilePath);
    }
  }

  public class CptIntroPage : CptCourseFile {

    public CptIntroPage(string FilePath) {
      this.FilePath = FilePath;
    }
    public override void WriteContent(DocWriter docWriter) {
      docWriter.AddIntroPage(FilePath);
    }
  }

  public class CptLab : CptCourseFile {

    private string ModuleNumber;
    Word.Document document;

    public CptLab(string FilePath) {
      this.FilePath = FilePath;
      ModuleNumber = CptModule.GetCurrentModuleNumber();
    }

    public override void OpenFile() {
      BuildEnv.ActivateWord();
      document = BuildEnv.WordApplication.Documents.Open(FilePath, Visible: MsoTriState.msoFalse);
      BuildEnv.RefreshView(document);
    }

    public override void CloseFile() {
      try { ((Word._Document)document).Close(); }
      catch { }
    }

    override public void AddDescription(CptCourseDescription description) {
      CptCourseLab lab = new CptCourseLab();
      lab.Title = document.Paragraphs.First.Range.Text.Replace("\r", " ").Replace("\t", " ").Replace("\v", " ");
      foreach (Word.Paragraph p in document.Paragraphs) {
        Word.Style style = p.get_Style();
        if (style.NameLocal.Equals("Heading 3")) {
          lab.Exercises.Add(p.Range.Text.Replace("\r", " ").Replace("\t", " ").Replace("\v", " "));
        }
      }
      description.Modules.Last().Labs.Add(lab);
    }


    public override void WriteContent(DocWriter docWriter) {
      docWriter.AddLab(FilePath, ModuleNumber);
    }
  }

  public class CptAppendix : CptCourseFile {

    private int AppendixNumber;
    private static int AppendixCount = 0;
    public static void Reset() { AppendixCount = 0; }

    public CptAppendix(string FilePath) {
      AppendixCount += 1;
      AppendixNumber = AppendixCount;
      this.FilePath = FilePath;
    }

    public override void WriteContent(DocWriter docWriter) {
      docWriter.AddAppendix(FilePath, AppendixNumber);
    }
  }

  public class CptBackPage : CptCourseFile {

    public CptBackPage(string FilePath) {
      this.FilePath = FilePath;
    }
    public override void WriteContent(DocWriter docWriter) {
      docWriter.AddBackPage(FilePath);
    }
  }
}
