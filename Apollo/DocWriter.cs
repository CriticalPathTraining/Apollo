using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office;
using Microsoft.Office.Core;
using Microsoft.Office.Interop;
using Word = Microsoft.Office.Interop.Word;

namespace Apollo {

  public class DocWriter {

    private Word.Document document;
    private CptCourseInfo documentInfo;

    private Word.Range EndOfDoc {
      get { return document.Bookmarks[((object)@"\endofdoc")].Range; }
    }

    private string StyleNameLevel1 = "Lab Step Numbered";
    private string StyleNameLevel2 = "Lab Step Numbered Level 2";
    private string StyleNameLevel3 = "Lab Step Numbered Level 3";
    private string StyleNameLevel4 = "Lab Step Numbered Level 4";
    private Word.Style StyleLevel1;
    private Word.Style StyleLevel2;
    private Word.Style StyleLevel3;
    private Word.Style StyleLevel4;

    private string LabStepsListTemplateName = "LabStepsTemplate";
    private Word.ListTemplate LabStepsListTemplate;
  

    public DocWriter(Word.Document doc, CptCourseInfo docInfo) {
      if (doc == null) { throw new ApplicationException("document cannot be null"); }
      this.document = doc;
      documentInfo = docInfo;
      document.CopyStylesFromTemplate(BuildEnv.StyleTemplateFilePath);

      StyleLevel1 = document.Styles[StyleNameLevel1];
      StyleLevel2 = document.Styles[StyleNameLevel2];
      StyleLevel3 = document.Styles[StyleNameLevel3];
      StyleLevel4 = document.Styles[StyleNameLevel4];

      foreach (Word.ListTemplate temp in document.ListTemplates) {
        if (temp.Name.Equals(LabStepsListTemplateName)) {
          LabStepsListTemplate = document.ListTemplates[LabStepsListTemplateName];
        }
      }

      if (LabStepsListTemplate == null) {
        LabStepsListTemplate = document.ListTemplates.Add(true, LabStepsListTemplateName);
      }

       
      LabStepsListTemplate.ListLevels[1].LinkedStyle = StyleNameLevel1;
      //LabStepsListTemplate.ListLevels[1].Font.Size = 9;
      LabStepsListTemplate.ListLevels[1].NumberFormat = "%1.";
      //LabStepsListTemplate.ListLevels[1].NumberPosition = 0;
      LabStepsListTemplate.ListLevels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
      LabStepsListTemplate.ListLevels[1].ResetOnHigher = 0;
      LabStepsListTemplate.ListLevels[1].StartAt = 1;
      //LabStepsListTemplate.ListLevels[1].TabPosition = 18;
      //LabStepsListTemplate.ListLevels[1].TextPosition = 18;
      LabStepsListTemplate.ListLevels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;

      LabStepsListTemplate.ListLevels[2].LinkedStyle = StyleNameLevel2;
      //LabStepsListTemplate.ListLevels[2].Font.Size = 9;
      //LabStepsListTemplate.ListLevels[2].NumberFormat = "%2)";
      //LabStepsListTemplate.ListLevels[2].NumberPosition = 20;
      //LabStepsListTemplate.ListLevels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLowercaseLetter;
      LabStepsListTemplate.ListLevels[2].ResetOnHigher = 1;
      LabStepsListTemplate.ListLevels[2].StartAt = 1;
      //LabStepsListTemplate.ListLevels[2].TabPosition = 36;
      //LabStepsListTemplate.ListLevels[2].TextPosition = 36;
      LabStepsListTemplate.ListLevels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;

      LabStepsListTemplate.ListLevels[3].LinkedStyle = StyleNameLevel3;
      //LabStepsListTemplate.ListLevels[3].Font.Size = 9;
      LabStepsListTemplate.ListLevels[3].NumberFormat = "%3.";
      //LabStepsListTemplate.ListLevels[3].NumberPosition = 20;
      LabStepsListTemplate.ListLevels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLowercaseRoman;
      LabStepsListTemplate.ListLevels[3].ResetOnHigher = 2;
      LabStepsListTemplate.ListLevels[3].StartAt = 1;
      //LabStepsListTemplate.ListLevels[3].TabPosition = 46;
      //LabStepsListTemplate.ListLevels[3].TextPosition = 46;
      LabStepsListTemplate.ListLevels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;

      LabStepsListTemplate.ListLevels[4].LinkedStyle = StyleNameLevel4;
      //LabStepsListTemplate.ListLevels[4].Font.Size = 9;
      LabStepsListTemplate.ListLevels[4].NumberFormat = "(%4)";
      //LabStepsListTemplate.ListLevels[4].NumberPosition = 20;
      LabStepsListTemplate.ListLevels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
      LabStepsListTemplate.ListLevels[4].ResetOnHigher = 3;
      LabStepsListTemplate.ListLevels[4].StartAt = 1;
      //LabStepsListTemplate.ListLevels[4].TabPosition = 56;
      //LabStepsListTemplate.ListLevels[4].TextPosition = 56;
      LabStepsListTemplate.ListLevels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
      
      AddIntroSection();
    }

    private void AddIntroSection() {

      document.Content.PageSetup.DifferentFirstPageHeaderFooter = -1;
      document.Content.PageSetup.TopMargin = 36;
      document.Content.PageSetup.BottomMargin = 36;
      document.Content.PageSetup.LeftMargin = 36;
      document.Content.PageSetup.RightMargin = 36;

      BuildEnv.ActivateWord();

      Word.Section section = document.Sections[1];
      SetIntroHeaders(section);

      AddParagraph(documentInfo.CourseCode, "Course Code");
      AddParagraph(documentInfo.CourseTitle, Word.WdBuiltinStyle.wdStyleTitle);
      AddParagraph(documentInfo.CourseSubtitle, Word.WdBuiltinStyle.wdStyleSubtitle);


      EndOfDoc.InsertBreak(Word.WdBreakType.wdPageBreak);

      ScollDocument(EndOfDoc);

      AddParagraph(Properties.Resources.LegalBody1, "Legal Body");
      AddParagraph(Properties.Resources.LegalBody2, "Legal Body");
      AddParagraph(Properties.Resources.LegalBody3, "Legal Body");
      AddParagraph(Properties.Resources.LegalBody4, "Legal Body");
      AddParagraph("Course Manual Version " + documentInfo.Version);

      EndOfDoc.InsertParagraphAfter();


      EndOfDoc.InsertBreak(Word.WdBreakType.wdPageBreak);

      ScollDocument(EndOfDoc);

      Word.Paragraph toc = document.Content.Paragraphs.Add(EndOfDoc);
      document.TablesOfContents.Add(toc.Range, -1, 1, 3);
      document.Bookmarks.Add("TOC", toc.Range);
      toc.set_Style(Word.WdBuiltinStyle.wdStyleTOC1);
      toc.Range.InsertParagraphAfter();

    }

    public void AddModuleSection(string moduleTitle) {
      AddSection(moduleTitle, Word.WdSectionStart.wdSectionOddPage);
      AddParagraph(moduleTitle, Word.WdBuiltinStyle.wdStyleHeading1);
    }

    private void AddSection(string contentHeader, Word.WdSectionStart sectionStart) {

      Word.Section section = document.Sections.Add(EndOfDoc);
      section.PageSetup.SectionStart = sectionStart;

      SetHeaders(section, contentHeader);

    }

    public void AddModuleIntro(string moduleDescription) {
      AddParagraph(moduleDescription);
    }

    public void AddParagraph(string content) {
      AddParagraph(content, Word.WdBuiltinStyle.wdStyleNormal);
    }

    public void AddParagraph(string content, Word.WdBuiltinStyle style) {
      Word.Paragraph paragraph = document.Content.Paragraphs.Add(EndOfDoc);
      paragraph.Range.Text = content;
      object h1 = style;
      paragraph.set_Style(ref h1);
      paragraph.Range.InsertParagraphAfter();
    }

    public void AddParagraph(string content, string style) {
      Word.Paragraph paragraph = document.Content.Paragraphs.Add(EndOfDoc);
      paragraph.Range.Text = content;
      object h1 = style;
      paragraph.set_Style(ref h1);
      paragraph.Range.InsertParagraphAfter();
    }

    public void AddList(List<string> items) {
      foreach (string item in items) {
        Word.Paragraph paragraph = document.Content.Paragraphs.Add(EndOfDoc);
        paragraph.Range.Text = item;
        object h1 = Word.WdBuiltinStyle.wdStyleListParagraph;
        paragraph.set_Style(ref h1);
        paragraph.Range.ListFormat.ApplyBulletDefault();
        paragraph.Range.InsertParagraphAfter();
      }
    }

    public void AddList(List<string> items, string styleName) {
      foreach (string item in items) {
        Word.Paragraph paragraph = document.Content.Paragraphs.Add(EndOfDoc);
        paragraph.Range.Text = item;
        object style = styleName;
        paragraph.set_Style(ref style);
        paragraph.Range.ListFormat.ApplyBulletDefault();
        paragraph.Range.InsertParagraphAfter();
      }
    }


    public void AddTopicsList(List<string> items) {

      Word.Range rTable = EndOfDoc;
      Word.Table TopicsTable = rTable.Tables.Add(rTable, 1, 2);

      TopicsTable.Rows.LeftIndent = 0;
      TopicsTable.TopPadding = 2;
      TopicsTable.RightPadding = 4;
      TopicsTable.BottomPadding = 0;
      TopicsTable.LeftPadding = 4;

      int Counter = 0;
      foreach (string para in items) {
        Word.Paragraphs paras;
        if (Counter <= (items.Count/2)) {
          paras = TopicsTable.Rows[1].Cells[1].Range.Paragraphs;
        }
        else {
          paras = TopicsTable.Rows[1].Cells[2].Range.Paragraphs;
        }

        if (paras.Count == 1 && paras.First.Range.Text.Equals("\r\a")) {
          paras.First.Range.Text = para;
          paras.First.Range.set_Style(((object)"Topics Covered Item"));

        }
        else {
          Word.Paragraph p = paras.Add();
          p.Range.Text = para;
          p.Range.set_Style(((object)"Topics Covered Item"));

        }
        Counter += 1;
      }

      //TopicsTable.Columns[1].AutoFit();


    }
     
    public void AddAgendaSlide(string path) {
      Word.Paragraph paragraph = document.Content.Paragraphs.Add(EndOfDoc);
      ScollDocument(EndOfDoc);
      Word.InlineShape slide = paragraph.Range.InlineShapes.AddPicture(path);
      slide.LockAspectRatio = MsoTriState.msoTrue;
      slide.ScaleHeight = 40;
      slide.ScaleWidth = 40;
      paragraph.Range.set_Style(((object)"Agenda Slide"));
      paragraph.Range.InsertParagraphAfter();
      EndOfDoc.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
    }




    public void AddSlide(string path) {
      AddSlide(path, true);
    }

    public void AddSlide(string path, bool PageBreakBefore) {

      Word.Paragraph paragraph = document.Content.Paragraphs.Add(EndOfDoc);
      if (PageBreakBefore) {
        paragraph.Range.InsertBreak(Word.WdBreakType.wdPageBreak);
      }
      else {
        paragraph.Range.InsertBreak(Word.WdBreakType.wdLineBreak);
      }

      ScollDocument(EndOfDoc);

      Word.InlineShape slide = paragraph.Range.InlineShapes.AddPicture(path);
      slide.LockAspectRatio = MsoTriState.msoTrue;
      slide.ScaleHeight = 60;
      slide.ScaleWidth = 60;
      paragraph.Range.set_Style(((object)"Slide"));
      paragraph.Range.InsertParagraphAfter();
      EndOfDoc.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
    }

    public void AddCourseHeaderImage() {
      string path = documentInfo.BuildFolder + @"\TitleHeaderImage.gif";
      Word.InlineShape slide = EndOfDoc.InlineShapes.AddPicture(path);
      slide.LockAspectRatio = MsoTriState.msoTrue;
      slide.ScaleHeight = 100;
      slide.ScaleWidth = 100;
    }

    public void AddSectionHeaderImage() {
      //string path = documentInfo.BuildFolder + @"\SectionHeaderImage.gif";
      //Word.InlineShape slide = EndOfDoc.Sections.Last.Range.Paragraphs.First.Range.InlineShapes.AddPicture(path);
      //slide.LockAspectRatio = MsoTriState.msoTrue;
      //slide.ScaleHeight = 100;
      //slide.ScaleWidth = 100;
    }

    public void AddModuleDescription(TextRange2 range) {
      range.Copy();
      Word.Range notes = EndOfDoc.Paragraphs.Add().Range;
      notes.PasteAndFormat(Word.WdRecoveryType.wdUseDestinationStylesRecovery);
      notes = document.Range(notes.Start, EndOfDoc.End);

      object style = "Module Description";
      notes.set_Style(ref style);
    }

    public void AddModuleDescription() {
      Word.Range notes = EndOfDoc.Paragraphs.Add().Range;
      notes.Text = "TODO: Add Module Description. It was not found under first slide in slide deck.";
      object style = "Module Description";
      notes.set_Style(ref style);
    }
    public void AddNotes() {
      // don't add anything if slide has no notes underneath
      AddParagraph("");
    }

    public void AddNotes(TextRange2 range) {
      try {
        range.Copy();
        Word.Range notes = EndOfDoc.Paragraphs.Add().Range;
        notes.PasteAndFormat(Word.WdRecoveryType.wdUseDestinationStylesRecovery);
        notes = document.Range(notes.Start, EndOfDoc.End);

        //object style = "Slide Notes";
        //notes.set_Style(ref style);
      }
      catch { }
    }

    public void AddIntroPage(string FileName) {
      AddSection("Intro Page", Word.WdSectionStart.wdSectionOddPage);

      ScollDocument(EndOfDoc);

      EndOfDoc.InsertFile(FileName);
      string PageTitle = document.Sections.Last.Range.Paragraphs.First.Range.Text;

      //AddSectionHeaderImage();

      SetHeaders(EndOfDoc.Sections.Last, PageTitle);

    }

    public void AddLab(string FileName, string ModuleNumber) {

      Console.WriteLine(" - adding lab...");
      AddSection("Module " + ModuleNumber + " Lab", Word.WdSectionStart.wdSectionOddPage);
      ScollDocument(EndOfDoc);

      EndOfDoc.InsertFile(FileName);

      string LabPrefix = "Module " + ModuleNumber + " Lab: ";
      string LabTitle = LabPrefix + document.Sections.Last.Range.Paragraphs.First.Range.Text.Replace("\r", "").Replace("\t", "");
      document.Sections.Last.Range.Paragraphs.First.Range.InsertBefore(LabPrefix);

      ScollDocument(EndOfDoc.Sections.Last.Range.Paragraphs.First.Range);


      EnsureListNumber(EndOfDoc.Sections.Last);

      SetHeaders(EndOfDoc.Sections.Last, LabTitle);
    }

    public void AddAppendix(string FileName, int LabNumber) {
      Console.WriteLine(" - adding appendix...");
      string AppendixPrefix = "Appendix " + char.ConvertFromUtf32(64 + LabNumber).ToString() + " - ";
      AddSection(AppendixPrefix, Word.WdSectionStart.wdSectionEvenPage);

      ScollDocument(EndOfDoc);

      EndOfDoc.InsertFile(FileName);
      string AppendixTitle = AppendixPrefix + document.Sections.Last.Range.Paragraphs.First.Range.Text;
      document.Sections.Last.Range.Paragraphs.First.Range.InsertBefore(AppendixPrefix);

      AddSectionHeaderImage();

      SetHeaders(EndOfDoc.Sections.Last, AppendixTitle);
    }

    public void AddBackPage(string FileName) {
      Console.WriteLine(" - adding backpage...");
      AddSection("BackPage", Word.WdSectionStart.wdSectionEvenPage);

      ScollDocument(EndOfDoc);

      EndOfDoc.InsertFile(FileName);
      string PageTitle = document.Sections.Last.Range.Paragraphs.First.Range.Text;

      AddSectionHeaderImage();

      SetHeaders(EndOfDoc.Sections.Last, PageTitle);
    }

    public void SetIntroHeaders(Word.Section section) {

      string contentFooter = this.documentInfo.FooterText;

      section.PageSetup.DifferentFirstPageHeaderFooter = -1;
      section.PageSetup.TopMargin = 36;
      section.PageSetup.BottomMargin = 36;
      section.PageSetup.LeftMargin = 36;
      section.PageSetup.RightMargin = 36;

      string course_info = documentInfo.CourseCode + ": " + documentInfo.CourseTitle;

      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.SpaceAfter = 0;
      //section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.InlineShapes.AddPicture(documentInfo.BuildFolder + @"\ApolloCoverHeader.png");
      //section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.LineSpacing = 24;

      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = course_info + "\t\tVersion " + documentInfo.Version;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.set_Style(((object)Word.WdBuiltinStyle.wdStyleHeader));
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.SpaceAfter = 12;

      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.SpaceBefore = 0;
      //Word.InlineShape footer_image = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.InlineShapes.AddPicture(documentInfo.BuildFolder + @"\ApolloCover.jpg");
      //Word.Shape s = footer_image.ConvertToShape();

      //section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.LineSpacing = 160;

      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Fields.Add(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range, Word.WdFieldType.wdFieldPage);
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.InsertBefore(contentFooter + "\t\t");
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.set_Style(((object)Word.WdBuiltinStyle.wdStyleFooter));
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.InsertAfter("\r" + "www.CriticalPathTraining.com");
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.SpaceBefore = 0;

  

    }

    public void SetHeaders(Word.Section section, string contentHeader) {

      string contentFooter = this.documentInfo.FooterText;

      section.PageSetup.DifferentFirstPageHeaderFooter = -1;
      section.PageSetup.TopMargin = 36;
      section.PageSetup.BottomMargin = 36;
      section.PageSetup.LeftMargin = 36;
      section.PageSetup.RightMargin = 36;

      string course_info = documentInfo.CourseCode + ": " + documentInfo.CourseTitle;


      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.set_Style(((object)Word.WdBuiltinStyle.wdStyleHeader));
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.SpaceAfter = 0;

      try {
        section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
      }
      catch { }
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = course_info + "\r" + contentHeader + "\t\tVersion " + documentInfo.Version;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.set_Style(((object)Word.WdBuiltinStyle.wdStyleHeader));
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.SpaceAfter = 12;

      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Fields.Add(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range, Word.WdFieldType.wdFieldPage);
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.InsertBefore(contentFooter + "\t\t");
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.set_Style(((object)Word.WdBuiltinStyle.wdStyleFooter));
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.InsertAfter("\r" + "www.CriticalPathTraining.com");
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.SpaceBefore = 12;

      try {
        section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
      }
      catch { }
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Fields.Add(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range, Word.WdFieldType.wdFieldPage);
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.InsertBefore(contentFooter + "\t\t");
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.set_Style(((object)Word.WdBuiltinStyle.wdStyleFooter));
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.InsertAfter("\r" + "www.CriticalPathTraining.com");
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.SpaceBefore = 12;


    }

    public void EnsureListNumber(Word.Section section) {

      bool ListLevel1Reset = false;
      foreach (Word.Paragraph p in section.Range.Paragraphs) {
        // get paragraph style name
        string style = ((Word.Style)p.get_Style()).NameLocal;
        // reset list numbering for each new exercise (heading section)
        if ( (style.Equals("Heading 2")) || 
             (style.Equals("Heading 3"))  ) {
          ListLevel1Reset = true;
        }

        if(style.Equals(StyleNameLevel1) ){
          if (ListLevel1Reset == true) {
            p.Range.ListFormat.RemoveNumbers();
            p.Range.ListFormat.ApplyListTemplate(LabStepsListTemplate, false, Word.WdListApplyTo.wdListApplyToWholeList);
            ListLevel1Reset = false;
          }
          else {
            p.Range.ListFormat.RemoveNumbers();
            p.set_Style(StyleLevel1);
          }
        }

        if (style.Equals(StyleNameLevel2)) {            
          p.set_Style(StyleLevel2);
        }

        if (style.Equals(StyleNameLevel3)) {
            //p.Range.ListFormat.RemoveNumbers();
            p.set_Style(StyleLevel3);
        }

        if (style.Equals(StyleNameLevel4)){
            //p.Range.ListFormat.RemoveNumbers();
            p.set_Style(StyleLevel4);
        }
      }
    }
    
    public void ScollDocument(Word.Range range) {        
      BuildEnv.ScollDocument(range);
    }

    public void UpdateToc() {
      Console.WriteLine(" - updating TOC...");

      ScollDocument(document.TablesOfContents[1].Range);

      foreach (Word.TableOfContents toc in document.TablesOfContents) {
        toc.Update();
      }

     
    }

    public void RemoveAllComments()
    {
        document.RemoveDocumentInformation(Word.WdRemoveDocInfoType.wdRDIComments);
    }

  }
}
