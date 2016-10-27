using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace Apollo {
  public class CptSellsheetWriter {

    private static Word.Document document;
    private static Word.Range EndOfDoc {
      get { return document.Bookmarks[((object)@"\endofdoc")].Range; }
    }

    private static void SetPageLayout() {
      document.Content.PageSetup.DifferentFirstPageHeaderFooter = -1;
      document.Content.PageSetup.TopMargin = 36;
      document.Content.PageSetup.BottomMargin = 72;
      document.Content.PageSetup.LeftMargin = 36;
      document.Content.PageSetup.RightMargin = 36;
    }

    private static void SetHeaders() {
      Word.Section section = document.Sections.First;
      string contentHeader = courseDescription.CourseCode + ": " + courseDescription.CourseTitle;
      string contentFooter = "© Critical Path Training. 2016. All Rights Reserved";


      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.set_Style(((object)Word.WdBuiltinStyle.wdStyleHeader));
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.SpaceAfter = 0;

      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = contentHeader + "\t\tVersion " + courseDescription.Version;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.set_Style(((object)Word.WdBuiltinStyle.wdStyleHeader));
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.SpaceAfter = 12;

      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "www.CriticalPathTraining.com\tinfo@criticalpathtraining.com\t(866)475-4440";
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.set_Style(((object)Word.WdBuiltinStyle.wdStyleFooter));
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.ParagraphFormat.SpaceBefore = 12;

      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Fields.Add(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range, Word.WdFieldType.wdFieldPage);
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.InsertBefore(contentFooter + "\t\t");
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.set_Style(((object)Word.WdBuiltinStyle.wdStyleFooter));
      section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
      section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.SpaceBefore = 12;
    }

    private static CptCourseDescription courseDescription;

    public static void CreateSellSheet(CptCourseDescription CourseDescription, string TargetFileName, string WordStyleTemplate) {
      courseDescription = CourseDescription;
      Word.Application app = new Word.Application();
      app.Visible = true;
      document = app.Documents.Add();
      SetPageLayout();
      SetHeaders();
      document.CopyStylesFromTemplate(WordStyleTemplate);

      AddBanner();
      AddParagraph(CourseDescription.CourseTitle, Word.WdBuiltinStyle.wdStyleTitle);
      AddParagraph(CourseDescription.CourseSubtitle, Word.WdBuiltinStyle.wdStyleSubtitle);

      Word.Range rTable = EndOfDoc;
      Word.Table CourseTable = rTable.Tables.Add(rTable, 6, 2);

      //CourseTable.Borders.Enable = 1;
      //CourseTable.LeftPadding = 4;
      //CourseTable.RightPadding = 4;
      //CourseTable.TopPadding = 4;
      //CourseTable.BottomPadding = 4;

      //CourseTable.TopPadding = 36;

      CourseTable.Rows.LeftIndent = 0;
      CourseTable.TopPadding = 2;
      CourseTable.RightPadding = 4;
      CourseTable.BottomPadding = 0;
      CourseTable.LeftPadding = 4;
      CourseTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthAuto;

      CourseTable.Rows[1].Cells[1].Range.Text = "Course Code";
      CourseTable.Rows[1].Cells[2].Range.Text = CourseDescription.CourseCode;
        
      CourseTable.Rows[2].Cells[1].Range.Text = "Audience";
      CourseTable.Rows[2].Cells[2].Range.Text = CourseDescription.Audience;
      CourseTable.Rows[3].Cells[1].Range.Text = "Format";
      CourseTable.Rows[3].Cells[2].Range.Text = CourseDescription.Format;
      CourseTable.Rows[4].Cells[1].Range.Text = "Length";
      CourseTable.Rows[4].Cells[2].Range.Text = CourseDescription.Length;

      CourseTable.Rows[5].Cells[1].Range.Text = "Course Description";

      foreach (string para in CourseDescription.Description) {
        Word.Paragraphs paras = CourseTable.Rows[5].Cells[2].Range.Paragraphs;
        if (paras.Count == 1 && paras.First.Range.Text.Equals("\r\a")) {
          paras.First.Range.Text = para;
        }
        else {
          Word.Paragraph p = paras.Add();
          p.Range.Text = para;
        }
      }

      CourseTable.Rows[6].Cells[1].Range.Text = "Student Prerequisites";

      foreach (string para in CourseDescription.Prerequisites) {
        Word.Paragraphs paras = CourseTable.Rows[6].Cells[2].Range.Paragraphs;
        if (paras.Count == 1 && paras.First.Range.Text.Equals("\r\a")) {
          paras.First.Range.Text = para;
        }
        else {
          Word.Paragraph p = paras.Add();
          p.Range.Text = para;
        }

      }

      //CourseTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
      //CourseTable.UpdateAutoFormat();
      CourseTable.Columns[1].AutoFit();
      CourseTable.Columns[2].AutoFit();
      CourseTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
      CourseTable.Borders.InsideLineWidth = Word.WdLineWidth.wdLineWidth100pt;
      CourseTable.Borders.InsideColor = Word.WdColor.wdColorGray50;

      CourseTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
      CourseTable.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth100pt;
      CourseTable.Borders.OutsideColor = Word.WdColor.wdColorGray50;
      CourseTable.Rows.DistanceLeft = 0;
      CourseTable.Rows.DistanceRight = 0;


      AddParagraph("Course Modules", "Course Heading");


      foreach (CptCourseModule module in CourseDescription.Modules) {
        AddParagraph(module.Title, "Module Title Numbered Item");

      }

      // Word.Section section = document.Sections.Add(EndOfDoc);
      // section.PageSetup.SectionStart = Word.WdSectionStart.wdSectionNewPage;

      AddParagraph("Course Module Detailed Outline", "Section Heading");


      foreach (CptCourseModule module in CourseDescription.Modules) {
        AddParagraph("Module " + module.Number + ": " + module.Title, "Module Title");
        AddParagraph(module.Description, "Module Description");

        AddParagraph("Topics Covered", "Module Heading");

        foreach (string topic in module.AgendaTopics) {
          if (!string.IsNullOrWhiteSpace(topic)) {
            AddParagraph(topic, "Module Bullet Point");
          }
        }

        foreach (CptCourseLab lab in module.Labs) {
          AddParagraph("Hands-on Lab: " + lab.Title, "Module Heading");
          foreach (string Exercise in lab.Exercises) {
            if (!string.IsNullOrWhiteSpace(Exercise)) {
              AddParagraph(Exercise, "Module Bullet Point");
            }
          }
        }
      }
      // save doc
      document.SaveAs2(TargetFileName);

      //if (courseInfo.BuildPDF) {
      //  manual.SaveAs(BuildEnv.CourseManualPdfFilePath, Word.WdSaveFormat.wdFormatPDF);
      //}

      ((Word._Document)document).Close(Word.WdSaveOptions.wdDoNotSaveChanges);
  
    }

    public static void AddBanner() {
      Word.Paragraph paragraph = document.Content.Paragraphs.First;
      string pathBannerImage = BuildEnv.BuildFolder + @"\CptCourseDescriptionHeader.png";
      paragraph.Range.InlineShapes.AddPicture(pathBannerImage);
      object h1 = "Top Banner";
      paragraph.set_Style(ref h1);
      paragraph.Range.InsertParagraphAfter();
    }

    public static void AddParagraph(string content) {
      AddParagraph(content, Word.WdBuiltinStyle.wdStyleNormal);
    }

    public static void AddParagraph(string content, Word.WdBuiltinStyle style) {
      Word.Paragraph paragraph = document.Content.Paragraphs.Add(EndOfDoc);
      paragraph.Range.Text = content;
      object h1 = style;
      paragraph.set_Style(ref h1);
      paragraph.Range.InsertParagraphAfter();
    }

    public static void AddParagraph(string content, Word.WdBuiltinStyle style, Word.Range range) {
      Word.Paragraph paragraph;
      if (range.Paragraphs.Count == 1 && string.IsNullOrEmpty(range.Paragraphs.First.Range.Text)) {
        paragraph = range.Paragraphs.First;
      }
      else {
        paragraph = range.Paragraphs.Add();
      }
      paragraph.Range.Text = content;
      object h1 = style;
      paragraph.set_Style(ref h1);
    }

    public static void AddParagraph(string content, string style) {
      Word.Paragraph paragraph = document.Content.Paragraphs.Add(EndOfDoc);
      paragraph.Range.Text = content;
      object h1 = style;
      paragraph.set_Style(ref h1);
      paragraph.Range.InsertParagraphAfter();
    }

    public static void AddParagraph(string content, string style, Word.Range range) {
      Word.Paragraph paragraph;
      if (range.Paragraphs.Count == 1 && string.IsNullOrEmpty(range.Paragraphs.First.Range.Text)) {
        paragraph = range.Paragraphs.First;
      }
      else {
        paragraph = range.Paragraphs.Add();
      }
      paragraph.Range.Text = content;
      object h1 = style;
      paragraph.set_Style(ref h1);
    }


  }
}
