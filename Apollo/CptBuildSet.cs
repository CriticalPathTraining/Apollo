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
  
  [XmlRoot(ElementName = "BuildSet", Namespace = "")]
  public class CptBuildSet {
    [XmlElement(ElementName = "CourseInfo", Namespace = "")]
    public List<CptCourseInfo> Courses { get; set; }
    public CptBuildSet() {
      Courses = new List<CptCourseInfo>();
    }
  }

  public class CptCourseInfo {
    public string SourceDirectory { get; set; }
    public string OmittedModules { get; set; }
    public string CourseTitle { get; set; }
    public string CourseSubtitle { get; set; }
    public string CourseCode { get; set; }
    public string Version { get; set; }
    public string FooterText { get; set; }
    public string BuildFolder { get; set; }
    public string OutputFolder { get; set; }
    public string OutputFileName { get; set; }
    public bool BuildPDF { get; set; }
  }

}
