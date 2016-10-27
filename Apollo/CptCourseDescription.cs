using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace Apollo {

  public class Namespaces {
    public const string CptCourseware = "http://www.criticalpathtraining.com/schema/coursedescription/1.0";
  }

  [XmlRoot(ElementName = "CourseDescription", Namespace = Namespaces.CptCourseware)]
  public class CptCourseDescription {
    [XmlElement(Order = 1)]
    public string CourseTitle { get; set; }
    [XmlElement(Order = 2)]
    public string CourseSubtitle { get; set; }
    [XmlElement(Order = 3)]
    public string Audience { get; set; }
    [XmlElement(Order = 4)]
    public string Format { get; set; }
    [XmlElement(Order = 5)]
    public string Length { get; set; }
    [XmlElement(Order = 6)]
    public string CourseCode { get; set; }
    [XmlArray("Description", Order = 7), XmlArrayItem("p")]
    public List<string> Description { get; set; }
    [XmlArray("Prerequisites", Order = 8), XmlArrayItem("p")]
    public List<string> Prerequisites { get; set; }
    [XmlElement(Order = 9)]
    public string Version { get; set; }
    [XmlArray("Modules", Order = 10), XmlArrayItem("Module")]
    public List<CptCourseModule> Modules = new List<CptCourseModule>();
  }

  public class CptCourseModule {
    [XmlElement("Number", Order=1)]
    public string Number { get; set; }
    [XmlElement("Title", Order=2)] 
      public string Title { get; set; }
    [XmlElement("Description", Order=3)] 
      public string Description { get; set; }
    [XmlArray("AgendaTopics", Order = 4), XmlArrayItem("AgendaTopic")]
    public List<string> AgendaTopics = new List<string>();
    [XmlArray("Labs", Order = 5), XmlArrayItem("Lab")]
    public List<CptCourseLab> Labs = new List<CptCourseLab>();
  }

  public class CptCourseLab {
    [XmlElement("Title", Order=1)] 
    public string Title { get; set; }
    [XmlArray("Exercises", Order = 2), XmlArrayItem("Exercise")]
    public List<string> Exercises = new List<string>();
  }

}
