using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace Apollo {
  class Program {

    static void Main(string[] args) {

      if (args[0] == null) {
        Console.WriteLine("Apollo requires a BuildInfo.xml file");
        CptBuildSet tempBuildSet = new CptBuildSet();
        tempBuildSet.Courses.Add(new CptCourseInfo {});
        XmlSerializer tempSerializer = new XmlSerializer(typeof(CptBuildSet));
        string FileName = Directory.GetCurrentDirectory() +  @"\BuildInfo.xml";
        FileStream tempStream = new FileStream(FileName, FileMode.Create);
        tempSerializer.Serialize(tempStream, tempBuildSet);
        tempStream.Close();
      }

      string BuildFile = args[0];
      XmlSerializer seriaizer = new XmlSerializer(typeof(CptBuildSet));
      FileStream stream = new FileStream(BuildFile, FileMode.Open);
      object rehydration = seriaizer.Deserialize(stream);
      CptBuildSet BuildSet = (CptBuildSet)rehydration;

        // change to build manual
      bool BuildManual = true;
      if (args.Contains("/nomanual")) {
        BuildManual = false;
      }

      bool RefreshUIEnabled = false;
      if (args.Contains("/refreshui")) {
        RefreshUIEnabled = true;
      }
      
      foreach (CptCourseInfo courseInfo in BuildSet.Courses) {
        Console.WriteLine();
        Console.WriteLine("Building " + courseInfo.CourseCode + ": " + courseInfo.CourseTitle);
        BuildEnv.Initialize(courseInfo, BuildManual, RefreshUIEnabled);
        CptCourse course = new CptCourse(courseInfo);
        course.CreateOutput();
      }

      BuildEnv.QuitPowerPoint();
      //BuildEnv.QuitWord();
      

    }

  }
}
