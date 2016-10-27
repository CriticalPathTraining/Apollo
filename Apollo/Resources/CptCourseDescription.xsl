<?xml version="1.0" encoding="utf-8"?>

<xsl:stylesheet
  version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:cd="http://www.criticalpathtraining.com/schema/coursedescription/1.0">

  <xsl:output method="html" indent="yes" doctype-public="html" />

  <xsl:template match="/cd:CourseDescription">

    <html>

      <head>
        <title>
          <xsl:value-of select="cd:CourseCode"/>: <xsl:value-of select="cd:CourseTitle"/>
        </title>
        <link rel="stylesheet" type="text/css" href="CptCourseDescription.css"/>
      </head>

      <body>

        <div id="course_title">
          <xsl:value-of select="cd:CourseTitle"/>
        </div>
        <div id="course_subtitle">
          <xsl:value-of select="cd:CourseSubtitle"/>
        </div>

        <table id="course_table">
          <tr>
            <td>Course Code:</td>
            <td>
              <xsl:value-of select="cd:CourseCode"/>
            </td>
          </tr>
          <tr>
            <td>Audience:</td>
            <td>
              <xsl:value-of select="cd:Audience"/>
            </td>
          </tr>

          <tr>
            <td>Format:</td>
            <td>
              <xsl:value-of select="cd:Format"/>
            </td>
          </tr>
          <tr>
            <td>Length:</td>
            <td>
              <xsl:value-of select="cd:Length"/>
            </td>
          </tr>
          <tr>
            <td>Description:</td>
            <td>
              <xsl:for-each select="/cd:CourseDescription/cd:Description/cd:p">
                <div class="course_description">
                  <xsl:value-of select="text()"/>
                </div>
              </xsl:for-each>
            </td>
          </tr>
          <tr>
            <td>Prerequisites:</td>
            <td>
              <xsl:for-each select="/cd:CourseDescription/cd:Prerequisites/cd:p">
                <div class="course_prerequisites">
                  <xsl:value-of select="text()"/>
                </div>
              </xsl:for-each>
            </td>
          </tr>

        </table>


        <p class="course_header" >Course Modules</p>

        <ol class="module_list">
          <xsl:for-each select="/cd:CourseDescription/cd:Modules/cd:Module/cd:Title">
            <li>
              <xsl:value-of select="text()"/>
            </li>
          </xsl:for-each>
        </ol>


        <p class="section_header" >Course Module Contents</p>


        <xsl:for-each select="/cd:CourseDescription/cd:Modules/cd:Module">
          <div class="module_section">
            <h1 class="module_title" >
              Module <xsl:value-of select="cd:Number"/>: <xsl:value-of select="cd:Title"/>
            </h1>
            <div class="module_description">
              <xsl:value-of select="cd:Description"/>
            </div>
            <div class="agenda_topics_caption">Agenda Topics</div>
            <ul class="module_topics" >
              <xsl:for-each select="cd:AgendaTopics/cd:AgendaTopic">
                <li>
                  <xsl:value-of select="text()"/>
                </li>
              </xsl:for-each>
            </ul>
            <xsl:for-each select="cd:Labs/cd:Lab">
              <div class="module_lab">
                <h2 class="lab_title" >
                  Lab: <xsl:value-of select="cd:Title"/>
                </h2>
                <ul class="lab_exercises" >
                  <xsl:for-each select="cd:Exercises/cd:Exercise">
                    <li class="lab_exercise">
                      <xsl:value-of select="text()"/>
                    </li>
                  </xsl:for-each>
                </ul>
              </div>
            </xsl:for-each>
          </div>
        </xsl:for-each>

      </body>
    </html>

  </xsl:template>

</xsl:stylesheet>