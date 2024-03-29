<?xml version="1.0" encoding="utf-16"?>

<xsl:stylesheet
  version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:cd="http://www.criticalpathtraining.com/schema/coursedescription/1.0">

  <xsl:template match="/cd:CourseDescription">

    <h1 class="CourseTitle">
      <xsl:value-of select="cd:CourseCode"/>: <xsl:value-of select="cd:CourseTitle"/>
    </h1>


    <table class="QuickFactsTable" cellpadding="4" cellspacing="0" border="1">


      <tr>
        <td class="QuickFactsTableHeader" valign="top">
          Description:
        </td>
        <td class="QuickFactsTableContent" valign="top">
          <xsl:value-of select="cd:Description"/>
        </td>
      </tr>

      <tr>
        <td class="QuickFactsTableHeader" valign="top">
          Prerequisites:
        </td>
        <td class="QuickFactsTableContent" valign="top">
          <xsl:value-of select="cd:Prerequisites"/>
        </td>
      </tr>
    </table>

    <p class="CourseSectionHeader" >Schedule of Lectures</p>

    <ol class="ListOfLectures">
      <xsl:for-each select="/cd:CourseDescription/cd:Day/cd:Lectures/cd:Lecture/cd:Title">
        <li>
          <xsl:value-of select="text()"/>
        </li>
      </xsl:for-each>
    </ol>


    <xsl:for-each select="/cd:CourseDescription/cd:Day">
      <div style="background-color:#DDDDDD">
        <span style="font-family:'Arial Black';font-size:12pt;color:#488000">
          &#160;Day <xsl:value-of select="cd:Number"/>
        </span>
        <span style="font-family:'Arial Black'; font-size:9pt;color:#9C6838 ;width:400pt">
          &#160;&#160;(runs from <xsl:value-of select="cd:Schedule"/>)
        </span>
      </div>
      <br/>

      <xsl:for-each select="cd:Lectures/cd:Lecture">
        <h1 class="LectureTitle" >
          <xsl:value-of select="cd:Title"/>
          <span style="font-family:'Arial Black'; font-size:7pt;color:#658530">
            &#160;&#160;<xsl:value-of select="cd:Schedule"/>
          </span>
        </h1>
        <ul class="LectureTopics">
          <xsl:for-each select="cd:LectureTopics/cd:Topic">
            <li>
              <xsl:value-of select="text()"/>
            </li>
          </xsl:for-each>
        </ul>
      </xsl:for-each>

    </xsl:for-each>

  </xsl:template>

</xsl:stylesheet>