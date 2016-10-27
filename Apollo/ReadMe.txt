
Course Assumptions:
- Course "Title" and "Subtitle" created from BuildInfo.xml file

Module Assumptions:
- Module Title created using Slide Presentation Title from PPTX file metadata
- A "Topic Slide" is slide whose Title not equal to "Agenda", "Summary" or "Demo"
- First slide is "Intro Slide" and not considered a topic slide
- Module description created using notes section under "Intro Slide"
- "Lecture Agenda" created from bullet points of first slide with Title of "Agenda"
- "Topics covered" created from Slide Title of all topic slides
- "Instructor Demos" created from body text of all slides with Title of "Demo"


Existing Slide Deck Fixes
- Ensure PPTX file has correct Title in document's Title property
- Add Module description into notes section under intro slide
- Clean up notes section for slides if required


Existing Lab File Fixes
- Remove graphic image from Heading1 paragraph 
