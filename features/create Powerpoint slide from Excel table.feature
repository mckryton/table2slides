Ability: create Powerpoint slides from Excel table
  Telling a story often starts with assembling a story line by scetching
  the aspects of the story as short text blocks. Excel provides a lot of
  flexibility for reordering those text blocks by still keeping the whole
  story in sight. On the other hand Powerpoints offers the better tools 
  to present the story in a more readable form by distributing small 
  chunks on Powerpoints slides.
  This application will convert the story contents from
  an Excel table into Powerpoint slides to combine the advantages from 
  both tools.
  

  Rule: a table with a column named description should result in a slide with description as title
    @wip
    Scenario: table with consecutive descriptions
      Given an Excel sheet with a table 
       | description |
       | title 1     |
       | title 2     |
       When the table is converted into slides
       Then a new presentations is created
        And the presentation has 2 new slides
        And slide 1 has the title "title 1"
        And slide 2 has the title "title 2"

    Scenario: empty sheet

    Scenario: table without data 

    Scenario: table without description column
      