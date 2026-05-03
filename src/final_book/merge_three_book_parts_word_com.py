import win32com.client as win32
from win32com.client import constants
import os

# ======================
# USER CONFIGURATION
# ======================

BOOK_BODY_PATH = r"C:\Projects\Python\book_contents_formatter_package\release\v1\HITS_AND_HAPPINESS_BODY.docx"
FRONT_MATTER_PATH = r"C:\Projects\Python\book_contents_formatter_package\release\v1\HITS_AND_HAPPINESS_FRONT_MATTER.docx"
OUTPUT_PATH = r"C:\Projects\Python\book_contents_formatter_package\release\v1\HITS_AND_HAPPINESS_FULL_BOOK.docx"

INTRO_WORD = "INTRODUCTION"

# ======================
# WORD APPLICATION SETUP
# ======================

word = win32.Dispatch("Word.Application")
word.Visible = False

doc = word.Documents.Open(BOOK_BODY_PATH)

# ======================
# FIND INTRODUCTION
# ======================

find_range = doc.Content
find = find_range.Find
find.Text = INTRO_WORD
find.Forward = True
find.MatchCase = False

if not find.Execute():
    raise RuntimeError("INTRODUCTION not found in the document.")

intro_range = find_range.Duplicate

# ======================
# INSERT SECTION BREAK
# ======================

intro_range.Collapse(constants.wdCollapseStart)
intro_range.InsertBreak(constants.wdSectionBreakNextPage)

# ======================
# MOVE TO FRONT SECTION
# ======================

front_section = doc.Sections(1)
front_range = front_section.Range
front_range.Collapse(constants.wdCollapseStart)

# ======================
# INSERT FRONT MATTER FILE
# ======================

front_range.InsertFile(FRONT_MATTER_PATH)

# ======================
# FRONT MATTER PAGE NUMBERING
# ======================

front_header = front_section.Headers(constants.wdHeaderFooterPrimary)
front_footer = front_section.Footers(constants.wdHeaderFooterPrimary)

front_header.LinkToPrevious = False
front_footer.LinkToPrevious = False

front_footer.PageNumbers.RestartNumberingAtSection = True
front_footer.PageNumbers.NumberStyle = constants.wdPageNumberStyleLowercaseRoman
front_footer.PageNumbers.StartingNumber = 1

front_footer.PageNumbers.Add(
    PageNumberAlignment=constants.wdAlignPageNumberCenter
)

# ======================
# BOOK BODY NUMBERING
# ======================

body_section = doc.Sections(2)
body_header = body_section.Headers(constants.wdHeaderFooterPrimary)
body_footer = body_section.Footers(constants.wdHeaderFooterPrimary)

body_header.LinkToPrevious = False
body_footer.LinkToPrevious = False

body_footer.PageNumbers.RestartNumberingAtSection = True
body_footer.PageNumbers.NumberStyle = constants.wdPageNumberStyleArabic
body_footer.PageNumbers.StartingNumber = 1

body_footer.PageNumbers.Add(
    PageNumberAlignment=constants.wdAlignPageNumberCenter
)

# ======================
# SAVE FINAL DOCUMENT
# ======================

if os.path.exists(OUTPUT_PATH):
    os.remove(OUTPUT_PATH)

doc.SaveAs(OUTPUT_PATH)
doc.Close()
word.Quit()

print("✅ Full book successfully created:")
print(OUTPUT_PATH)
