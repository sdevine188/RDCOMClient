# https://stackoverflow.com/questions/51957310/error-converting-docx-to-pdf-using-pandoc

# convert word docx to pdf, for example after knitting to docx
# IMPORTANT: it seems like the file path to the docx cannot have any spaces in it!!
# I got errors trying to convert
# "C:/Users/Stephen/Desktop/University of Wisconsin/classes/DS710/ds710spring2019assignment1/Assignment_1_R.docx"
# but no errors converting the same doc saved to
# "C:/Users/Stephen/Desktop/R/RDCOMClient/Assignment_1_R.docx"

library(RDCOMClient)
library(fs)

setwd("C:/Users/Stephen/Desktop/R/RDCOMClient")
dir_ls()
file <- "C:/Users/Stephen/Desktop/R/RDCOMClient/test_document2.docx"
file <- "C:/Users/Stephen/Desktop/R/RDCOMClient/Assignment_1_R.docx"

file <- "C:/Users/Stephen/Desktop/University of Wisconsin/classes/DS710/ds710spring2019assignment1/Assignment_1_R.docx"


wordApp <- COMCreate("Word.Application")  # create COM object
wordApp[["Visible"]] <- TRUE #opens a Word application instance visibly
wordApp[["Documents"]]$Add() #adds new blank docx in your application
wordApp[["Documents"]]$Open(Filename=file) #opens your docx in wordApp

#THIS IS THE MAGIC    
wordApp[["ActiveDocument"]]$SaveAs("C:/Users/Stephen/Desktop/R/RDCOMClient/test_document2_converted.pdf", 
                                   FileFormat=17) #FileFormat=17 saves as .PDF
wordApp[["ActiveDocument"]]$SaveAs("C:/Users/Stephen/Desktop/R/RDCOMClient/Assignment_1_R_converted.pdf", 
                                   FileFormat=17) #FileFormat=17 saves as .PDF

wordApp$Quit() #quit wordApp


################################


# library(reticulate)
# 
# com <- import("win32com.client")
# 
# file <- "C:/Users/Stephen/Desktop/University of Wisconsin/classes/DS710/ds710spring2019assignment1/Assignment_1_R.docx"
# 
# wordPy <- com$gencache$EnsureDispatch("Word.Application")
# wordPyOpen <- wordPy$Documents$Open(file)
# wordPyOpen$SaveAs("C:/path/to your/doc.pdf",
#                   FileFormat=17)
# wordPy$Quit()