library(tidyverse)
library(RDCOMClient)
library(RDCOMOutlook)
library(lubridate)
library(fs)


# setwd
setwd("H:/R/RDCOMClient")


############################################################################################


# create get_email_info function
get_email_info <- function(email) {
        
        # get email info
        subject <- email$Subject()
        date_received <- ymd_hms(as.POSIXct(email$ReceivedTime() * (24 * 60 * 60), origin="1899-12-30", tz = "GMT")) 
        sender_name <- email$SenderName()
        sender_email_address <- email$SenderEmailAddress()
        cc_email_address <- email$CC()
        importance <- email$Importance()
        number_attachments <- email$Attachments()$Count()
        attachment_names <- map(.x = seq(2), .f = function(.x) {email_test$Attachments[.x]}) %>% unlist() %>% str_c(string = ., collapse = "; ")
        body <- email$Body()
        
        # combine email info in tbl
        email_info_tbl <- tibble(subject = subject, date_received = date_received, sender_name = sender_name, 
                                 sender_email_address = sender_email_address, cc_email_address = cc_email_address,
                                 importance = importance, number_attachments = number_attachments, 
                                 attachment_names = attachment_names, body = body)
        return(email_info_tbl)
}

# test
email <- emails[[1]][[1]]
email
emails %>% pmap_dfr(.l = ., .f = get_email_info)


##########


# create function to search outlook inbox
# https://rdrr.io/github/mdneuzerling/RDCOMOutlook/src/R/search_emails.R

search_emails <- function(search_term, folder = "Inbox", scope = "subject", partial_match = TRUE, 
                          search_subfolders = TRUE, search_time = 10) {
        
        # initialize outlook application
        outlook_app <- COMCreate("Outlook.Application")
        
        # define scope of search
        scope <- if(scope == "subject") {"subject"} else 
                if (scope == "body") {"textdescription"} else 
                if(scope == "attachment_names") {"attachmentfilename"} else 
                if(scope == "from_name") {"fromname"} else 
                if(scope == "from_email") {"fromemail"} else 
                if(scope == "cc_name") {"displaycc"} else 
                if(scope == "to_name") {"displayto"} 
        
        # build search_query
        search_query <- str_c("urn:schemas:httpmail:", scope, if(partial_match == TRUE) {" LIKE '%"} else {" = '"},
                search_term, if(partial_match == TRUE) {"%'"} else {"'"})
        
        # execute search
        search <- outlook_app$AdvancedSearch(folder, search_query, search_subfolders)
        
        # let system sleep while search executes
        Sys.sleep(search_time)
        
        # get search results into tibble
        results <- search$results()
        number_results <- results$Count()
        emails <- map(.x = seq(number_results), .f = function(.x) {results$Item(.x)}) %>% tibble(email = .)
        return(emails)
}

# get email_info_tbl
get_email_info_tbl <- function(emails){
        
        # get email_info_tbl
        email_info_tbl <- emails %>% pmap_dfr(.l = ., .f = get_email_info)
        return(email_info_tbl)
}

# test
search_term = "email is for Stephen"
folder <- "Inbox"
scope <- "subject"
partial_match <- TRUE
search_subfolders <- TRUE
search_time <- 10

# run search_emails
search_output <- search_emails(search_term = "email is for Stephen", folder = "Inbox", scope = "subject", partial_match = TRUE, 
              search_subfolders = TRUE, search_time = 10)        
search_output

# run get_email_info
email_info_tbl <- get_email_info_tbl(emails = search_output)
email_info_tbl %>% slice(1) %>% pull(body) %>% cat()
email_info_tbl %>% slice(1) %>% pull(attachment_names)


############################################################################################


# save email and attachment
# https://github.com/mdneuzerling/RDCOMOutlook/blob/master/R/save_attachments.R

# create save_attachments function
save_attachments <- function(email) {
        
        # create current_email
        current_email <- email
        
        # get current_email_subject
        current_email_subject <- current_email$Subject()
        current_email_subject <- str_replace_all(string = current_email_subject, pattern = " ", replacement = "_")
        
        # get current_email_date_recieved
        current_email_date_received <- ymd_hms(as.POSIXct(current_email$ReceivedTime() * (24 * 60 * 60), origin="1899-12-30", tz = "GMT")) 
        current_email_date_received <- floor_date(current_email_date_received, "day")
        current_email_date_received <- str_replace_all(string = current_email_date_received, pattern = " ", replacement = "_")
        
        # create attachment_folder
        attachment_folder <- str_c(current_email_subject, "__", current_email_date_received)
        dir_create(path = attachment_folder)
        
        # get count of attachments on current_email
        attachment_count <- current_email$Attachments()$Count()
        
        map(.x = seq(attachment_count), .f = function(.x) {
                current_attachment <- current_email$Attachments(.x)
                current_attachment_location <- str_c(getwd(), "/", attachment_folder, "/", current_attachment$FileName())
                print(str_c("now saving ", current_attachment$FileName()))
                current_attachment$SaveAsFile(current_attachment_location)
        })
}


# test
# can't seem to save email itself, only it's constituent data (eg body/subject, etc) and attachments
email <- emails[[1]][[1]]
email$SaveAsFile("email_test.msg")
email$SaveAsFile("email_test")


search_output <- search_emails(search_term = "save_attachments test", folder = "Inbox", scope = "subject", partial_match = TRUE, 
                               search_subfolders = TRUE, search_time = 10) 
search_output
search_output %>% pmap(.l = ., .f = function(email, ...) {email}) 
search_output %>% pwalk(.l = ., .f = save_attachments) 



############################################################################################
###########################################################################################
###########################################################################################


# # https://stackoverflow.com/questions/42573699/how-to-retrieve-outlook-inbox-emails-using-r-rdcomclient
# 
# # read outlook emails from inbox
# folder_name <- "AUX"
# 
# ## create outlook object
# OutApp <- COMCreate("Outlook.Application")
# outlookNameSpace = OutApp$GetNameSpace("MAPI")
# 
# folder <- outlookNameSpace$Folders(1)$Folders(folder_name)  ## i get an error at this step??
# # Check that we got the right folder
# folder$Name(1)
# 
# emails <- folder$Items
# 
# # Can't figure out how to get number of items, so just doing first 10
# for (i in 1:10)
# {
#         subject <- emails(i)$Subject(1)
#         # Replace "#78" with the text you are looking for in Email Subject line
#         if (grepl("#78", subject)[1]){
#                 print(emails(i)$Body())
#                 break
#         } 
# }


###############################################################################################


# # check inbox for specific subject line
# outlook_app <- COMCreate("Outlook.Application")
# search <- outlook_app$AdvancedSearch(
#         "Inbox",
#         # "urn:schemas:httpmail:subject = 'this email is for Stephen'" # this is full subject line
#         "urn:schemas:httpmail:subject LIKE '%email is for Stephen%'" # this is a wildcard search, with partial subject line
#         
# )
# 
# results <- search$Results()
# results
# 
# # get email sent date
# results$Item(1)
# results$Item(1)$ReceivedTime() # Received time of first search result
# as.Date("1899-12-30") + floor(results$Item(1)$ReceivedTime()) # Received dat
# 
# # get email subject and body
# search$Results()$Item(1)
# search$Results()$Item(1)$Subject()
# search$Results()$Item(1)$Body()

