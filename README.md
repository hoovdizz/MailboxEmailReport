# MailboxEmailReport
Exchange 2010 Email Report of Exchange server
#	Exchange 2010 no longer gives you the option to simply see the size of all mailboxes from long shot. 
# this script will show mailbox count, database size, server size, top ten largest mailboxes and 
# the size of all mailboxes. (to include disconnected mailboxes.) -Alix Hoover 
#
# 12/10/2015 	Added exchange snap in so you can right click run
# 01/07/2016 	Added mailbox count break down by database
# 01/07/2016 	Added status output to see what is being done. 
# 02/02/2016 	Added Prompt to keep the window from auto closing when using the right click function. 
# 02/02/2016 	Added notice that enter can skip email function


#ensure you modify $MailServer and $reportsender
