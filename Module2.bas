Attribute VB_Name = "Module2"
Option Explicit

Global Const version_date_s = "4/25/2019"
Global Const major_version = 3
Global Const minor_version = 7006
Global Const file_format = 6  ' Defined 2/2/2003
'Global Const file_format = 5  ' Defined 5/16/2000
'Global Const file_format = 5  ' Defined 5/7/2001  Added credit cards

' Revision history
' 3.7006  4/25/2019
' Added the Averages to the Monthly Balance form.
' Changed the Balance form to highlight the month were are on in yellow and to cyan when on the current actual month

' 3.7004  4/19/2019
' Added Average balances to the Monthly Balance form.

' 3.7003 4/18/2019
' Added the Delta column in the Balance Summary screen which shows how you did for the month

' 3.7002 2/2/2019
' When cleaning out junk I saw that twips.bas was not being loaded with the project and
' was located 1 folder above. I added twips.bas as a module so now it loads and runs properly.
' Changed the height of the Cardtrak screen to show empty lines at the bottom.

' 3.7001 2/2/2019
' In the last revision I changed a line which allowed CardTrack to display totals even if not paid.
' This messed up displaying things correctly and even messed with the checkbook balance.
' I took that line out and replaced it with the original. cardtrak.frm (Module3)
' Updated rev number


' 3.6006 6/3/2018
' Haven't started but need to change the Copy/Paste Selected to use the single Copy/Paste menu click
' Updated the Card Track Summary form to add in a Net field = -paid-purchases-interest-late
' This primarily affects ct_summary_form
' In ct_summary_form added the hover over a Picture1 bar to show the value
' Removed the requirement to be registered to run in module2.bas
' Added Net and Est Balance to printout
' For Card Track removed the Paid qualification to be shown in the grid which allows future entries to be included
' Alighened the center buttons on the main form

' 3.6005 12/13/2017
' Working on the Copy/Paste Selected - I had removed it when trying to combine the Copy/Paste into a
' single menu item rather than the Selected menu item and I'm putting it back to the way it was originally
' For this release until I have time to fix it

' 3.6004 1/28/2017
' Added set_center_tab to module2.bas
' Changing GoTo M/Y to call this
' Added the menu items and buttons for Next/Previous Year and Center Tab
' Moved the Next/Previous Month group of buttons to be centered above the 6th tab
' Added the gif bitmaps for these buttons
' Added the double click any tab will set the center tab
' Added the todays_date_button in the navigation group which does the same function as todays_date_label_Click

' 3.6003 1/28/2017
' Fixed the Go To Month/Year menu click
' Fixed the Month/Year form to add the list of years up to 2051 - I will be 100 years old on August 25, 1951 so good luck!
' Fixed the GoTo M/Y to make the center tab active
' Set the Debug flag to 0 in ? so that the Balance column shows correctly
' Added the Today button on the GoTo Y/D
' Made the year drop down on GoTo Y/D to go from high to low years
' Validate the year entry in GoTo Y/D so that it can be typed in

' 3.6002 1/26/2017
' Commented out the entire section in Function read_database below by adding the
' Exit Function line. With this entire section executing it was taking over 10 seconds to
' rebuild the database. By removing this section it loads very quickly in under 1 second.
' I still don't remember why I felt it was necessary to rebuild the ebtire
' database, validate all records and delete duplicate records.
' All the calculations seem to be the same when loading a file with over 6600 records.
' I also noticed that wheel scrolling stopped working.
' Fixed the Runtime Error Invalid Row when double clicking on a calendar date. The problem was in main_form update_entry_tabs
' where the variable i was not being declared locally and was being corrupted. This error seems to
' be present in the stable version 3.4011.
' Reenabled mouse wheel scrolling in wheel.bas by commentint out the Exit Sub line in Sub WheelHook.
' For some unknown reason back in 2005 this Exit Sub was inserted which would prevent
' mouse wheel scrolling of the FlexGrid on the main form. By commenting out this line the
' mouse wheel scrolling is operational again. I am using the touch pad on my laptop to verify it works. Maybe
' there was some issue when using a regular wheel mouse. I'm going to leave this wheel scrolling
' active. The stable version, 3.4011, used this function with no known problems.

' 3.6001 1/25/2017
' Commented out the line to delete the duplicate records
' The entire section to create a rebuilt database was added after the stable version of 3.4011 5/10/2004
' See the comments in Function read_database below
' Updated web site references to www.mycheck2check.com
' Updated frmAbout.frm fields

' 3.5002 1/25/2017
' Recompiled after reinstallation of Visual Basic 6.0 (SP6) on my laptop

' 3.5001 11/21/2010
' Added OnError when opening a non-existent file or drive

' 3.4011 5/10/2004
' Commented out the Spanish menu item
'
' 3.4010 5/1/2004
' Added some sounds on Enter key hit and startup and exit
'
' 3.4009 4/6/2004
' Started chaning the dd/mm/yy format to use the regional settings
'
' 3.4008 3/8/2004
' Restart in language last used
' Word spacing in CT form
' Added auto reload last file
'
' 3.4007 3/7/2004
' Converted the remaining screens to Spanish - still have msgbox everywhere
'
' 3.4006 3/6/2004
' Changed the way the transactions, monthly notes, and misc notes are displayed
'
' 3.4005 3/5/2004
' Changed a couple words in the spanish translations
' Added the misc notes button and window
' Changed the View Quick Accounts button tool tip
' When editing an amount use the sign of the previous amount
' Added a go to previous bottom and next top month buttons
' Added buttons to go to the top and bottom of current month
'
' 3.4004 2/1/2004
' Added wheel mouse scrolling
'
' 3.4003 7/7/2003
' Added name to data mining message
' Fixed the copy month so that it will not chop off the last record

' 3.4    4/22/2003
' Moved the redim statements in sub main to before the expiration screen is used to prevent a runtime error
' Fixed the auto check number in the edit screen and card data form
' Show the current reference in the reference number frame on the notes
' Fixed the notes font size
'
' 3.3002 4/20/2003
' 3.3001 4/19/2003
' Started adding Spanish and multilangage
'
' 3.3   4/6/2003
' Updated the version number only
'
' 3.2   3/22/2003
' Fixed the move selected so it wouldn't clear the paid column
' Change read only attributes to normal on file save
' Added internet data mining
' Added clear filter button to filter form
'
' 3.104 3/6/03
' Adding the cardtrak summary graph
' Added cardtrak printing
'
' 3.103 3/6/03
' Added more cardtrak and summary screen
'
' 3.102 2/14/03
' Added more card trak
'
' 3.101 2/2/03
' Quick transactions are marked as cleared when done
' Added card trak
'
' 3.1   9/9/02
' Uncommented the password button allowing passwords to work
'
' 3.0   5/14/2002
' Raised price to $19.95
'
' Changed rev number
' 2.7   5/12/2002
' Added the quick save/deposit accounts
' Added passwords
'
' 2.6   1/4/2002
' Fixed the run time problem with the file new
'
' 2.5   1/2/2002
' Fixed the run time problem with the form load and an invalid month
'
' 2.4   8/7/01
' Started to add the credit card stuff but kept it commented out
' Added copy tags with arrange
' Added copy/cut selected, paste selected
' Changed price to $9.95
'
' 2.3   4/5/01
' Fixed the reconcile problem with changing record zero
' Added the changed flag to transaction delete
'
' 2.2   3/24/01
' Changed the reconcile code to use redim to save memory
'
' 2.2   3/19/01
' Added full checkbook reconcilation
' Added Added a transaction edit screen
' Added paste month preferences
' Added date due column and to transactions
'
' 2.1   1/9/01
'
' 2.0   1/8/01
' New release
'
' 1.35  1/3/01
' Fixed the go to web and email in about form
'
' 1.34  1/3/01
' Changed the tag screen
'
' 1.33  1/2/01
' Added filter for tags
' Added keyboard entry for tag number
' Fixed filter on status
'
' 1.32  12/31/00
' Jazzed up the screens, added transaction filter and button
' Added view balances
' Added view tags
' Added view summary
' Added check integrity form
' Added notes font size
'
' 1.3   6/16/00
' Added tons of stuff
'
' 1.2   5/23/00
' Fixed the cut trans so it doesn't give a delete prompt
'
'


Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Global Const SCR = 0
Global Const PTR = 1


Global Const MAX_DATA_TABLE = 10000
Global Const MAX_RECORDS_IN_MONTH = 200
Global MONTH_STRINGS(12) As String
Global LONG_MONTH_STRINGS(12) As String
Global DAY_STRINGS(7) As String
Global Const MAX_NOTES = 240  ' 20 years

' Paid constants
Global Const PAID_BLANK = 0
Global Const PAID_DONE = 1
Global Const PAID_QUESTION = 2
Global Const PAID_DASH = 3

Type version_block_type
  major_version As Integer
  minor_version As Integer
  file_format As Integer
  spare2 As Integer
  spare3 As Integer
  spare4 As Integer
  spare5 As Integer
  spare6 As Integer
End Type


' Calendar form structure
Public Type calendar_type
  Month As Integer
  day As Integer
  Year As Integer
  name As String
End Type

' These are the preference indexes for paste month
Public Const PREF_MONTH_STATUS0_INDEX = 0
Public Const PREF_MONTH_STATUS1_INDEX = 1
Public Const PREF_MONTH_STATUS2_INDEX = 2
Public Const PREF_MONTH_STATUS3_INDEX = 3
Public Const PREF_MONTH_CHECK_NUMBER_INDEX = 4
Public Const PREF_MONTH_CLEARED_INDEX = 5

Type preferences_type
  prompt_for_move As Boolean
  auto_insert As Boolean
  show_splash_screen As Boolean
  show_name_colors As Boolean
  show_override_columns As Boolean
  auto_check_done As Boolean
  prompt_for_delete As Boolean
  prompt_for_paste_notes As Boolean
  auto_negative_numbers As Boolean  ' True for autoinsert negative numbers
  auto_check_done_on_check As Boolean  ' True to auto check the done box when check number entered
  unlocked As Boolean  ' True when running a full version
  notes_font_size As Integer  ' Size of the notes
  language As Integer  ' 0=English, 1=Spanish
  paste_month(6) As Integer  ' Elements 0-3= Status(0=blank, 1=done, 2=pending, 3=skip), Element 4=check number (0=Same, 1=Blank), Element 5=Cleared column (0=same, 1=blank)
  save_recovery_file As Boolean  ' Save a duplicate copy of the file
  auto_load_last_file As Boolean  ' Reload the last file at startup if set to true
  play_sounds As Boolean  ' True is play sounds on Enter hit, false is quiet
End Type
  

' Declare the main data record type
Public Type r_type
  this As Integer
  previous As Integer
  next As Integer
  day As Byte
  Month As Byte
  Year As Integer
  name As String
  amount As Double
  balance As Double
  exclude As Boolean  ' false=no exclude, true=exclude
  override As Boolean
  override_amount As Double
  transaction_number As Integer
  sub_transaction_number As Integer
  paid As Integer  ' 0=blank, 1=done, 2=pending, 3=skip
  check As Integer  ' Check number, -1=no check number used
  tags As Byte
  cleared As Byte  ' 0=no, 1=yes, 2=permanent
  due As Byte  ' Day of the month this is due, 0=default
  pad(7) As Byte
End Type

' Declare the version 3 data record type
Type r3_type
  this As Integer
  previous As Integer
  next As Integer
  day As Byte
  Month As Byte
  Year As Integer
  name As String
  amount As Double
  balance As Double
  exclude As Boolean
  override As Boolean
  override_amount As Double
  transaction_number As Integer
  sub_transaction_number As Integer
  paid As Integer
  pad(12) As Byte
End Type

Type data_type
  db_name As String
  first As Integer
  last As Integer
  current As Integer
  number_of_records As Integer
  password As String
  number_of_notes As Integer
  bank_balance_beginning As Double
  bank_balance_ending As Double
  last_check_number As Integer
  pad(80) As Byte
End Type


Type view_type
  start_month As Byte
  start_year As Integer
  start_day As Integer
  current_day As Byte
  current_month As Byte
  current_year As Integer
  number_of_days As Integer
  entry_row As Integer
  last_balance As Double
  records_in_month As Integer
  quick_date_start As Integer
End Type

Type table_image_entry_type  ' Contains a copy of the dates and references for a single record
  day As Byte
  this As Integer
  next As Integer
  prev As Integer
End Type

Type table_image_type  ' Contains a copy of the dates and reference fields for the entire month
  table(MAX_RECORDS_IN_MONTH) As table_image_entry_type
  last As Integer
  Month As Integer
  Year As Integer
End Type



' ------------------ Define all the cardtrak stuff ----------------
Public Const MAX_CARD_TRANSACTION = 100  ' 100 transactions per month
Public Const MAX_CARD_INTEREST = 10  ' 10 interests for each month
Public Const MAX_CARD_TRANSACTIONS = 2000  ' Make space for 2000 months of card data
Public Const MAX_CARDS = 30  ' Allow for 20 credit cards
Public Const CT_EDIT = 0     ' Then we are editing an existing record
Public Const CT_CREATE = 1   ' Then we are making a blank record
Public Const CT_CONVERT = 2  ' Then we are converting a normal transaction to a ct transaction
Public Const CT_ADD = 3      ' Then we are adding new cards only

Type card_interest_type
  name As String
  balance As String
  percent As String
  payments As String
  charges As String
  interest As String
End Type

Type card_transaction_type
  active As Boolean
  card_number As Integer
  name As String
  spare_int As Integer  ' Record number in main database that this applies to
  card_this As Integer  ' Record number in the card transaction database that his applies to
  day As Integer
  Month As Integer
  Year As Integer
  
  previous_balance As Double
  new_balance As Double
  total_purchases As Double
  total_payments As Double
  total_interest As Double
  total_late As Double
  
  cleared As Integer
  status As Integer  ' Blank, done, pending, skip
  exclude As Boolean
  date_due As Integer
  date_paid As Integer
  amount_due As Double
  amount_paid As Double
  check_number As Integer
  tags As Byte
  
  transactions As String  ' Contains all the transactions for this record
  interest(MAX_CARD_INTEREST) As card_interest_type
  notes As String
End Type

Type card_info_type
  active As Boolean
  balance As Double
  created_day As Double  ' Month, day and year this card was created
  created_month As Double
  created_year As Double
  due_date As Integer  ' Day of the month this card payment is due
  interest(MAX_CARD_INTEREST) As card_interest_type
  name As String  ' Name of this card
  account As String  ' Account number
  notes As String  ' Misc information about this card
  s1 As String  ' Spare stuff
  s2 As String
  s3 As String
  s4 As String
  s5 As String
  spare_integer(5) As Integer
  spare_double(5) As Double
End Type
  
Type cardtrak_single_summary_type
  active As Boolean
  name As String
  balance As Double
  purchases As Double
  interest As Double
  paid As Double
  finance As Double
  late As Double
  minimum As Double
End Type

Public cardtrak_summary(MAX_CARDS) As cardtrak_single_summary_type
Public cardtrak_monthly_summary(12) As cardtrak_single_summary_type  ' 12 months
Public cardtrak_summary_single_month As cardtrak_single_summary_type  ' 1 month
Public cardtrak_filter As Integer  ' Filter on this card number, 0=no filter


' Undo for the entire month
Type copy_of_month_type
  Month As Integer
  Year As Integer
  table(MAX_RECORDS_IN_MONTH) As r_type  ' Enough room for all transactions for a month
  notes As String
End Type

Type month_notes_type
  Month As Integer
  Year As Integer
  s As String
End Type

Type cardtrak_month_type
  table() As card_transaction_type  ' This will be redim to MAX_RECORDS_IN_MONTH - Enough room for a cardtrak transaction for the month
End Type



' ------------ Balance information -----------
Type balance_summary_type
  beginning As Double
  ending As Double
  low As Double
  begin_found As Boolean
  end_found As Boolean
  tags(8) As Double
End Type
Public Const beginning = 0
Public Const ending = 1
Global balance_summary(12) As balance_summary_type  ' For the 12 displayed months


' ---------- Tag information ----------
Type tag_type
  total As Double
  pending As Double
  done As Double
  blank As Double
  skip As Double
  number As Integer
End Type
Public Const MAX_TAG = 3
Public tags(8) As tag_type  ' Define 8 tags
Global tag_mask(8)  ' Gets loaded on the tags_form with 1, 2, 4, 8...


' ---------- Summary information ----------
Public Const MAX_PAID = 3  ' 0=blank, 1=done, 2=question, 3=dash
Type summary_type
  income(MAX_PAID) As Double
  expense(MAX_PAID) As Double
  number_income(MAX_PAID) As Integer
  number_expense(MAX_PAID) As Integer
End Type
Global summary As summary_type  ' Keep track of month totals



' ------------------ Filter Parameters -------------------
Type filter_type
  active As Boolean
  name As String
  amount_from As Double
  amount_to As Double
  check As Integer
  status_ignore As Boolean
  status_blank As Boolean
  status_done As Boolean
  status_question As Boolean
  status_dash As Boolean
  filtered As Boolean  ' True when we have filtered transactions
  filtered_out_count As Integer  ' Count of the number of filtered out transactions
  filtered_in_count As Integer  ' Count of the number of filtered in transactions
  total_amount As Double  ' Amount of displayed transactions
  tags_ignore As Boolean  ' True means to ignore tag filtering
  tags(MAX_TAG) As Boolean   ' True means to filter this tag
End Type

Global filter As filter_type  ' Define the main filter variable





' --------------- Declare the undo stuff -------------
' -------- Define the types of actions that are possible for undo -----------
Global Const WHAT_NONE = 0
Global Const WHAT_PASTE_RECORD = 1        ' Ok
Global Const WHAT_CUT_RECORD = 2          ' Ok
Global Const WHAT_DELETE_RECORD = 3       ' Ok
Global Const WHAT_PASTE_MONTH = 4         ' Ok
Global Const WHAT_CUT_MONTH = 5           ' Ok
Global Const WHAT_MOVE_RECORD = 6         ' Ok
Global Const WHAT_CUT_TAGS = 7            '
Global Const WHAT_PASTE_TAGS = 8          '
Global Const WHAT_EDIT_TRANSACTION = 9    '
Global Const WHAT_MOVE_SELECTED = 10      '


Type undo_type
  doing_undo As Boolean
  what_was_done As Integer
  rec_num As Integer
  r As r_type  ' Record that was deleted
  copy_of_month As copy_of_month_type ' Contains of copy of the entire current month for the copy and paste month
  selected_rec_num(MAX_RECORDS_IN_MONTH) As Integer
  cardtrak As card_transaction_type  ' Contains the cardtrak if that's what it was
End Type





' ------------------ Declare all the Quick Accounts -------------
Public Const MAX_QUICK_ACCOUNT = 30
Type quick_account_type
  date As String
  name As String
  needed As Double
  total As Double
End Type

Public Type quick_accounts_type
  'account(MAX_QUICK_ACCOUNT + 1) As quick_account_type
  account() As quick_account_type
End Type

Global QUICK_ACCOUNTS As quick_accounts_type





' ------------------ Declare the variables ----------------

' Declare the undo stuff
Global undo As undo_type
Global undo_cardtrak_month As cardtrak_month_type

' Declare the file data
Global version_block As version_block_type
Global data As data_type
Global db(MAX_DATA_TABLE + 1) As r_type
Global db_temp(MAX_DATA_TABLE + 1) As r_type

' Declare the record type
Global this As r_type
Private r3_this As r3_type  ' Use this when reading in file format 3 type
Global copy_of_this As r_type   ' Used for cut and copy
Global copy_of_cardtrak As card_transaction_type

Global view As view_type  ' Contains the parameters for the current view in the table

Dim i, j, k As Integer
Global table_image As table_image_type  ' Contains of copy of the entire current month
Global copy_of_month As copy_of_month_type ' Contains of copy of the entire current month for the copy and paste month
Global copy_of_cardtrak_month As cardtrak_month_type

Global notes(1000) As month_notes_type

Global line_count As Integer   ' Used when printing
Global page_number As Integer   ' Used when printing
Global printer_error As Boolean  ' Used when printing

Global preferences As preferences_type
Global changed_flag As Boolean
Global cards_info(MAX_CARDS) As card_info_type
Global cards() As card_transaction_type









Public Sub strings_initialize()
  ' Set the version number
  version_block.major_version = major_version
  version_block.minor_version = minor_version
  version_block.file_format = file_format
  
  MONTH_STRINGS(0) = ""
  MONTH_STRINGS(1) = words(JAN_N)
  MONTH_STRINGS(2) = words(FEB_N)
  MONTH_STRINGS(3) = words(MAR_N)
  MONTH_STRINGS(4) = words(APR_N)
  MONTH_STRINGS(5) = words(MAY_N)
  MONTH_STRINGS(6) = words(JUN_N)
  MONTH_STRINGS(7) = words(JUL_N)
  MONTH_STRINGS(8) = words(AUG_N)
  MONTH_STRINGS(9) = words(SEP_N)
  MONTH_STRINGS(10) = words(OCT_N)
  MONTH_STRINGS(11) = words(NOV_N)
  MONTH_STRINGS(12) = words(DEC_N)
  
  DAY_STRINGS(1) = words(SUN_N)
  DAY_STRINGS(2) = words(MON_N)
  DAY_STRINGS(3) = words(TUE_N)
  DAY_STRINGS(4) = words(WED_N)
  DAY_STRINGS(5) = words(THU_N)
  DAY_STRINGS(6) = words(FRI_N)
  DAY_STRINGS(7) = words(SAT_N)
  
  LONG_MONTH_STRINGS(0) = ""
  LONG_MONTH_STRINGS(1) = words(JAN_LONG_N)
  LONG_MONTH_STRINGS(2) = words(FEB_LONG_N)
  LONG_MONTH_STRINGS(3) = words(MAR_LONG_N)
  LONG_MONTH_STRINGS(4) = words(APR_LONG_N)
  LONG_MONTH_STRINGS(5) = words(MAY_LONG_N)
  LONG_MONTH_STRINGS(6) = words(JUN_LONG_N)
  LONG_MONTH_STRINGS(7) = words(JUL_LONG_N)
  LONG_MONTH_STRINGS(8) = words(AUG_LONG_N)
  LONG_MONTH_STRINGS(9) = words(SEP_LONG_N)
  LONG_MONTH_STRINGS(10) = words(OCT_LONG_N)
  LONG_MONTH_STRINGS(11) = words(NOV_LONG_N)
  LONG_MONTH_STRINGS(12) = words(DEC_LONG_N)
End Sub


Sub get_record()
    this = db(data.current)
End Sub

Sub save_this()
  db(this.this) = this
End Sub


Function get_next_record() As Boolean
    With data
      If (db(.current).next = -1) Then
        ' We are already at the end of the file
        get_next_record = False
        Exit Function
      End If
      
      ' We have more records to go
      .current = db(.current).next
    End With
    
    this = db(data.current)
    get_next_record = True
End Function


Function get_previous_record() As Boolean
    With data
      If (db(.current).previous = -1) Then
        ' We are already at the beginning of the file
        get_previous_record = False
        Exit Function
      End If
      
      ' We have more records to go
      .current = db(.current).previous
    End With
    
    this = db(data.current)
    get_previous_record = True
End Function


Function find_first(ByVal m As Integer, ByVal y As Integer) As Boolean
  Dim found, done, d1, d2
  
  ' Get the first record that matches the month
  d2 = view.current_year * 12 + view.current_month
  With data
    find_first = False
  
    If (.number_of_records = 0) Then Exit Function  ' No records so we don't have anything to do
    
    ' Get the first record
    this = db(.first)
    view.last_balance = this.balance
    d1 = this.Year * 12 + this.Month
    'See if we are looking before the first record in the file
    If (d2 < d1) Then
      view.last_balance = 0
      Exit Function
    End If
    
    .current = .first
    found = False
    done = False
    While Not done
      If (d1 = d2) Then
           ' We have found the first record that matches
           found = True
           done = True
      End If
      If (Not found) Then If (Not get_next_record) Then done = True
      d1 = this.Year * 12 + this.Month
      If (d1 <= d2) Then view.last_balance = this.balance
      If (d1 > d2) Then done = True
    Wend
  End With
  
  If (found = True) Then
    find_first = True
  End If
  
End Function


Function find_next() As Boolean
  Dim m, y, found, done, d1, d2
  
  ' Get the next record that matches the month
  d2 = view.current_year * 12 + view.current_month
  With data
    find_next = False
    m = this.Month
    y = this.Year
    
    found = False
    done = Not get_next_record
    
    While Not done
      If (this.Month = m) And (this.Year = y) Then
           ' We have found the first record that matches
           found = True
           done = True
      End If
      If (Not found) Then If (Not get_next_record) Then done = True
      If (this.this = data.last) Then done = True
      
      ' See if we are past the current month and year
      d1 = this.Year * 12 + this.Month
      If (d1 > d2) Then
        done = True
      End If
      If (d1 <= d2) Then view.last_balance = this.balance
      
    Wend
  End With
  
  If (found = True) Then
    find_next = True
  End If
  
End Function


Sub delete_record(n As Integer)
  ' Delete the record
  With data
    ' First delete any associated cardtrak
    cards(db(n).sub_transaction_number).active = False
    
    db(n).this = -1
    
    If (db(n).next >= 0) Then
      ' We have a record following this one
      db(db(n).next).previous = db(n).previous
    Else
      ' We have no next record
      .last = db(n).previous
    End If
    
    If (db(n).previous >= 0) Then
      db(db(n).previous).next = db(n).next  ' We have a previous record
    Else
      ' We have no previous record
      .first = db(n).next
    End If
    
    .number_of_records = .number_of_records - 1
  End With
End Sub


Sub insert_record_after_this_one(ByRef prev_rec As Integer, ByRef new_rec As Integer)
  With data
    data.current = new_rec  ' Undo stuff
    db(new_rec).this = new_rec
    db(new_rec).next = db(prev_rec).next
    db(new_rec).previous = prev_rec
    db(prev_rec).next = new_rec
          
    ' Now adjust the next record
    If (db(new_rec).next >= 0) Then
      ' We have another record after the newly inserted one
      db(db(new_rec).next).previous = new_rec
    Else
      .last = new_rec
    End If
  
    ' Validate the next record
    If (db(new_rec).next >= 0) Then
      If (db(db(new_rec).next).this = -1) Then
        MsgBox "Pointing to null record"
        Stop
      End If
    End If
  End With
End Sub

Sub insert_record(ByVal prev_rec As Integer)
  ' Put THIS in the db
  ' prev_rec is the record number to insert after.  If it is -1 then insert at the first date
  ' Scan the record list and insert a new record
  ' Find the first available record
  ' Then put CT_THIS into the cards db if this.sub_transaction is > 0
  
  Dim new_rec As Integer
  Dim this_day, this_month, this_year, this_name, this_amount
  Dim first_time_through As Boolean
  
  With data
    For i = 0 To MAX_DATA_TABLE
      If (db(i).this = -1) Then
        ' We have found an empty slot   i = the new slot or record number
        new_rec = i
        
        If (.number_of_records = 0) Then
          ' This is the first record in the database
          db(0) = this
          db(0).this = 0
          db(0).previous = -1
          db(0).next = -1
          data.first = 0
          data.last = 0
          data.number_of_records = 1
          data.current = 0  ' Undo stuff
          this.this = new_rec  ' esk ?????
          Exit Sub
        End If
        
        .current = new_rec  ' Undo stuff
        
        db(new_rec) = this  ' Copy all the data to the new record
        this_day = this.day
        this_month = this.Month
        this_year = this.Year
        
        this_name = this.name
        this_amount = this.amount
        
        If (prev_rec >= 0) Then
          ' Insert the record after the record designated by prev_rec
          Call insert_record_after_this_one(prev_rec, new_rec)
        Else
          ' Find the place in the table where the new record goes
          ' See if it goes in the first record
          first_time_through = True
          .current = .first
          get_record
          If (this_year < this.Year) Or _
              ((this_year = this.Year) And (this_month < this.Month)) Or _
              ((this_year = this.Year) And (this_month = this.Month) And (this_day < this.day)) Then
            ' Insert the new record as the first record
            db(new_rec).this = new_rec
            db(new_rec).next = .first
            db(new_rec).previous = -1
            db(.first).previous = new_rec
            .first = new_rec
            .current = .first  ' Undo stuff
          Else
            ' The new record does not go before the first record so see where it belongs
            For j = 0 To .number_of_records - 1
              If (first_time_through) Then
                get_record
              Else
                k = get_next_record
              End If
            
              first_time_through = False
            
              If (this_year < this.Year) Or _
                  ((this_year = this.Year) And (this_month < this.Month)) Or _
                  ((this_year = this.Year) And (this_month = this.Month) And (this_day < this.day)) Then
              'If ((this_day < this.day) And (this_year = this.year)) Then
                ' We found a record with a date that is before the new record so insert the new record after this one
                prev_rec = this.previous
          
                Call insert_record_after_this_one(prev_rec, new_rec)
                Exit For
              End If
            
            Next j
          
            If (j = .number_of_records) Then
              ' We went the whole way through the database and didn't find where it goes so attach it to the end
              Call insert_record_after_this_one(.last, new_rec)
            End If
          End If
          
        End If
        .number_of_records = .number_of_records + 1
        Exit For
      End If
    Next i
    
    this.this = new_rec  ' esk ?????
      
    ' Put the cardtrak in now
    If (this.sub_transaction_number > 0) Then
    End If
    
    
    If (i = MAX_DATA_TABLE) Then
      ' We don't have any room
      MsgBox "No room in the data table"
    End If
  End With
End Sub

Sub insert_this_record(ByVal n As Integer)
  ' Insert THIS in the db
  ' Create a new record if n = -1
  ' Use this.this as the record number if n >= 0
  ' Put THIS in it
  ' Fix the pointers
  
  Dim r As r_type
  Dim p As Integer
  Dim t As Integer
  
  r = this  ' Save a temp copy
  If (n < 0) Then
    ' Create a new record and put this in it
    insert_record (-1)
    data.current = this.this
    'get_record
    'this = r
    'this.next = n
    'this.this = t
    'this.previous = p
  Else
    this = r
    db(this.this) = this
  End If
End Sub

Public Function insert_cardtrak_record(t As r_type, ct As card_transaction_type) As Integer
  ' Find a slot to put the incoming ct record
  ' Copy imcoming record to it
  ' Mark it as active
  ' Return the index number or zero if an error
  insert_cardtrak_record = 0
  For i = 1 To MAX_CARD_TRANSACTIONS
    ' Loop through all the card transactions and if they don't point to an active main then delete it
    If (Not cards(i).active) Then
      ' We found a slot
      cards(i) = ct  ' Copy the incoming ct record to it
      cards(i).active = True
      cards(i).card_this = i  ' This points to the ct db
      insert_cardtrak_record = i
      Exit For
    End If
  Next i
End Function


Public Sub save_to_month(what As Integer, t As r_type)
  Static j As Integer
  If (what = 0) Then
    j = 0
  End If
  If (what = 1) Then
    ' Add t to the month
    copy_of_month.table(j) = t
    ' Save the cardtrak also
    copy_of_cardtrak_month.table(j) = cards(t.sub_transaction_number)
    j = j + 1
  End If
End Sub


Public Sub clear_attributes(s As String)
  ' If the file exists then clear out any read only attributes and make it normal
  If (Dir(s) <> "") Then
    ' We have this file so now see if it's read only
    If ((GetAttr(s) And vbReadOnly) <> 0) Then
      ' We have a read only file so change it to normal
      SetAttr s, vbNormal
    End If
  End If
End Sub


' ------------------- write_database ----------------
Function write_database() As Boolean
  Dim s1, s2, s3
  Dim n As Integer
  
  main_form.MousePointer = vbHourglass
  
  ' Add up the number of notes
  data.number_of_notes = 0
  For i = 0 To MAX_NOTES
    If (notes(i).s <> "") Then data.number_of_notes = data.number_of_notes + 1
  Next i
  
  version_block.file_format = file_format
  version_block.major_version = major_version
  version_block.minor_version = minor_version
  
  s1 = data.db_name + ".czz"
  s2 = data.db_name + ".c2c"
  s3 = data.db_name + ".bak"
 
  On Error GoTo error_h
  
  clear_attributes (s1)
  clear_attributes (s2)
  clear_attributes (s3)
  
  
  If (Dir(s2) <> "") Then
    ' We have this file so now see if it's read only
    If ((GetAttr(s2) And vbReadOnly) <> 0) Then
      ' We have a read only file so change it to normal
      SetAttr s2, vbNormal
    End If
  End If
  
  If (Dir(s3) <> "") Then
    ' We have this file so now see if it's read only
    If ((GetAttr(s3) And vbReadOnly) <> 0) Then
      ' We have a read only file so change it to normal
      SetAttr s3, vbNormal
    End If
  End If
  
  ' Open the data file
  Open s1 For Binary As #1
  
  
  ' ------------- Write out the revision block ------------
  Put #1, , version_block
  
  ' ------------ Write out the header ------------
  Put #1, , data
  
  ' -------------- Write out the database --------------
  j = 0
  i = 0
  Do While j < data.number_of_records
    If (db(i).this >= 0) Then
      ' We have an active record here
      Put #1, , db(i)
      j = j + 1
    End If
    i = i + 1
  Loop
  
  
  ' ----------- Write out the notes -------------
  ' Write out the notes now
  For i = 0 To MAX_NOTES
    If (notes(i).s <> "") Then
      ' We have a valid notes file
      Put #1, , notes(i)
    End If
  Next i
  
  
  ' -------------- Write out the quick accounts ---------------
  If (version_block.file_format > 4) Then
    ' Write out the quick accounts
    Put #1, , MAX_QUICK_ACCOUNT
    For i = 0 To MAX_QUICK_ACCOUNT
      Put #1, , QUICK_ACCOUNTS.account(i)
    Next i
  End If
  
  
  ' ------------------ Write out the credit cards -----------------
  If (version_block.file_format > 5) Then
    ' Write out the credit card info
    Put #1, , MAX_CARDS
    For i = 0 To MAX_CARDS
      Put #1, , cards_info(i)
    Next i
  
    ' Add up the number of credit card transactions
    n = 0
    For i = 0 To MAX_CARD_TRANSACTIONS
      If (cards(i).active) Then n = n + 1
    Next i
    ' Write out the credit card transactions
    Put #1, , n
    For i = 0 To MAX_CARD_TRANSACTIONS
      If (cards(i).active) Then Put #1, , cards(i)
    Next i
  End If
  
  
  Close #1
  
  If (Error = "") Then
    If Dir(s3) <> "" Then Kill s3  ' Kill the backup file
    If Dir(s2) <> "" Then Name s2 As s3  ' Rename the original file as the backup
    If Dir(s3) <> "" Then Kill s3  ' Kill the backup file
    Name s1 As s2  ' Rename the new one
  Else
    MsgBox Error
  End If
  
  main_form.MousePointer = vbDefault
  Exit Function
  
error_h:
  ' We have an error so handle it
  MsgBox "Error writing to disk. " + Err.Description
  main_form.MousePointer = vbDefault
  Close #1
End Function

Function write_backup_database() As Boolean
  Dim s, s1
  
  s = data.db_name
  s1 = s + "_bak"
  If (Dir(s1) <> "") Then Kill s1  ' If the backupfile already exists then delete it
  data.db_name = s1
  write_database
  data.db_name = s
  
End Function

Function read_database() As Boolean
  Dim s1, s2, r, s_db
  Dim n As Integer
  Dim temp_card As card_transaction_type
  
  main_form.MousePointer = vbHourglass
  
  s1 = data.db_name + ".c2c"
  s_db = data.db_name
  
  ' Open the data file
  Open s1 For Binary Access Read As #1
  
  ' -------------- Read in the version block ----------------
  Get #1, , version_block
  
  If (version_block.file_format > file_format) Or (version_block.file_format < 3) Then
        r = MsgBox("Unsupported file format " + Format(version_block.file_format), vbInformation + vbOKOnly, "Information")
        Close #1
        read_database = False
        Exit Function
  End If
  
  ' -------------- Read in the header -------------
  Get #1, , data
  data.db_name = s_db
  
  ' ---------------- Check the password now ---------------
  If (data.password <> "") Then
    ' Put up the password box
    i = 0
    Do While (i < 4)
      i = i + 1
      If (password_form.execute(0)) Then
        If UCase(data.password) = UCase(password_form.password) Then
          i = 10
          Exit Do
          End If
      Else
        i = 0  ' Break out because cancel was hit
        Exit Do
      End If
      r = MsgBox("Invalid password", vbInformation + vbOKOnly, "Information")
    Loop
        
    If (i < 10) Then
      r = MsgBox("Invalid password", vbInformation + vbOKOnly, "Information")
      Close #1
      read_database = False
      data.number_of_notes = 0
      data.number_of_records = 0
      main_form.new_menu_Click
      Exit Function
    End If
  End If
  
  
  ' --------------- Zero out the array ----------------
  For i = 0 To MAX_DATA_TABLE
    db(i).this = -1
  Next i
  
  ' ---------------- Read in the database ---------------
  If (data.number_of_records > 0) Then
    For i = 1 To data.number_of_records
      If (version_block.file_format = 3) Then
        ' Read in file format 3
        Get #1, , r3_this
        Call copy_3_to_this(r3_this)
        If (this.paid = -1) Then this.paid = 1
        db(this.this) = this
      Else
        If (version_block.file_format <= 6) Then ' <<<<<<<<<<<<<
          ' Read in file format 4, 5, 6
          Get #1, , this
          If (this.paid = -1) Then this.paid = 1
          db(this.this) = this
        End If
      End If
    Next i
  End If
  
  ' -------------- Read in the notes now ---------------
  If (data.number_of_notes > 0) Then
    For i = 0 To data.number_of_notes - 1
      Get #1, , notes(i)
    Next i
  End If
  
  
  ' -------------- Read in the quick accounts ----------------
  ReDim QUICK_ACCOUNTS.account(MAX_QUICK_ACCOUNT + 1)
  If (version_block.file_format > 4) Then
    ' Read in the quick accounts
    Get #1, , n  ' Read in the number of quick accounts
    ReDim QUICK_ACCOUNTS.account(n + 1)
    For i = 0 To n
      Get #1, , QUICK_ACCOUNTS.account(i)
    Next i
  End If
  
  ' -------------- Read in the credit card info --------------
  If (version_block.file_format > 5) Then
    ' Read in the credit card info
    Get #1, , n  ' Read in the number of cards
    For i = 0 To n
      Get #1, , cards_info(i)
    Next i
  
    ' Read in the credit card transactions
    ' Zero out the card transactions
    For i = 0 To MAX_CARD_TRANSACTIONS
      cards(i).active = False
    Next i
    Get #1, , n  ' Read in the number of card transactions
    If (n > 0) Then
      For i = 0 To n
        Get #1, , temp_card
        cards(temp_card.card_this) = temp_card
      Next i
    End If
  End If
  
  
  ' -------------- Close the file -------------
  Close #1
  ReDim copy_of_cardtrak_month.table(MAX_RECORDS_IN_MONTH)
  ReDim undo_cardtrak_month.table(MAX_RECORDS_IN_MONTH)
  
  main_form.MousePointer = vbDefault
  read_database = True
  
  
  ' ----- Create a rebuilt database -----
  ' ESK 1/26/2017
  ' I don't remember why the following section was added whic rebuilds the database
  ' and validates each record and deletes duplicate records. This section was not in the stable
  ' version of 3.4011 so I'm taking it out by putting in the Exit Function. This is essentially
  ' doing what 3.4011 did.
  Exit Function
  
  
  ' ----- Create a rebuilt database -----
  ' Copy the existing database records to a temp value
  ' and zero the records in the existing database
  For i = 0 To MAX_DATA_TABLE
    db_temp(i) = db(i)
    db(i).this = -1
  Next i
  
  ' Create a database all empty records
  data.first = 0
  data.last = 0
  data.number_of_records = 0
  
  ' Go through the entire temp database and validate each record and add each record if it is valid
Dim last_day
Dim last_month
Dim last_year
Dim last_name
Dim last_amount
Dim j

  For i = 0 To MAX_DATA_TABLE
    If ((db_temp(i).this >= 0) And (db_temp(i).day > 0)) Then
        ' We have a valid record so add it
        this = db_temp(i)
           
        insert_this_record (-1)
    
    End If
  Next i
  
  ' Loop through the database and delete duplicate records
  data.current = data.first
  get_record
  For i = 1 To MAX_DATA_TABLE
        
        j = get_next_record()
        If (j = False) Then Exit For
        
        If (this.day = last_day) And _
           (this.Month = last_month) And _
           (this.Year = last_year) And _
           (this.name = last_name) And _
           (this.amount = last_amount) Then
            
            '   ESK 1/25/2017
            ' Commented out the following line because a valid database is allowed
            ' to have duplicate records. I don't remember why I thought it was important
            ' to remove duplicate records unless somehow they were in error.
            ' Also, I don't remember why this whole section on creating a rebuilt database was
            ' necessary unless it was to organize the records differently such as by date.
            ' The previous and stable version that I have been using for many years was
            ' version "3.4011 5/10/2004"
            ' This new version with the next line commented out seems to produce the same
            ' totals from a large database on 2/25/2017. My plan is to recompile this version
            ' and test it against 3.4011.
            ' ESK 1/26/2017
            ' I uncommented the line below because I inserted the Exit Function which
            ' prevents this entire section to not execute
            delete_record (data.current)
            
        End If
    
        last_day = this.day
        last_month = this.Month
        last_year = this.Year
        last_name = this.name
        last_amount = this.amount
  Next i
  
  
  
End Function


Private Sub copy_3_to_this(ByRef r3_this As r3_type)
  ' Copy the r3_this record to this record
  this.this = r3_this.this
  this.previous = r3_this.previous
  this.next = r3_this.next
  this.day = r3_this.day
  this.Month = r3_this.Month
  this.Year = r3_this.Year
  this.name = r3_this.name
  this.amount = r3_this.amount
  this.balance = r3_this.balance
  this.exclude = r3_this.exclude
  this.override = r3_this.override
  this.override_amount = r3_this.override_amount
  this.transaction_number = r3_this.transaction_number
  this.sub_transaction_number = r3_this.sub_transaction_number
  this.paid = r3_this.paid
  this.check = -1
End Sub



Sub MSFlexGridEdit(MSFlexGrid As Control, Edt As Control, KeyAscii As Integer)

    ' Use the character that was typed.
    Select Case KeyAscii

    ' A space means edit the current text.
    Case 0 To 32
        Edt = MSFlexGrid
        Edt.SelStart = 1000

    ' Anything else means replace the current text.
    Case Else
        Edt = Chr(KeyAscii)
        Edt.SelStart = 1
    End Select

    ' Show Edt at the right place.
    Edt.Move MSFlexGrid.CellLeft + MSFlexGrid.Left, MSFlexGrid.CellTop + MSFlexGrid.Top, _
    MSFlexGrid.CellWidth, MSFlexGrid.CellHeight
    Edt.Visible = True
    Edt.CausesValidation = True

    ' And let it work.
    Edt.SetFocus
End Sub

Sub EditKeyCode(MSFlexGrid As Control, Edt As _
Control, KeyCode As Integer, Shift As Integer)

    ' Standard edit control processing.
    Select Case KeyCode

    Case 27 ' ESC: hide, return focus to MSFlexGrid.
        Edt.Visible = False
        MSFlexGrid.SetFocus
        
    Case 13 ' ENTER return focus to MSFlexGrid.
        MSFlexGrid.SetFocus
        play_sound (2)

    Case 38     ' Up.
        MSFlexGrid.SetFocus
        DoEvents
        If MSFlexGrid.row > MSFlexGrid.FixedRows Then
            MSFlexGrid.row = MSFlexGrid.row - 1
        End If

    Case 40     ' Down.
        MSFlexGrid.SetFocus
        DoEvents
        If MSFlexGrid.row < MSFlexGrid.Rows - 1 Then
            MSFlexGrid.row = MSFlexGrid.row + 1
        End If
    End Select
End Sub
    

Function GetCommandLine(Optional MaxArgs)
   'Declare variables.
   Dim c, CmdLine, CmdLnLen, InArg, i, NumArgs
   'See if MaxArgs was provided.
   If IsMissing(MaxArgs) Then MaxArgs = 10
   'Make array of the correct size.
   ReDim ArgArray(MaxArgs)
   NumArgs = 0: InArg = False
   'Get command line arguments.
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   'Go thru command line one character
   'at a time.
   For i = 1 To CmdLnLen
      c = Mid(CmdLine, i, 1)
      'Test for space or tab.
      If (c <> " " And c <> vbTab) Then
         'Neither space nor tab.
         'Test if already in argument.
         If Not InArg Then
         'New argument begins.
         'Test for too many arguments.
            If NumArgs = MaxArgs Then Exit For
            NumArgs = NumArgs + 1
            InArg = True
         End If
         'Concatenate character to current argument.
         ArgArray(NumArgs) = ArgArray(NumArgs) & c
      Else
         'Found a space or tab.
         'Set InArg flag to False.
         InArg = False
      End If
   Next i
   'Resize array just enough to hold arguments.
   ReDim Preserve ArgArray(NumArgs)
   'Return Array in Function name.
   GetCommandLine = ArgArray()
End Function


Function strip_filename(ByRef fn As Variant) As String
  strip_filename = fn
  For i = Len(fn) To 1 Step -1
    If (Mid(fn, i, 1) = "\") Then
      strip_filename = Mid(fn, 1, i)
      Exit For
    End If
  Next i
  
End Function



Function get_drive(ByRef fn As Variant) As String
    For i = 1 To Len(fn)
      get_drive = get_drive + Mid(fn, i, 1)
      If (Mid(fn, i, 1) = ":") Then
      Exit For
      End If
    Next i
  
End Function


' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' ---------------------- sub main ------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' This is the start of program execution
' --------------------------------------------------------
'
Public Sub Main()
  Dim s As Long
  'main_form.register_menu.Visible = False
  'main_form.register_menu_dash.Visible = False
  
  ReDim cards(MAX_CARD_TRANSACTIONS)
  ReDim copy_of_cardtrak_month.table(MAX_RECORDS_IN_MONTH)
  ReDim undo_cardtrak_month.table(MAX_RECORDS_IN_MONTH)
  
  If (Not register_form.ok_to_run) Then
    main_form.EvaluationExpired  ' Disable the controls since evaluation period expired
  End If
  
  main_form.show
  
  s = Len(cards(0))
End Sub



Public Function currency_s(ByVal v As Double) As String
  currency_s = Format(v, "###,###,##0.00 ")
End Function


Public Function amount_color(amt As Double) As Long
  If (amt < -0.0001) Then
    ' Change the color of the cell to red
    amount_color = vbRed
  Else
    If (amt > 0.0001) Then
      ' Change the color of the cell to blue
      amount_color = vbBlue
    Else
      amount_color = vbBlack
    End If
  End If
End Function


Public Function filter_check() As Boolean
  Dim f As Boolean, ff As Boolean
  Dim show_it As Boolean
  Dim show As Boolean
  
  ' Return true if the transaction THIS is displayable
  ' Return false if it is filtered out
  ' Return true if filter is not active
  If (filter.active = False) Then
    filter_check = True
    Exit Function
  End If
  
  
  ' ----- Check the name -----
  show_it = True
  If (filter.name <> "") Then
    If (InStr(UCase(this.name), UCase(filter.name)) <= 0) Then show_it = False
  End If
  
  ' ----- Check the status -----
  If (filter.status_ignore = False) Then
    ' Ignore is not checked so check the rest of them
    show = False
    If (this.paid = 0) And (filter.status_blank) Then show = True
    If (this.paid = 1) And (filter.status_done) Then show = True
    If (this.paid = 2) And (filter.status_question) Then show = True
    If (this.paid = 3) And (filter.status_dash) Then show = True
    If (show = False) Then show_it = False
  End If
    
  ' ----- Check the tags -----
  If (filter.tags_ignore = False) Then
    ' Ignore is not checked so check the rest of them
    show = False
    If (((this.tags And tag_mask(0)) <> 0) And (filter.tags(0))) Then show = True
    If (((this.tags And tag_mask(1)) <> 0) And (filter.tags(1))) Then show = True
    If (((this.tags And tag_mask(2)) <> 0) And (filter.tags(2))) Then show = True
    If (((this.tags And tag_mask(3)) <> 0) And (filter.tags(3))) Then show = True
    If (show = False) Then show_it = False
  End If
    
  ' ----- Check the check number -----
  If (filter.check > -1) Then
    If (this.check <> filter.check) Then show_it = False
  End If
  
  ' ----- Check the amount from and to -----
  If ((filter.amount_from <> -55) And (filter.amount_to) <> -55) Then
    ' We have something to check
    If (this.amount < filter.amount_from) Or (this.amount > filter.amount_to) Then show_it = False
  End If
  
  ' ----- See if it's exclude -----
  'If (this.exclude) Then show_it = False
  
  
  ' ---------- Now see what we have to show ----------
  If (show_it = False) Then
    filter.filtered = True
    filter.filtered_out_count = filter.filtered_out_count + 1
    filter_check = False
  Else
    filter.filtered_in_count = filter.filtered_in_count + 1
    filter_check = True
  End If
End Function


Public Function quick_date_this() As Integer
  quick_date_this = this.Year * 12 + this.Month
End Function

Public Function quick_date_this_day() As Integer
  quick_date_this_day = (this.Year - 2000) * 31 * 12 + this.Month * 31 + this.day
End Function

Public Function quick_index() As Integer
  quick_index = quick_date_this - view.quick_date_start
End Function

Public Sub put_this_in_balance_summary()
  Dim start As Integer
  Dim ending As Integer
  Dim n As Integer
  Dim qi As Integer
  
  qi = quick_index
  
  ' See if the current data matches where we are
  If (quick_index >= 0) And (quick_index <= 11) And (this.exclude = False) Then
    ' We have a tab that it this record can go on
    ' See if it's the first record in the list
    If (balance_summary(qi).begin_found = False) Then
      ' We have a new tab to put it on
      balance_summary(qi).beginning = this.balance - this.amount
      If (this.override) Then balance_summary(qi).beginning = this.override_amount
      balance_summary(qi).begin_found = True
    End If
    
    balance_summary(qi).ending = this.balance
    If (this.balance < balance_summary(qi).low) Then balance_summary(qi).low = this.balance
  End If
  
  If (quick_index >= 0) And (quick_index <= 10) Then
    For i = quick_index + 1 To 11
    balance_summary(i).beginning = balance_summary(i - 1).ending
    balance_summary(i).ending = balance_summary(i - 1).ending
    balance_summary(i).low = balance_summary(i - 1).ending
    Next i
  End If

End Sub


Public Function strip_spaces(s As String) As String
  strip_spaces = ""
  For i = 1 To Len(s)
    If (Mid(s, i, 1) <> " ") Then strip_spaces = strip_spaces + Mid(s, i, 1)
  Next i
End Function


Public Function save_axgrid_into_string(c As axgrid) As String
  Dim i, j, m, n
  Dim s As String
  
  ' Save the contents of the grid and return the string of it
  
  ' Find the last row and column used
  For i = c.Rows To 1 Step -1
    For j = c.Cols To 1 Step -1
      If (c.TextMatrix(i, j) <> "") Then Exit For
    Next j
    If (c.TextMatrix(i, j) <> "") And (j > 0) Then Exit For
  Next i
  
  ' Now save the cells up to the last cell
  For m = 1 To i
    For n = 1 To c.Cols
    s = s + Chr(6) + c.TextMatrix(m, n)
    If (m = i) And (n = j) Then Exit For
    Next n
  Next m
  
  save_axgrid_into_string = s
End Function




Public Sub load_axgrid_from_string(s As String, c As axgrid)
  Dim i, j, m, n
  Dim done
  
  ' Save the contents of the grid and return the string of it
  
  ' Find the last row and column used
  m = 1
  done = False
  For i = 1 To c.Rows
    For j = 1 To c.Cols
      If (done) Then
        c.TextMatrix(i, j) = ""
      Else
        ' Find the next number
        m = InStr(m, s, Chr(6)) + 1
        If (m = 1) Then done = True  'Exit For
        n = InStr(m, s, Chr(6))
        If (n = 0) Then n = Len(s) + 1
        c.TextMatrix(i, j) = Mid(s, m, n - m)
      End If
    Next j
  Next i
End Sub


Public Function GetExtension(Filename As String) As String

Dim intPathPos As Integer, intExtPos As Integer

  For i = Len(Filename) To 1 Step -1
    'Cycle through filename, character by character

    If Mid(Filename, i, 1) = "." Then
      'If current character is a period then
      intExtPos = i
  
      'Alter variable to reflect position

      For j = Len(Filename) To 1 Step -1
        If Mid(Filename, j, 1) = "\" Then
           intPathPos = j
           Exit For
        End If
      Next j

      'Then find the last slash in the string
      '- we'll perform basic filename checking
      'later. We'll ensure this slash is *before*
      'the last period, otherwise there's no valid
      'extension
      Exit For
    End If
  Next i

  If intPathPos > intExtPos Then
     Exit Function
  Else
     If intExtPos = 0 Then Exit Function
     GetExtension = Mid(Filename, intExtPos + 1, Len(Filename) - intExtPos)
  End If

  'Finally, extract the characters after that final period

End Function



Public Sub set_center_tab(Month As Integer, Year As Integer)
' Set the center tab to reflect the date passed to this function
    
    Month = Month - 5  ' Subtract off 5 so that we can set the active tab to the center
    If (Month <= 0) Then
      ' We got here because when we subtracted 5 from the month turned it to 0 or negative
      ' which means we are in the previous year so fix the month and decrement the year
      Month = Month + 12
      Year = Year - 1
    End If
    
    view.start_month = Month
    view.start_year = Year
    
    main_form.entry_tab.Tab = 5
    main_form.update_entry_tabs

End Sub
